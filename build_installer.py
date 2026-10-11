import ctypes as c
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import uuid
from pathlib import Path

msi = c.WinDLL('msi')
U = c.c_uint
P = c.POINTER(U)
for name, args in {
    'MsiOpenDatabaseW': [c.c_wchar_p, c.c_void_p, P],
    'MsiDatabaseOpenViewW': [U, c.c_wchar_p, P],
    'MsiViewExecute': [U, U],
    'MsiViewFetch': [U, P],
    'MsiViewGetColumnInfo': [U, c.c_int, P],
    'MsiRecordGetStringW': [U, U, c.c_wchar_p, P],
    'MsiRecordGetFieldCount': [U],
    'MsiCloseHandle': [U],
    'MsiCreateRecord': [U],
    'MsiRecordSetStringW': [U, U, c.c_wchar_p],
    'MsiRecordSetInteger': [U, U, c.c_int],
    'MsiRecordSetStreamW': [U, U, c.c_wchar_p],
    'MsiRecordReadStream': [U, U, c.c_void_p, P],
    'MsiDatabaseCommit': [U],
    'MsiDatabaseImportW': [U, c.c_wchar_p, c.c_wchar_p],
    'MsiGetSummaryInformationW': [U, c.c_wchar_p, U, P],
    'MsiSummaryInfoSetPropertyW': [U, U, U, c.c_int, c.c_void_p, c.c_wchar_p],
    'MsiSummaryInfoPersist': [U],
}.items():
    getattr(msi, name).argtypes = args
    getattr(msi, name).restype = U

def check(result):
    if result:
        raise RuntimeError(f'Windows Installer error {result}')

def string(record, field):
    size = U(0)
    status = msi.MsiRecordGetStringW(record, field, None, c.byref(size))
    if status not in (0, 234):
        check(status)
    size.value += 1
    buf = c.create_unicode_buffer(size.value)
    check(msi.MsiRecordGetStringW(record, field, buf, c.byref(size)))
    return buf.value

def query(database, sql):
    view = U()
    check(msi.MsiDatabaseOpenViewW(database, sql, c.byref(view)))
    try:
        check(msi.MsiViewExecute(view, 0))
        columns = U()
        check(msi.MsiViewGetColumnInfo(view, 0, c.byref(columns)))
        names = [string(columns, i) for i in range(1, msi.MsiRecordGetFieldCount(columns) + 1)]
        msi.MsiCloseHandle(columns)
        rows = []
        while True:
            record = U()
            result = msi.MsiViewFetch(view, c.byref(record))
            if result == 259:
                break
            check(result)
            rows.append(dict(zip(names, (string(record, i) for i in range(1, len(names) + 1)))))
            msi.MsiCloseHandle(record)
        return rows
    finally:
        msi.MsiCloseHandle(view)

def execute(database, sql, *values, stream_field=None):
    view, record = U(), msi.MsiCreateRecord(len(values))
    check(msi.MsiDatabaseOpenViewW(database, sql, c.byref(view)))
    try:
        for i, value in enumerate(values, 1):
            if i == stream_field:
                check(msi.MsiRecordSetStreamW(record, i, str(value)))
            elif isinstance(value, int):
                check(msi.MsiRecordSetInteger(record, i, value))
            else:
                check(msi.MsiRecordSetStringW(record, i, str(value)))
        check(msi.MsiViewExecute(view, record))
    finally:
        msi.MsiCloseHandle(record)
        msi.MsiCloseHandle(view)


def export_cabinet(database, destination):
    media = query(database, 'SELECT * FROM `Media`')
    assert len(media) == 1 and media[0]['Cabinet'].startswith('#'), 'Expected one embedded cabinet'
    name = media[0]['Cabinet'][1:]
    view, record = U(), U()
    check(msi.MsiDatabaseOpenViewW(database, f"SELECT `Data` FROM `_Streams` WHERE `Name` = '{name}'", c.byref(view)))
    try:
        check(msi.MsiViewExecute(view, 0))
        check(msi.MsiViewFetch(view, c.byref(record)))
        buffer = c.create_string_buffer(65536)
        with destination.open('wb') as target:
            while True:
                count = U(len(buffer))
                check(msi.MsiRecordReadStream(record, 1, buffer, c.byref(count)))
                if count.value == 0:
                    break
                target.write(buffer.raw[:count.value])
    finally:
        msi.MsiCloseHandle(record)
        msi.MsiCloseHandle(view)
    return name


def assembly_metadata(build):
    # Reflection reads assembly identities without loading application code or registering COM.
    script = r'''
$metadata = @{}
Get-ChildItem -LiteralPath $env:CLAWPDF_BUILD_FOLDER -File | Where-Object { $_.Extension -in '.dll','.exe' } | ForEach-Object {
    try {
        $identity = [Reflection.AssemblyName]::GetAssemblyName($_.FullName)
        $metadata[$_.Name.ToLowerInvariant()] = @{
            Name=$identity.Name; Version=$identity.Version.ToString()
            Culture=if($identity.CultureName){$identity.CultureName}else{'neutral'}
            PublicKeyToken=[BitConverter]::ToString($identity.GetPublicKeyToken()).Replace('-','').ToUpperInvariant()
            ProcessorArchitecture=$identity.ProcessorArchitecture.ToString()
            FileVersion=$_.VersionInfo.FileVersionRaw.ToString()
        }
    } catch {}
}
$metadata | ConvertTo-Json -Depth 4
'''
    env = dict(os.environ, CLAWPDF_BUILD_FOLDER=str(build))
    result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script],
                            env=env, check=True, capture_output=True, text=True)
    return json.loads(result.stdout)


def run_tool(args, folder):
    result = subprocess.run(args, cwd=folder, capture_output=True)
    (folder / (Path(args[0]).stem + '.log')).write_bytes(result.stdout + result.stderr)
    if result.returncode:
        raise RuntimeError(f'{args[0]} failed; see its log in {folder}')


def main():
    parser = argparse.ArgumentParser(description='Build the Word SVG MSI from the upstream installer and Release binaries.')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--build-dir', type=Path, help='Release directory; defaults to src/_Build/Release.')
    args = parser.parse_args()
    repo = Path(__file__).resolve().parent
    build = args.build_dir.resolve() if args.build_dir else repo / 'src/_Build/Release'
    metadata = assembly_metadata(build)
    version = '.'.join(metadata['clawpdf.exe']['Version'].split('.')[:3])
    if version != '0.9.6':
        raise RuntimeError('Build the 0.9.6 Release binaries first.')
    if (build / 'clawPDF.tlb').read_bytes()[:4] != b'MSFT':
        raise RuntimeError('Export clawPDF.tlb with TlbExp.exe before building the installer.')
    work = repo / '.build-tools' / ('installer-' + uuid.uuid4().hex)
    stage, cabinets = work / 'files', work / 'cabinet'
    stage.mkdir(parents=True)
    cabinets.mkdir()
    candidate = work / 'candidate.msi'
    shutil.copyfile(repo / 'src/_MSI/clawPDF_0.9.3_setup.msi', candidate)
    database = U()
    check(msi.MsiOpenDatabaseW(str(candidate), c.c_void_p(1), c.byref(database)))
    try:
        # Set the MSI string pool encoding before writing Chinese UI text.
        (work / '_ForceCodepage.idt').write_bytes(b'\r\n\r\n936\t_ForceCodepage\r\n')
        check(msi.MsiDatabaseImportW(database, str(work), '_ForceCodepage.idt'))
        translations = json.loads((repo / 'installer_zh_cn.json').read_text(encoding='utf-8'))
        for table, keys, field in (('Control', ('Dialog_', 'Control'), 'Text'),
                                   ('RadioButton', ('Property', 'Order'), 'Text'),
                                   ('Dialog', ('Dialog',), 'Title')):
            for row in query(database, f'SELECT * FROM `{table}`'):
                match = re.match(r'^(\{\\[^}]+\})?(.*)$', row[field], re.DOTALL)
                prefix, text = match.group(1) or '', match.group(2)
                if text in translations:
                    where = ' AND '.join(f'`{key}`=?' for key in keys)
                    execute(database, f'UPDATE `{table}` SET `{field}`=? WHERE {where}',
                            prefix + translations[text], *(int(row[key]) if key == 'Order' else row[key] for key in keys))
        execute(database, 'UPDATE `TextStyle` SET `FaceName`=?', 'Microsoft YaHei UI')
        cab_name = export_cabinet(database, work / 'original.cab')
        run_tool(['expand.exe', '-F:*', str(work / 'original.cab'), str(stage)], work)
        files = query(database, 'SELECT * FROM `File`')
        original_count = len(files)
        components = {row['Component']: row for row in query(database, 'SELECT * FROM `Component`')}
        directories = {row['Directory']: row for row in query(database, 'SELECT * FROM `Directory`')}

        def directory_path(key):
            if key == 'TARGETDIR':
                return Path()
            row = directories[key]
            name = row['DefaultDir'].split(':')[0].split('|')[-1]
            return directory_path(row['Directory_Parent']) / ('' if name == '.' else name)

        destinations = {}
        for row in files:
            relative = directory_path(components[row['Component_']]['Directory_']) / row['FileName'].split('|')[-1]
            destinations[row['File']] = relative
            replacement = build / relative
            if replacement.is_file():
                shutil.copyfile(replacement, stage / row['File'])

        # Include runtime files newly generated by the build, including the bridge configuration.
        existing = {str(path).lower() for path in destinations.values()}
        for extra in build.iterdir():
            if not extra.is_file() or extra.name.startswith('WordPdfRegression') or extra.name.lower() in existing:
                continue
            if extra.suffix.lower() not in ('.dll', '.exe', '.config'):
                continue
            key = 'F_Word_' + hashlib.sha256(extra.name.encode()).hexdigest()[:24]
            component = 'C_' + key
            guid = '{' + str(uuid.uuid5(uuid.NAMESPACE_URL, 'https://github.com/sch401/clawPDF/' + extra.name)).upper() + '}'
            sequence = len(files) + 1
            execute(database, 'INSERT INTO `Component` (`Component`,`ComponentId`,`Directory_`,`Attributes`,`KeyPath`) VALUES (?,?,?,?,?)',
                    component, guid, 'TARGETDIR', 0, key)
            execute(database, 'INSERT INTO `FeatureComponents` (`Feature_`,`Component_`) VALUES (?,?)', 'DefaultFeature', component)
            execute(database, 'INSERT INTO `File` (`File`,`Component_`,`FileName`,`FileSize`,`Attributes`,`Sequence`) VALUES (?,?,?,?,?,?)',
                    key, component, extra.name, extra.stat().st_size, 512, sequence)
            files.append(dict(File=key, Component_=component, FileName=extra.name, Version='', Sequence=str(sequence)))
            destinations[key] = Path(extra.name)
            shutil.copyfile(extra, stage / key)

        for row in files:
            payload = stage / row['File']
            if not payload.is_file():
                raise RuntimeError(f'Missing cabinet file {row["File"]}')
            identity = metadata.get(destinations[row['File']].name.lower())
            file_version = identity['FileVersion'] if identity else row['Version']
            execute(database, 'UPDATE `File` SET `FileSize`=?, `Version`=? WHERE `File`=?', payload.stat().st_size, file_version, row['File'])
            execute(database, 'DELETE FROM `MsiFileHash` WHERE `File_`=?', row['File'])
            if not file_version:
                # MsiFileHash is the four signed DWORDs of the file's MD5 digest.
                import struct
                parts = struct.unpack('<4i', hashlib.md5(payload.read_bytes()).digest())
                execute(database, 'INSERT INTO `MsiFileHash` (`File_`,`Options`,`HashPart1`,`HashPart2`,`HashPart3`,`HashPart4`) VALUES (?,?,?,?,?,?)',
                        row['File'], 0, *parts)
            if identity:
                for entry in query(database, f"SELECT * FROM `MsiAssemblyName` WHERE `Component_` = '{row['Component_']}'"):
                    if entry['Name'] not in identity:
                        continue
                    value = identity[entry['Name']]
                    if value:
                        execute(database, 'UPDATE `MsiAssemblyName` SET `Value`=? WHERE `Component_`=? AND `Name`=?', value, row['Component_'], entry['Name'])
                    else:
                        execute(database, 'DELETE FROM `MsiAssemblyName` WHERE `Component_`=? AND `Name`=?', row['Component_'], entry['Name'])

        for name, value in {'ProductVersion': version, 'ProductName': 'clawPDF ' + version + ' Word SVG 中文版',
                            'ProductLanguage': '2052',
                            'ProductCode': '{' + str(uuid.uuid5(uuid.NAMESPACE_URL, 'https://github.com/sch401/clawPDF/' + version)).upper() + '}',
                            'ARPURLINFOABOUT': 'https://github.com/sch401/clawPDF/tree/preserve-word-svg'}.items():
            execute(database, 'UPDATE `Property` SET `Value`=? WHERE `Property`=?', value, name)
        upgrades = query(database, 'SELECT * FROM `Upgrade`')
        execute(database, 'DELETE FROM `Upgrade`')
        for upgrade in upgrades:
            execute(database, 'INSERT INTO `Upgrade` (`UpgradeCode`,`VersionMin`,`VersionMax`,`Language`,`Attributes`,`Remove`,`ActionProperty`) VALUES (?,?,?,?,?,?,?)',
                    upgrade['UpgradeCode'], upgrade['VersionMin'], version, upgrade['Language'],
                    int(upgrade['Attributes']), upgrade['Remove'], upgrade['ActionProperty'])
        # Replace the old product inside the transaction before staging new files.
        # This also replaces equal-version libraries rebuilt without the upstream signing key.
        execute(database, 'UPDATE `InstallExecuteSequence` SET `Sequence`=? WHERE `Action`=?', 1520, 'RemoveExistingProducts')
        execute(database, 'UPDATE `Media` SET `LastSequence`=? WHERE `DiskId`=?', len(files), 1)
        summary = U()
        check(msi.MsiGetSummaryInformationW(database, None, 3, c.byref(summary)))
        check(msi.MsiSummaryInfoSetPropertyW(summary, 1, 2, 936, None, None))
        check(msi.MsiSummaryInfoSetPropertyW(summary, 7, 30, 0, None, 'Intel;2052'))
        check(msi.MsiSummaryInfoSetPropertyW(summary, 9, 30, 0, None, '{' + str(uuid.uuid4()).upper() + '}'))
        check(msi.MsiSummaryInfoPersist(summary))
        msi.MsiCloseHandle(summary)

        ddf = work / 'payload.ddf'
        directives = ['.OPTION EXPLICIT', '.Set CabinetNameTemplate=payload.cab',
                      f'.Set DiskDirectoryTemplate={cabinets}', '.Set CompressionType=LZX',
                      '.Set CompressionMemory=21', '.Set Cabinet=ON', '.Set Compress=ON',
                      '.Set MaxDiskSize=0', '.Set MaxCabinetSize=0']
        directives += [f'"{stage / row["File"]}" {row["File"]}' for row in sorted(files, key=lambda row: int(row['Sequence']))]
        ddf.write_text('\n'.join(directives), encoding='mbcs')
        run_tool(['makecab.exe', '/F', str(ddf)], work)
        if len(list(cabinets.glob('*.cab'))) != 1:
            raise RuntimeError('Expected a single cabinet.')
        execute(database, 'UPDATE `_Streams` SET `Data`=? WHERE `Name`=?', cabinets / 'payload.cab', cab_name, stream_field=1)
        check(msi.MsiDatabaseCommit(database))
        (work / 'payload-map.json').write_text(json.dumps({key: str(value) for key, value in destinations.items()}, indent=2), encoding='utf-8')
    finally:
        msi.MsiCloseHandle(database)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(candidate, args.output)
    print(json.dumps(dict(output=str(args.output.resolve()), version=version, files=len(files),
                          added_files=len(files) - original_count, work=str(work)), indent=2))


if __name__ == '__main__':
    main()
