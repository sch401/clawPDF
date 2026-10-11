import hashlib
import json
import shutil
import struct
import sys
import uuid
from pathlib import Path

repo = Path(sys.argv[1]).resolve()
package = Path(sys.argv[2]).resolve()
sys.path.insert(0, str(repo))
import build_installer as api

folder = repo.parent / ('installer-validation-' + uuid.uuid4().hex)
folder.mkdir()
layout = folder / 'application'
layout.mkdir()
databases = []
try:
    for path in (repo / 'src/_MSI/clawPDF_0.9.3_setup.msi', package):
        handle = api.U()
        api.check(api.msi.MsiOpenDatabaseW(str(path), None, api.c.byref(handle)))
        databases.append(handle)
    original, new = databases
    props = {r['Property']: r['Value'] for r in api.query(new, 'SELECT * FROM `Property`')}
    old_props = {r['Property']: r['Value'] for r in api.query(original, 'SELECT * FROM `Property`')}
    expected_version = '.'.join(api.assembly_metadata(repo / 'src/_Build/Release')['clawpdf.exe']['Version'].split('.')[:3])
    assert props['ProductVersion'] == expected_version
    assert props['ProductLanguage'] == '2052'
    controls = api.query(new, 'SELECT * FROM `Control`')
    welcome = next(row['Text'] for row in controls if row['Dialog_'] == 'WelcomeForm' and row['Control'] == 'BannerText')
    assert '欢迎安装' in welcome
    assert all(row['FaceName'] == 'Microsoft YaHei UI' for row in api.query(new, 'SELECT * FROM `TextStyle`'))
    assert props['ProductCode'] != old_props['ProductCode']
    assert props['UpgradeCode'] == old_props['UpgradeCode']
    upgrade = api.query(new, 'SELECT * FROM `Upgrade`')
    assert len(upgrade) == 1 and upgrade[0]['VersionMax'] == expected_version
    assert upgrade[0]['UpgradeCode'] == old_props['UpgradeCode']
    assert api.query(new, 'SELECT * FROM `CustomAction`') == api.query(original, 'SELECT * FROM `CustomAction`')
    sequence = {r['Action']: int(r['Sequence']) for r in api.query(new, 'SELECT * FROM `InstallExecuteSequence`')}
    assert sequence['InstallInitialize'] < sequence['RemoveExistingProducts'] < sequence['ProcessComponents']

    directories = {r['Directory']: r for r in api.query(new, 'SELECT * FROM `Directory`')}
    components = {r['Component']: r for r in api.query(new, 'SELECT * FROM `Component`')}
    def relative_directory(key):
        if key == 'TARGETDIR':
            return Path()
        row = directories[key]
        name = row['DefaultDir'].split(':')[0].split('|')[-1]
        return relative_directory(row['Directory_Parent']) / ('' if name == '.' else name)

    for handle, name in ((original, 'original'), (new, 'new')):
        api.export_cabinet(handle, folder / (name + '.cab'))
        (folder / name).mkdir()
        api.run_tool(['expand.exe', '-F:*', str(folder / (name + '.cab')), str(folder / name)], folder)

    files = api.query(new, 'SELECT * FROM `File`')
    assert len(files) == 113
    assert int(api.query(new, 'SELECT * FROM `Media`')[0]['LastSequence']) == max(int(r['Sequence']) for r in files)
    hashes = {r['File_']: r for r in api.query(new, 'SELECT * FROM `MsiFileHash`')}
    identities = api.assembly_metadata(repo / 'src/_Build/Release')
    drivers = replacements = 0
    for row in files:
        content = (folder / 'new' / row['File']).read_bytes()
        assert len(content) == int(row['FileSize']), row['FileName']
        relative = relative_directory(components[row['Component_']]['Directory_']) / row['FileName'].split('|')[-1]
        if relative.parts[0] == 'clawmon':
            assert content == (folder / 'original' / row['File']).read_bytes(), f'Upstream driver differs: {relative}'
            drivers += 1
        built = repo / 'src/_Build/Release' / relative
        if built.is_file():
            assert content == built.read_bytes(), f'Build mismatch: {relative}'
            replacements += 1
        target = layout / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(folder / 'new' / row['File'], target)
        if not row['Version']:
            parts = struct.unpack('<4i', hashlib.md5(content).digest())
            assert parts == tuple(int(hashes[row['File']]['HashPart' + str(i)]) for i in range(1, 5)), relative
        identity = identities.get(relative.name.lower())
        if identity:
            for entry in api.query(new, f"SELECT * FROM `MsiAssemblyName` WHERE `Component_` = '{row['Component_']}'"):
                assert entry['Value'] == identity[entry['Name']], (relative, entry)
            if not identity['PublicKeyToken']:
                assert not any(entry['Name'] == 'PublicKeyToken' for entry in api.query(new, f"SELECT * FROM `MsiAssemblyName` WHERE `Component_` = '{row['Component_']}'"))
    assert drivers == 21 and (layout / 'clawPDF.Bridge.exe.config').is_file()
    assert (layout / 'clawPDF.WordAddIn.dll').is_file()
    report = dict(validation='PASS', version=props['ProductVersion'], files=len(files),
                  preserved_driver_files=drivers, matched_build_files=replacements,
                  sha256=hashlib.sha256(package.read_bytes()).hexdigest(), layout=str(layout),
                  actual_installation_executed=False)
    (folder / 'validation.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(json.dumps(report, indent=2))
finally:
    for handle in databases:
        api.msi.MsiCloseHandle(handle)
