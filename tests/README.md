# Word SVG and navigation regression

Requires desktop Microsoft Word, Python with Pillow, pypdf and pdfplumber, and a
Release build of clawPDF. Run from the repository root. Pass the path to Visual
Studio's Roslyn `csc.exe` as `$compiler`.

```powershell
python tests/create_word_fixture.py ../regression
$buildFolder = Join-Path (Get-Location) 'src/_Build/Release'
$references = @('clawPDF.Core.dll', 'clawPDF.Settings.dll', 'clawPDF.Shared.dll',
    'clawPDF.Utilities.dll', 'clawPDF.PDFProcessing.dll', 'SystemInterface.dll',
    'SystemWrapper.dll') | ForEach-Object { '/reference:' + (Join-Path $buildFolder $_) }
& $compiler /nologo /target:exe "/out:$buildFolder/WordPdfRegression.exe" /reference:Microsoft.CSharp.dll @references tests/WordPdfRegression.cs
Copy-Item "$buildFolder/clawPDF.exe.config" "$buildFolder/WordPdfRegression.exe.config"
& "$buildFolder/WordPdfRegression.exe" (Join-Path (Get-Location) '../regression')
python tests/check_word_pdf.py ../regression
```

The generated document contains an SVG, three headings, a TOC, REF and PAGEREF
fields without hyperlink switches, and an external link. Assertions cover:

- Unchanged source file bytes, uppercase DOCX recognition and missing-input rejection.
- Native Word routing without changing the default printer; queue submission
  with the original document title, correct page count and retained PDF copy.
- Word export and the existing Ghostscript conversion pipeline.
- Three pages, correct chapter destinations, clickable TOC and cross-references.
- Preserved external link, vector graphics and zero bitmap images in both PDFs.

Only the child UI process launch is intercepted; the export and queue files are
real. The test uses its own spool folder. This is a document export and conversion
test. Installed printer registration,
COM registration and UI save dialogs are outside its scope. The fixture covers
TOC, REF and PAGEREF; NOTEREF and locked fields are not exercised by this fixture.

To test a real document, pass its path as the second argument:

Pass `--snapshot` as the third argument to create a Flat OPC copy and exercise
the Word print add-in's snapshot export and queue route instead of file import.

```powershell
& "$buildFolder/WordPdfRegression.exe" ../real-document 'C:/Documents/Doc1.docx'
```

This runs Word export, queue submission and Ghostscript conversion with the
source file preserved. The three-page fixture checker applies to the generated
fixture only. In the reported `Doc1.docx`, the native PDF had 929 curves and 102
lines; the final PDF had 877 curves and 102 lines. Both had zero bitmap images.
The original PostScript PDF printed from Word had three bitmap images and no curves.
The 0.9.6 Word add-in was also checked separately using Word's actual File > Print
button with this document, and the Chinese clawPDF save window.

After building the MSI, run the offline installer check from the repository root:

```powershell
python tests/check_installer.py . path/to/clawPDF-0.9.6-WordPrintSVG-zh-CN-setup.msi
```

It verifies product and upgrade identifiers, the upgrade sequence, all cabinet
files and hashes, assembly identities, the unchanged upstream driver files and
custom actions, Chinese installer text and fonts, and agreement with the Release build. It extracts an application
layout into the parent working directory, where the Word regression executable
can be compiled and run. The check does not install the MSI or exercise its actual
installation, upgrade or uninstall operations.
