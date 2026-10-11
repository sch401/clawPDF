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
