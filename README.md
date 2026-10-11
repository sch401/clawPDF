# clawPDF - Virtual PDF/OCR/Image (Network) Printer

ClawPDF may seem like yet another Virtual PDF/OCR/Image Printer, but it actually comes packed with features that are typically found in enterprise solutions. With clawPDF, you can create documents in various formats, including PDF/A-1b, PDF/A-2b, PDF/A-3b, PDF/X, PDF/Image, OCR, SVG, PNG, JPEG, TIF, and TXT. You also have easy access to metadata and can remove it before sharing a document. In addition, you can protect your documents with a password and encrypt them with up to 256-bit AES.

ClawPDF offers a scripting interface that lets you automate processes and integrate it into your application. Moreover, you can install clawPDF on a print server and print documents over the network, not just locally.

ClawPDF is open-source and compatible with all major Windows client and server operating systems (x86/x64/ARM64), and it even supports multi-user environments!

# Download

## This fork: native Word PDF export and print integration (0.9.7)

The application and installed product are named **clawPDF**. Version 0.9.7 shares
one COM registration routine across 32-bit and 64-bit registry views, reports
registration failures, and checks the selected printer before reading driver
details. Print job titles use the document name instead of its full disk path.

Close Word before installing the MSI, reopen it, and use **File > Print >
clawPDF** with the normal **PDF** profile. The installed Word COM add-in cancels
the PostScript print job and submits a detached Flat OPC snapshot, including
unsaved edits, to the native Word PDF exporter and clawPDF queue. It intercepts
only printers using the clawPDF Virtual Printer driver. This route currently
exports the whole document at one page per sheet; page selection, copies and
multiple pages per sheet are not applied. Other printers retain normal printing.

The separate **Word to PDF (preserve SVG)** button still accepts saved Word files.
New users default to Simplified Chinese. On the first 0.9.6 startup, an empty or
English language preference migrates to Simplified Chinese; other languages are
retained. Later language choices remain editable in Application Settings. The
installer wizard is also localized into Simplified Chinese.

When a `.docx`, `.docm`, or `.doc` file is dropped into clawPDF or submitted through
`/PrintFile` or the COM `PrintFile` API, this fork uses the installed Microsoft Word
PDF exporter before submitting the PDF to the existing clawPDF conversion queue.
SVG graphics that Word preserves in its PDF export stay vector in the normal PDF
profile. Heading styles produce the PDF's chapter bookmarks. TOC, REF, PAGEREF and
NOTEREF fields in the document body are exported as clickable links, including
references originally inserted without the hyperlink option. Locked fields retain
their existing behavior. The changes apply to the export instance only.

Choose PDF for vector output; the PDF/Image and OCR profiles intentionally
render pages as images. Word export runs synchronously and requires desktop Word.

The source document is opened read-only in a separate, hidden Word instance with
macros disabled. Export errors are reported through the existing error handling;
they do not fall back to printing a rasterized version. Save edits in Word before
submitting the document file.

```powershell
clawPDF.exe /PrintFile="C:\Documents\figure.docx"
```

Run `powershell -File .\build.ps1` with Visual Studio's .NET desktop development
tools installed. This builds the unsigned application, bridge, setup helper and Word add-in;
compiler reference assemblies are downloaded locally when the .NET Framework 4.6.2
targeting pack is missing. Output is in `src\_Build\Release`. The native printer
driver is unchanged; an existing clawPDF 0.9.3 installation supplies that driver.

The ZIP build is for an existing clawPDF 0.9.3 installation.
Close any running clawPDF instance, extract the entire ZIP, and run its `clawPDF.exe`.
Drop a saved Word document onto that executable and choose the normal PDF profile.
It uses the existing installed printer and settings. The ZIP does not install or
replace the printer driver, register a new COM server, or change Windows file
associations. Its assemblies are unsigned.

The 0.9.7 MSI includes the upstream printer driver, shortcuts, Word add-in, COM registration
and uninstall actions. It uses the same upgrade family and a new product code.
Old products are removed within the installation transaction before new files are
installed, so libraries rebuilt without the upstream signing key are replaced too.
The MSI is unsigned. Close clawPDF and all Word windows before running the installer.

In Word's COM Add-ins list, **clawPDF 打印组件** must be loaded for the print
integration. If its registration is missing, repair the MSI or run
`SetupHelper.exe /ComInterface=Register` from the installed clawPDF directory as
administrator, then restart Word. Closing a document window can leave an older
Word automation instance running; a newly started Word instance should list
`clawPDF.WordPrint` as connected. Ordinary PDF profiles preserve the incoming
graphics; the Word add-in supplies the vector data for this print route.

To build the MSI after the Release build, use Visual Studio's .NET Framework SDK
`TlbExp.exe` to export `src\_Build\Release\clawPDF.exe` to
`src\_Build\Release\clawPDF.tlb`, then run:

```powershell
python build_installer.py --output clawPDF-0.9.7-setup.msi
```

If Word has loaded a DLL from the default build directory, compile into a separate
directory with `build.ps1 -OutputDirectory .build-tools/package-release` and pass
that directory to `build_installer.py --build-dir .build-tools/package-release`.
Export the type library into that same directory before packaging.

The builder uses Python's standard library, Windows Installer, `expand.exe` and
`makecab.exe`. It preserves the upstream MSI's driver and installation actions,
replaces the application files, updates assembly identities and file hashes, and
includes new runtime configuration files. It creates an MSI without installing it.
`installer_zh_cn.json` supplies the wizard strings. The builder sets the MSI's
Chinese language and string encoding before applying these translations.

`tests\create_word_fixture.py`, `tests\WordPdfRegression.cs` and
`tests\check_word_pdf.py` exercise a three-page document containing SVG, heading
bookmarks, a TOC, REF/PAGEREF fields and an external hyperlink through Word export
and the clawPDF conversion pipeline. See [test instructions](tests/README.md).

https://github.com/clawsoftware/clawPDF/releases/download/0.9.3/clawPDF_0.9.3_setup.msi


# Features

- Print to PDF, PDF/A-1b, PDF/A-2b, PDF/A-3b, PDF/X, PDF/Image, OCR, SVG, PNG, JPEG, TIF and TXT
- Print 100% valid [PDF/A-1b](https://github.com/clawsoftware/clawPDF/raw/master/docs/pdfa_valid/PDFA-1b.pdf), [PDF/A-2b](https://github.com/clawsoftware/clawPDF/raw/master/docs/pdfa_valid/PDFA-2b.pdf) and [PDF/A-3b](https://github.com/clawsoftware/clawPDF/raw/master/docs/pdfa_valid/PDFA-3b.pdf)
- [Optical Character Recognition (OCR)](https://github.com/clawsoftware/clawPDF/wiki/Optical-Character-Recognition-(OCR))
- [Scripting Interface (Python, Powershell, VBScript...)](https://github.com/clawsoftware/clawPDF/wiki/Scripting-Interface)
- [Shared Network Printing](https://github.com/clawsoftware/clawPDF/wiki/Install-as-Network-Printer)
- [SVG Export](https://github.com/clawsoftware/clawPDF/wiki/SVG-Export)
- [Drag and Drop Support](https://github.com/clawsoftware/clawPDF/wiki/Drag-and-Drop)
- [Merge Files](https://github.com/clawsoftware/clawPDF/wiki/Merge-Files)
- [Command Line Support](https://github.com/clawsoftware/clawPDF/wiki/Command-Line-Commands)
- [Silent Printing](https://github.com/clawsoftware/clawPDF/wiki/Silent-Printing)
- [Custom Paper Sizes / Standard Paper Sizes](https://github.com/clawsoftware/clawPDF/wiki/(Custom)-Paper-Sizes)
- 256-bit AES encryption
- Light/Dark Theme
- ARM64 Support
- Full Unicode Support
- Multiple Profiles
- [Post Actions](https://github.com/clawsoftware/clawPDF/wiki/Post-Actions)
- Create additional printers with assigned profile
- [24 translations. Add yours!](https://github.com/clawsoftware/clawPDF/wiki/Translations)
- Easy to deploy (MSI-Installer & Config)
- Many settings
- Easy to use
- No adware, spyware and nagware

# Demo

## Optical Character Recognition (OCR)

![OCR](docs/images/ImageOCR.gif?raw=true "OCR")

## [Scripting Interface](https://github.com/clawsoftware/clawPDF/blob/master/docs/com_examples/Powershell/CreatePDFwithPassword.ps1)

![ScriptingInterface](docs/images/ScriptingInterface.gif?raw=true "Scripting Interface")

## Print a PDF and protect it with a password

![PrintPDFwithPassword](docs/images/PrintPDFwithPassword.gif?raw=true "PrintPDFwithPassword")

## Merge multiple documents

![Merge](docs/images/MergeFiles.gif?raw=true "Merge")

# Tested under

- Windows Server 2022 RDS
- Windows Server 2019 RDS
- Windows Server 2016 RDS
- Windows 11 x64/ARM64
- Windows 10 x86/x64/ARM64
- Windows 8 x86/x64
- Windows 7 x86/x64

# Command Line

## Batch Printing
```
The GUID for the Profile parameter is located under: HKEY_CURRENT_USER\Software\clawSoft\clawPDF\Settings\ConversionProfiles\[id]\Guid

clawPDF.exe /PrintFile=D:\example.docx /profile=f81ea998-3a76-4104-a574-9a66d6f3039b
clawPDF.exe /PrintFile=D:\example.pdf /profile=JpegGuid
clawPDF.exe /PrintFile=D:\example.pdf /profile=JpegGuid /OutputPath=D:\batchjob

clawPDF.exe /PrintFile=D:\example.txt /printerName=clawPDF2
clawPDF.exe /PrintFile=D:\example.docx /printerName=clawJPG
```

## Overwrite Config
```
- To deploy a default configuration in an enterprise environment.
- To export a configuration select "Application Settings -> Debug -> Save settings to file".

clawPDF.exe /Config=D:\clawPDF.ini
```

## Printer Managment
```
SetupHelper.exe /Printer=Add /Name=ExamplePrinter
SetupHelper.exe /Printer=Remove /Name=ExamplePrinter
```

## ManagePrintJobs
```
clawPDF.exe /ManagePrintJobs
```


# Changelog

## v0.9.3 (2023.05.16)

- [bugfix] Fixed a bug where in some cases only administrators could use the shared network printer function
- [bugfix] Fixed Windows 7 issues caused since version 0.9.1 

[more](https://github.com/clawsoftware/clawPDF/wiki/Changelog)


# Requirements

- .Net Framework 4.6.2+
- [Visual C++ Redistributable 14](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist#visual-studio-2015-2017-2019-and-2022) (Download: [x86](https://aka.ms/vs/17/release/vc_redist.x86.exe)/[x64](https://aka.ms/vs/17/release/vc_redist.x64.exe))


# Build

- Visual Studio 2022

[more](https://github.com/clawsoftware/clawPDF/wiki/Build-it-yourself)

# Third-party

## clawPDF uses the following licensed software or parts of the source code:

- PDFCreator (https://github.com/pdfforge/PDFCreator), licensed under AGPL v3 license.
- Pdftosvg.net (https://github.com/dmester/pdftosvg.net), licensed under MIT license.
- iText7 (https://github.com/itext/itext7-dotnet), licensed under AGPL v3 license.
- Nlog (https://github.com/NLog/NLog), licensed under BSD 3-Clause.
- PdfScribe (https://github.com/stchan/PdfScribe), licensed under AGPL v3 license.
- clawmon (https://github.com/clawsoftware/clawPDF/tree/master/src/clawmon), licensed under GPL v2 license.
- Microsoft Postscript Printer Driver (https://docs.microsoft.com/en-us/windows-hardware/drivers/print/microsoft-postscript-printer-driver), copyright (c) Microsoft Corporation. All rights reserved.
- Ghostscript (https://www.ghostscript.com/download/gsdnld.html), licensed under AGPL v3 license.
- SystemWrapper (https://github.com/jozefizso/SystemWrapper), licensed under Microsoft Public license.
- Ftplib (https://archive.codeplex.com/?p=ftplib), licensed under MIT license.
- DataStorage.dll, licensed under pdfforge Freeware License.
- DynamicTranslator.dll, licensed under pdfforge Freeware License.
- Appbar_save (http://modernuiicons.com/), licensed under Attribution-NoDerivs 3.0 Unported.
- Appbar_cogs (http://modernuiicons.com/), licensed under Attribution-NoDerivs 3.0 Unported.
- Appbar_page_file_pdf (http://modernuiicons.com/), licensed under Attribution-NoDerivs 3.0 Unported.


# License

clawPDF is licensed under AGPL v3 license<br>
Copyright (C) 2023 // Andrew Hess // clawSoft
