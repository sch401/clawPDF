using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using NLog;

namespace clawSoft.clawPDF.PrintFile
{
    internal static class WordPdfExporter
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool IsWordDocument(string filename)
        {
            var extension = Path.GetExtension(filename);
            return string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(extension, ".docm", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsAvailable => Type.GetTypeFromProgID("Word.Application") != null;

        public static void Export(string filename, string outputFilename)
        {
            if (!IsWordDocument(filename) || !File.Exists(filename))
                throw new ArgumentException("A Word document that exists is required.", nameof(filename));

            ExportCore(filename, outputFilename);
        }

        public static void ExportSnapshot(string filename, string outputFilename)
        {
            if (!File.Exists(filename) || !string.Equals(Path.GetExtension(filename), ".xml", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("A Word Flat OPC snapshot is required.", nameof(filename));
            ExportCore(filename, outputFilename);
        }

        private static void ExportCore(string filename, string outputFilename)
        {

            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new InvalidOperationException("Microsoft Word must be installed for native PDF export.");

            dynamic word = null;
            dynamic documents = null;
            dynamic document = null;
            try
            {
                // Use a separate Word instance so open documents and unsaved edits stay with the user.
                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone
                word.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
                documents = word.Documents;
                document = documents.Open(FileName: Path.GetFullPath(filename), ConfirmConversions: false,
                    ReadOnly: true, AddToRecentFiles: false, PasswordDocument: "", PasswordTemplate: "",
                    WritePasswordDocument: "", WritePasswordTemplate: "", Visible: false,
                    OpenAndRepair: false, NoEncodingDialog: true);

                EnableNavigation(document);

                // ponytail: synchronous Office export; use an isolated worker if unattended jobs need timeouts.
                document.ExportAsFixedFormat(OutputFileName: Path.GetFullPath(outputFilename), ExportFormat: 17,
                    OpenAfterExport: false, OptimizeFor: 0, CreateBookmarks: 1, DocStructureTags: true,
                    UseISO19005_1: false); // PDF, print quality, heading bookmarks
                if (!File.Exists(outputFilename) || new FileInfo(outputFilename).Length == 0)
                    throw new IOException("Word did not produce a PDF.");
            }
            finally
            {
                try
                {
                    if (document != null)
                        document.Close(SaveChanges: 0); // wdDoNotSaveChanges
                }
                catch (Exception ex)
                {
                    Logger.Warn(ex, "Could not close the Word export document.");
                }
                Release((object)document);
                Release((object)documents);
                try
                {
                    if (word != null)
                        word.Quit(SaveChanges: 0);
                }
                catch (Exception ex)
                {
                    Logger.Warn(ex, "Could not quit the Word export instance.");
                }
                Release((object)word);
            }
        }

        private static void EnableNavigation(dynamic document)
        {
            dynamic fields = document.Fields;
            try
            {
                // Work backwards: updating a TOC may replace its nested fields.
                for (var index = (int)fields.Count; index >= 1; index--)
                {
                    dynamic field = fields[index];
                    dynamic codeRange = null;
                    try
                    {
                        int type = field.Type;
                        // REF, TOC, PAGEREF and NOTEREF all support the hyperlink switch.
                        if (type != 3 && type != 13 && type != 37 && type != 72)
                            continue;
                        if ((bool)field.Locked)
                            continue;
                        codeRange = field.Code;
                        string code = codeRange.Text;
                        if (!Regex.IsMatch(code, @"(?:^|\s)\\h(?=\s|$)", RegexOptions.IgnoreCase))
                            codeRange.Text = code.TrimEnd() + @" \h ";
                        if (!(bool)field.Update())
                            Logger.Warn("Word could not update navigation field: {0}", code);
                    }
                    finally
                    {
                        Release((object)codeRange);
                        Release((object)field);
                    }
                }
                document.Repaginate();
            }
            finally
            {
                Release((object)fields);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.FinalReleaseComObject(value);
        }
    }
}
