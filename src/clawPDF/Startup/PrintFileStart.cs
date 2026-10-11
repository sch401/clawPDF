using System.IO;
using clawSoft.clawPDF.Assistants;
using clawSoft.clawPDF.PrintFile;
using System.Windows;
using NLog;

namespace clawSoft.clawPDF.Startup
{
    internal class PrintFileStart : IAppStart
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public PrintFileStart(string printFile, string printerName, string snapshotTitle = null)
        {
            PrintFile = printFile;
            PrinterName = printerName;
            SnapshotTitle = snapshotTitle;
        }

        public string PrintFile { get; }
        public string PrinterName { get; }
        private string SnapshotTitle { get; }

        public bool Run()
        {
            _logger.Info("Launched printjob with PrintFile command.");

            if (string.IsNullOrEmpty(PrintFile))
            {
                _logger.Error("PrintFile Parameter has no argument");
                return false;
            }

            if (!File.Exists(PrintFile))
            {
                _logger.Error("The file \"{0}\" does not exist!", PrintFile);
                return false;
            }

            if (SnapshotTitle != null)
            {
                try
                {
                    if (new PrintCommand(PrintFile, PrinterName).PrintWordSnapshot(SnapshotTitle)) return true;
                    MessageBox.Show("Word 矢量打印失败，请检查文档和 Word 是否可正常打开。", "clawPDF", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                finally
                {
                    // Only remove a snapshot in the add-in's own temporary folder.
                    var folder = Path.GetDirectoryName(Path.GetFullPath(PrintFile));
                    if (string.Equals(Path.GetDirectoryName(folder), Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar), System.StringComparison.OrdinalIgnoreCase)
                        && Path.GetFileName(folder).StartsWith("clawPDF-WordPrint-", System.StringComparison.Ordinal)
                        && Path.GetFileName(PrintFile) == "document.xml")
                    {
                        File.Delete(PrintFile);
                        Directory.Delete(folder);
                    }
                }
            }

            var printFileAssistant = new PrintFileAssistant(PrinterName);

            if (!printFileAssistant.AddFile(PrintFile))
            {
                _logger.Warn("The file \"{0}\" is not printable!", PrintFile);
                return false;
            }

            if (!printFileAssistant.PrintAll())
            {
                _logger.Error("The file \"{0}\" could not be printed!", PrintFile);
                return false;
            }

            return true;
        }
    }
}
