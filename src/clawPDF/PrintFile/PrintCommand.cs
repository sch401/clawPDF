using System;
using System.Diagnostics;
using System.IO;
using clawSoft.clawPDF.Core.Jobs;
using clawSoft.clawPDF.PDFProcessing;
using clawSoft.clawPDF.Shared.Helper;
using clawSoft.clawPDF.Utilities;
using clawSoft.clawPDF.Utilities.Process;
using NLog;

namespace clawSoft.clawPDF.PrintFile
{
    internal enum PrintType
    {
        Print,
        PrintTo,
        ExportWordPdf,
        Unprintable
    }

    /// <summary>
    ///     The PrintCommand class provides a way to print a file with clawPDF
    /// </summary>
    internal class PrintCommand
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        ///     Create a new PrintCommand for the given file
        /// </summary>
        /// <param name="filename">The full path to the file that shall be printed</param>
        /// <param name="printer">The printer the command will print to</param>
        public PrintCommand(string filename, string printer)
        {
            if (filename == null)
                throw new ArgumentNullException(nameof(filename));

            Filename = filename;
            Printer = printer;

            if (!File.Exists(filename))
            {
                CommandType = PrintType.Unprintable;
                return;
            }

            var extension = Path.GetExtension(filename);
            if (string.IsNullOrEmpty(extension))
            {
                CommandType = PrintType.Unprintable;
                return;
            }

            if (WordPdfExporter.IsWordDocument(filename))
            {
                CommandType = WordPdfExporter.IsAvailable ? PrintType.ExportWordPdf : PrintType.Unprintable;
                return;
            }

            if (!SupportsPrint() && !SupportsPrintTo())
                CommandType = PrintType.Unprintable;
            else
                CommandType = SupportsPrintTo() ? PrintType.PrintTo : PrintType.Print;
        }

        public ProcessWrapperFactory ProcessWrapperFactory { get; set; } = new ProcessWrapperFactory();

        public PrintType CommandType { get; }

        public string Filename { get; }

        public bool Successful { get; private set; }

        public bool IsPrintable => CommandType != PrintType.Unprintable;

        public bool RequiresDefaultPrinter => CommandType == PrintType.Print;

        public string Printer { get; }

        /// <summary>
        ///     Prints the file.
        /// </summary>
        /// <returns>true, if printing was successful</returns>
        public bool Print()
        {
            return Print(TimeSpan.FromSeconds(60));
        }

        /// <summary>
        ///     Prints the file.
        /// </summary>
        /// <param name="processTimeout">The timespan to wait for the process to finish</param>
        /// <returns>true, if printing was successful</returns>
        public bool Print(TimeSpan processTimeout)
        {
            if (CommandType == PrintType.Unprintable)
                throw new InvalidOperationException("File is not printable");

            if (CommandType == PrintType.ExportWordPdf)
            {
                try
                {
                    ExportWordPdf();
                    Successful = true;
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, "Native Word PDF export failed.");
                    return false;
                }
                return Successful;
            }

            var printerHelper = new PrinterHelper();
            if (CommandType == PrintType.Print && Printer != printerHelper.GetDefaultPrinter())
                throw new InvalidOperationException("The default printer needs to be set in order to print this file");

            var p = ProcessWrapperFactory.BuildProcessWrapper(new ProcessStartInfo(Filename));

            if (SupportsPrintTo())
            {
                p.StartInfo.Verb = "Printto";
                p.StartInfo.Arguments = "\"" + Printer + "\"";
            }
            else
            {
                p.StartInfo.Verb = "Print";
            }

            try
            {
                p.Start();
                p.WaitForExit(processTimeout);

                if (!p.HasExited)
                {
                    Logger.Warn("Process was not finishing after {0} seconds, killing it now...",
                        processTimeout.TotalSeconds);
                    p.Kill();
                }
                else
                {
                    Successful = true;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Exception during print process"
                             + "\r\nType: " + ex.GetType()
                             + "\r\nMessage: " + ex.Message
                );
                return false;
            }

            return Successful;
        }

        public bool PrintWordSnapshot(string documentTitle)
        {
            try
            {
                ExportWordPdf(documentTitle);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Word print snapshot export failed.");
                return false;
            }
        }

        private void ExportWordPdf(string documentTitle = null)
        {
            var tempFolder = Path.Combine(Path.GetTempPath(), "clawPDF-Word-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempFolder);
            var pdfFile = Path.Combine(tempFolder, Path.GetFileNameWithoutExtension(Filename) + ".pdf");
            string infFile = null;
            var submitted = false;
            try
            {
                if (documentTitle == null)
                    WordPdfExporter.Export(Filename, pdfFile);
                else
                    WordPdfExporter.ExportSnapshot(Filename, pdfFile);
                infFile = PsFileHelper.TransformToInfFile(pdfFile, JobInfoQueue.Instance.SpoolFolder, Printer);
                if (string.IsNullOrEmpty(infFile))
                    throw new IOException("Could not queue the PDF exported by Word.");

                var jobInfo = new JobInfo(infFile);
                jobInfo.SourceFiles[0].DocumentTitle = documentTitle ?? Filename;
                jobInfo.SourceFiles[0].TotalPages = PDFProcessor.GetNumberOfPages(pdfFile);
                jobInfo.SaveInf();

                var executable = typeof(PrintCommand).Assembly.Location;
                var process = ProcessWrapperFactory.BuildProcessWrapper(new ProcessStartInfo(executable)
                {
                    Arguments = "/InfoDataFile=\"" + infFile + "\"",
                    WorkingDirectory = Path.GetDirectoryName(executable),
                    UseShellExecute = false
                });
                // The PDF is already copied to the spool; its normal job lifecycle owns that copy.
                process.Start();
                submitted = true;
            }
            finally
            {
                if (!submitted && !string.IsNullOrEmpty(infFile))
                    new JobInfo(infFile).Delete(true);
                try
                {
                    File.Delete(pdfFile);
                    Directory.Delete(tempFolder);
                }
                catch (IOException ex)
                {
                    Logger.Warn(ex, "Could not remove temporary Word export files.");
                }
            }
        }

        private bool SupportsPrint()
        {
            return FileUtil.Instance.FileAssocHasPrint(Path.GetExtension(Filename));
        }

        private bool SupportsPrintTo()
        {
            return FileUtil.Instance.FileAssocHasPrintTo(Path.GetExtension(Filename));
        }
    }
}
