using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using clawSoft.clawPDF.Core.Jobs;
using clawSoft.clawPDF.Core.Settings;
using clawSoft.clawPDF.Core.Settings.Enums;
using clawSoft.clawPDF.Shared.Helper;
using clawSoft.clawPDF.Utilities.Process;

internal static class WordPdfRegression
{
    private sealed class TempFolders : ITempFolderProvider
    {
        public string TempFolder { get; set; }
    }

    private sealed class CaptureProcessFactory : ProcessWrapperFactory
    {
        public ProcessStartInfo Submitted { get; private set; }

        private sealed class CaptureProcess : ProcessWrapper
        {
            private readonly Action<ProcessStartInfo> _capture;
            public CaptureProcess(ProcessStartInfo info, Action<ProcessStartInfo> capture) : base(info) { _capture = capture; }
            public override void Start() { _capture(StartInfo); }
        }

        public override ProcessWrapper BuildProcessWrapper(ProcessStartInfo info)
        {
            return new CaptureProcess(info, submitted => Submitted = submitted);
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            var folder = Path.GetFullPath(args[0]);
            Directory.CreateDirectory(folder);
            var source = Path.Combine(folder, "SVG-navigation.DOCX");
            Assert(File.Exists(source), "Run create_word_fixture.py first");
            Console.WriteLine("Exporting the fixture with Word...");
            var hash = Hash(source);
            var assembly = Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "clawPDF.exe"));
            var exporter = assembly.GetType("clawSoft.clawPDF.PrintFile.WordPdfExporter", true);
            Assert((bool)exporter.GetMethod("IsWordDocument").Invoke(null, new object[] { source }), "Uppercase DOCX detection");
            Assert(!(bool)exporter.GetMethod("IsWordDocument").Invoke(null, new object[] { "figure.svg" }), "Non-Word type detection");
            var export = exporter.GetMethod("Export");
            var wordPdf = Path.Combine(folder, "word-native.pdf");
            export.Invoke(null, new object[] { source, wordPdf });
            Assert(hash == Hash(source), "Source document unchanged");

            var commandType = assembly.GetType("clawSoft.clawPDF.PrintFile.PrintCommand", true);
            var command = Activator.CreateInstance(commandType, new object[] { source, "clawPDF" });
            Assert(commandType.GetProperty("CommandType").GetValue(command).ToString() == "ExportWordPdf", "Shared Word route");
            Assert(!(bool)commandType.GetProperty("RequiresDefaultPrinter").GetValue(command), "Word route leaves default printer unchanged");
            try
            {
                export.Invoke(null, new object[] { Path.Combine(folder, "missing.docx"), Path.Combine(folder, "missing.pdf") });
                throw new Exception("Missing input was accepted");
            }
            catch (TargetInvocationException ex)
            {
                Assert(ex.InnerException is ArgumentException, "Missing input is rejected before Word starts");
            }

            Console.WriteLine("Submitting through the shared Word command...");
            // Keep the test job out of the user's real spool; intercept only the UI process launch.
            var queueType = assembly.GetType("clawSoft.clawPDF.JobInfoQueue", true);
            var queue = queueType.GetProperty("Instance").GetValue(null);
            var testSpool = Path.Combine(folder, "submission");
            queueType.GetField("<SpoolFolder>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(queue, testSpool);
            var capture = new CaptureProcessFactory();
            commandType.GetProperty("ProcessWrapperFactory").SetValue(command, capture);
            Assert((bool)commandType.GetMethod("Print", new[] { typeof(TimeSpan) }).Invoke(command, new object[] { TimeSpan.FromSeconds(60) }), "Word command submission");
            Assert(capture.Submitted != null && capture.Submitted.FileName == assembly.Location, "Queue uses the new executable");
            var inf = capture.Submitted.Arguments.Substring("/InfoDataFile=".Length).Trim('"');
            Assert(Path.GetFullPath(inf).StartsWith(testSpool + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase), "Test job confined to its spool");
            var info = new JobInfo(inf);
            Assert(info.SourceFiles[0].DocumentTitle == source && info.SourceFiles[0].TotalPages == 3, "Source title and page count queued");
            Assert(File.Exists(info.SourceFiles[0].Filename), "Queued PDF survives export temporary cleanup");
            Assert(hash == Hash(source), "Source unchanged after command submission");
            Console.WriteLine("Converting the submitted PDF through clawPDF...");
            info.Metadata.Title = "SVG navigation regression";
            info.Metadata.Author = "clawPDF regression";
            var profile = new ConversionProfile { OutputFormat = OutputFormat.Pdf, OpenViewer = false };
            profile.AutoSave.Enabled = true;
            profile.AutoSave.EnsureUniqueFilenames = false;
            var job = new GhostscriptJob(info, profile, new TempFolders { TempFolder = Path.Combine(folder, "jobs") }, new JobTranslations());
            job.AutoCleanUp = false;
            job.OutputFilenameTemplate = Path.Combine(folder, "clawpdf-converted.pdf");
            Assert(job.RunJob() == JobState.Succeeded, "Existing conversion pipeline: " + job.ErrorMessage);
            Assert(job.OutputFiles.Count == 1 && File.Exists(job.OutputFiles[0]), "Final PDF exists");
            info.Delete(true);
            Console.WriteLine("PASS: native export, source preservation, Word routing, queue submission, missing input, conversion pipeline");
            Console.WriteLine(wordPdf);
            Console.WriteLine(job.OutputFiles[0]);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static string Hash(string file)
    {
        using (var hash = SHA256.Create())
        using (var stream = File.OpenRead(file))
            return BitConverter.ToString(hash.ComputeHash(stream));
    }

    private static void Assert(bool result, string check)
    {
        if (!result) throw new Exception("FAIL: " + check);
    }
}
