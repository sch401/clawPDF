using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyVersion("0.9.7.0")]
[assembly: AssemblyFileVersion("0.9.7.0")]
[assembly: ComVisible(false)]

namespace clawSoft.clawPDF.WordAddIn
{
    // Office's COM contract; neither a PIA nor VSTO runtime is required.
    [ComVisible(true), ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, int mode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addIn,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int mode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComVisible(true), Guid("D1406A4C-0FF1-4A91-A9A1-40E63A5B5ED7"), ProgId("clawPDF.WordPrint"), ClassInterface(ClassInterfaceType.None)]
    public sealed class WordPrintAddIn : IDTExtensibility2
    {
        private static readonly Guid Events = new Guid("00020A01-0000-0000-C000-000000000046");
        private delegate void BeforePrintHandler([MarshalAs(UnmanagedType.IDispatch)] object document, ref bool cancel);
        private object _word;
        private BeforePrintHandler _beforePrint;

        public void OnConnection(object application, int mode, object addIn, ref Array custom)
        {
            _word = application;
            _beforePrint = BeforePrint;
            ComEventsHelper.Combine(application, Events, 7, _beforePrint);
        }

        public void OnDisconnection(int mode, ref Array custom)
        {
            if (_word != null && _beforePrint != null)
                ComEventsHelper.Remove(_word, Events, 7, _beforePrint);
            _beforePrint = null;
            _word = null; // This Word instance belongs to the user.
        }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        private void BeforePrint(object document, ref bool cancel)
        {
            if (cancel) return;
            dynamic word = _word;
            string printer = FindPrinter((string)word.ActivePrinter);
            if (printer == null) return;
            cancel = true;
            string folder = null;
            try
            {
                dynamic doc = document;
                // ponytail: whole document, one page per sheet; print selection needs a print UI override.
                folder = Path.Combine(Path.GetTempPath(), "clawPDF-WordPrint-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(folder);
                string snapshot = Path.Combine(folder, "document.xml");
                // Flat OPC keeps SVG and unsaved edits; field updates happen only in the detached copy.
                File.WriteAllText(snapshot, (string)doc.WordOpenXML, new UTF8Encoding(false));
                string title = (string)doc.Name;
                string executable = Path.Combine(Path.GetDirectoryName(typeof(WordPrintAddIn).Assembly.Location), "clawPDF.exe");
                Process.Start(new ProcessStartInfo(executable)
                {
                    Arguments = "/WordSnapshot=\"" + snapshot + "\" /DocumentTitle=\"" + title + "\" /PrinterName=\"" + printer + "\"",
                    WorkingDirectory = Path.GetDirectoryName(executable),
                    UseShellExecute = false
                });
                folder = null; // The conversion process owns and removes the snapshot.
            }
            catch (Exception ex)
            {
                if (folder != null)
                {
                    File.Delete(Path.Combine(folder, "document.xml"));
                    Directory.Delete(folder);
                }
                MessageBox.Show("Word 矢量打印失败：\r\n" + ex.Message, "clawPDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string FindPrinter(string activePrinter)
        {
            if (string.IsNullOrWhiteSpace(activePrinter)) return null;
            using (var printers = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Print\Printers"))
            {
                if (printers == null) return null;
                foreach (string name in printers.GetSubKeyNames())
                {
                    if (!string.Equals(activePrinter, name, StringComparison.OrdinalIgnoreCase)
                        && !activePrinter.StartsWith(name + " ", StringComparison.OrdinalIgnoreCase)) continue;
                    using (var printer = printers.OpenSubKey(name))
                        if (string.Equals(printer?.GetValue("Printer Driver") as string, "clawPDF Virtual Printer", StringComparison.OrdinalIgnoreCase))
                            return name;
                }
            }
            return null;
        }

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            using (var key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Office\Word\Addins\clawPDF.WordPrint"))
            {
                key.SetValue("FriendlyName", "clawPDF 打印组件");
                key.SetValue("Description", "保留 SVG、章节书签和目录／交叉引用跳转");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            Registry.LocalMachine.DeleteSubKey(@"SOFTWARE\Microsoft\Office\Word\Addins\clawPDF.WordPrint", false);
        }
    }
}
