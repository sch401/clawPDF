using System;
using System.IO;
using System.Reflection;
using Microsoft.Win32;
using clawSoft.clawPDF.SetupHelper.Driver;
using clawSoft.clawPDF.Utilities;

namespace clawSoft.clawPDF.SetupHelper
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var showUsage = true;

            var clp = new CommandLineParser(args);

            if (clp.HasArgument("Driver"))
            {
                showUsage = false;
                try
                {
                    switch (clp.GetArgument("Driver").ToLower())
                    {
                        case "add":

                            Actions.InstallclawPDFPrinter();
                            for (int i = 0; i < 3; i++)
                            {
                                if (Actions.IsRepairRequired())
                                {
                                    Actions.UninstallclawPDFPrinter();
                                    Actions.WaitForPrintSpooler();
                                    Actions.InstallclawPDFPrinter();
                                }
                            }
                            break;

                        case "remove":
                            Actions.UninstallclawPDFPrinter();
                            break;

                        default:
                            showUsage = true;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Environment.ExitCode = 1;
                }
            }

            if (clp.HasArgument("Printer"))
            {
                showUsage = false;
                try
                {
                    string name = clp.GetArgument("Name");
                    switch (clp.GetArgument("Printer").ToLower())
                    {
                        case "add":
                            Actions.AddPrinter(name);
                            break;

                        case "remove":
                            Actions.RemovePrinter(name);
                            break;

                        default:
                            showUsage = true;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Environment.ExitCode = 1;
                }
            }

            if (clp.HasArgument("FileExtensions"))
            {
                showUsage = false;
                try
                {
                    switch (clp.GetArgument("FileExtensions").ToLower())
                    {
                        case "add":
                            AddExplorerIntegration();
                            break;

                        case "remove":
                            RemoveExplorerIntegration();
                            break;

                        default:
                            showUsage = true;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Environment.ExitCode = 1;
                }
            }

            if (clp.HasArgument("ComInterface"))
            {
                showUsage = false;
                try
                {
                    switch (clp.GetArgument("ComInterface").ToLower())
                    {
                        case "register":
                            RegisterComInterface();
                            break;

                        case "unregister":
                            UnregisterComInterface();
                            break;

                        default:
                            showUsage = true;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Environment.ExitCode = 1;
                }
            }

            if (showUsage)
                Usage();
        }

        private static void Usage()
        {
            Console.WriteLine("SetupHelper " + Assembly.GetExecutingAssembly().GetName().Version);
            Console.WriteLine();
            Console.WriteLine("usage:");
            Console.WriteLine("SetupHelper.exe [/Printer=Add|Remove /Name=Printer] [/FileExtensions=Add|Remove] [/ComInterface=Register|Unregister]");
        }

        private static void AddExplorerIntegration()
        {
            CallRegAsmForShell("clawPDFShell.dll", "/codebase");
        }

        private static void RemoveExplorerIntegration()
        {
            CallRegAsmForShell("clawPDFShell.dll", "/unregister");
        }

        private static void RegisterComInterface()
        {
            CallRegAsmForShell("clawPDF.exe", "/codebase /tlb");
            CallRegAsmForShell("clawPDF.WordAddIn.dll", "/codebase");
        }

        private static void UnregisterComInterface()
        {
            CallRegAsmForShell("clawPDF.exe", "/unregister");
            CallRegAsmForShell("clawPDF.WordAddIn.dll", "/unregister");
        }

        private static void CallRegAsmForShell(string fileName, string parameters)
        {
            var shellDll = Path.Combine(new AssemblyHelper().GetCurrentAssemblyDirectory(), fileName);
            var shellExecuteHelper = new ShellExecuteHelper();
            var paramString = $"\"{shellDll}\" {parameters}";
            var views = Environment.Is64BitOperatingSystem
                ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
                : new[] { RegistryView.Registry32 };
            foreach (var view in views)
            {
                string dotNetPath;
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (var framework = root.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework"))
                    dotNetPath = framework?.GetValue("InstallRoot") as string;
                if (string.IsNullOrEmpty(dotNetPath))
                    throw new InvalidOperationException("Cannot find .NET Framework for " + view);
                var regAsmPath = Path.Combine(dotNetPath, @"v4.0.30319\RegAsm.exe");
                Console.WriteLine(regAsmPath + " " + paramString);
                var result = shellExecuteHelper.RunAsAdmin(regAsmPath, paramString);
                if (result != ShellExecuteResult.Success)
                    throw new InvalidOperationException("COM registration command failed: " + fileName + " (" + view + ", " + result + ")");
            }
        }
    }
}
