using System;
using System.IO;
using System.Reflection;
class PrinterRegression
{
    static int Main(string[] args)
    {
        try
        {
            var assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
            var method = assembly.GetType("clawSoft.clawPDF.WordAddIn.WordPrintAddIn")
                .GetMethod("FindPrinter", BindingFlags.NonPublic | BindingFlags.Static);
            Check(method, null, null);
            Check(method, " ", null);
            Check(method, "clawPDF", "clawPDF");
            Check(method, "CLAWPDF on Ne01:", "clawPDF");
            Check(method, "clawPDF 在 Ne01:", "clawPDF");
            Check(method, "clawPDFExtra", null);
            Check(method, "Microsoft Print to PDF", null);
            Check(method, "Microsoft Print to PDF on Ne00:", null);
            Console.WriteLine("PASS: clawPDF printer recognition, localized ports, empty input and other printers");
            return 0;
        }
        catch(Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    static void Check(MethodInfo method, string input, string expected)
    {
        var actual = (string)method.Invoke(null, new object[]{input});
        if(actual != expected) throw new Exception("Unexpected printer match: " + input + " => " + actual);
    }
}
