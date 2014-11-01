using System;
using System.IO;
using System.Runtime.InteropServices;

namespace VbaLauncher
{
    class VbaLauncher
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage : VbaLauncher.exe {xlsm file} {macro function name}");
                Console.WriteLine("    ex) VbaLauncher.exe test.xlsm Sheet1.TestMacro");
                return;
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                Console.WriteLine("Cannot found Excel Application. Please execute after installing.");
                return;
            }

            dynamic app = Activator.CreateInstance(excelType);
            if (app == null)
            {
                Console.WriteLine("Failed to create Excel Application Instance.");
                return;
            }

            app.DisplayAlerts = false;

            dynamic wb = app.Workbooks.Open(Path.Combine(Environment.CurrentDirectory, args[0]));
            app.Run(args[1]);

            wb.Close();
            //Marshal.ReleaseComObject(wb);

            app.Quit();
            Marshal.ReleaseComObject(app);
        }
    }
}
