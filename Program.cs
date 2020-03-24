using System.Diagnostics;
using System.Collections.Generic;
using System;

namespace WindowTracker
{
    /// <summary>
    /// List of User32 Functions
    /// http://www.pinvoke.net/default.aspx/user32.GetLastInputInfo
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Process[] allProcesses = Process.GetProcesses();

            List<string> processesToIgnore = new List<string>();
            processesToIgnore.Add("explorer");
            processesToIgnore.Add("ApplicationFrameHost");

            string filename;
            Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Word.Application wordApp;

            ExcelInteropService eis = new ExcelInteropService();
            WordInteropService wis = new WordInteropService();

            foreach (Process proc in allProcesses)
            {
                if (proc.MainWindowHandle != IntPtr.Zero &&
                    !processesToIgnore.Contains(proc.ProcessName) &&
                    !String.IsNullOrWhiteSpace(proc.MainWindowTitle) &&
                    proc.Responding &&
                    WindowHelpers.WindowPlacementIsVisible(proc.MainWindowHandle))
                {
                    Console.WriteLine(proc.ProcessName + proc.Id + proc.Responding + ": " + proc.MainWindowTitle + " (" + proc.MainWindowHandle.ToString() + ")");
                    filename = WindowHelpers.GetProcessFileName(proc);
                    Console.WriteLine(filename);

                    if (filename.ToLower().Contains("excel"))
                    {
                        excelApp = eis.GetOpenExcelApplication(proc);
                        Console.WriteLine(excelApp.ActiveWorkbook.FullName);
                    }
                    else if (filename.ToLower().Contains("word"))
                    {
                        wordApp = wis.GetOpenWordApplication(proc);
                        foreach (Microsoft.Office.Interop.Word.Document doc in wordApp.Documents)
                        {
                            Console.WriteLine(doc.FullName);
                        }
                        
                    }
                }
            }


            Console.ReadKey();
        }
    }
}
