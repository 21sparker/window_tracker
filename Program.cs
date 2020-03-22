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

            foreach (Process proc in allProcesses)
            {
                if (proc.MainWindowHandle != IntPtr.Zero &&
                    !processesToIgnore.Contains(proc.ProcessName) &&
                    !String.IsNullOrWhiteSpace(proc.MainWindowTitle) &&
                    proc.Responding &&
                    WindowHelpers.WindowPlacementIsVisible(proc.MainWindowHandle))
                {

                    Console.WriteLine(proc.ProcessName + proc.Id + proc.Responding + ": " + proc.MainWindowTitle + " (" + proc.MainWindowHandle.ToString() + ")");
                    Console.WriteLine(proc.MainModule.FileName);

                }
            }
            Console.ReadKey();
        }
    }
}
