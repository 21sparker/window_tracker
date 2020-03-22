using System;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WindowTracker
{
    class WindowHelpers
    {
        #region Window Placement Is Visible
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        private struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Point rcNormalPosition;
        }

        /// <summary>
        /// Returns whether the window placement is visible
        /// </summary>
        /// <param name="hWnd"></param>
        /// <returns></returns>
        public static bool WindowPlacementIsVisible(IntPtr hWnd)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            GetWindowPlacement(hWnd, ref placement);

            return !placement.rcNormalPosition.IsEmpty;

        }
        #endregion

        #region Window is Visible
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        public static bool WindowIsVisible(IntPtr hWnd)
        {
            return IsWindowVisible(hWnd);
        }

        #endregion

        #region Get Foreground Window Handle
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private int GetForegroundWindowHandle()
        {
            IntPtr hndl = GetForegroundWindow();
            return (Int32)hndl;
        }
        #endregion

        #region Get Time Since Last User Input
        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public static readonly int SizeOf = Marshal.SizeOf(typeof(LASTINPUTINFO));

            [MarshalAs(UnmanagedType.U4)]
            public UInt32 cbSize;
            [MarshalAs(UnmanagedType.U4)]
            public UInt32 dwTime;
        }
        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        public static int GetLastInputTime()
        {
            int idleTime = 0;
            LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
            lastInputInfo.dwTime = 0;

            int envTicks = Environment.TickCount;

            if (GetLastInputInfo(ref lastInputInfo))
            {
                int lastInputTick = (int)lastInputInfo.dwTime;
                idleTime = envTicks - lastInputTick;
            }

            return ((idleTime > 0) ? (idleTime / 1000) : 0);
        }
        #endregion

        //https://stackoverflow.com/questions/8431298/process-mainmodule-access-is-denied

        #region Get Process Exe File Path
        [Flags]
        private enum ProcessAccessFlags : uint
        {
            QueryLimitedInformation = 0x00001000
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool QueryFullProcessImageName(
            [In] IntPtr hProcess,
            [In] int dwFlags,
            [Out] StringBuilder lpExeName,
            ref int lpdwSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(
            ProcessAccessFlags processAccess,
            bool bInheritHandle,
            int processId);

        /// <summary>
        /// Returns the executable path of the running process.
        /// </summary>
        /// <param name="process"></param>
        /// <returns></returns>
        public static string GetProcessFileName(Process process)
        {
            string processFileName = String.Empty;
            try
            {
                processFileName = process.MainModule.FileName;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // do nothing, 'Access is denied' is an error I get with some applications (ex: Task Manager)
            }

            if (!String.IsNullOrEmpty(processFileName))
            {
                return processFileName;
            }

            int capacity = 2000;
            StringBuilder builder = new StringBuilder(capacity);
            IntPtr ptr = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, false, process.Id);

            if (!QueryFullProcessImageName(ptr, 0, builder, ref capacity))
            {
                return String.Empty;
            }

            return builder.ToString();
        }

        #endregion
    }
}
