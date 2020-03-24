using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;


namespace WindowTracker
{
    public class WordInteropService
    {
        private const string WORD_CLASS_NAME = "_WwG";
        private const uint DW_OBJECTID = 0xFFFFFFF0;
        private static Guid rrid = new Guid("{00020400-0000-0000-C000-000000000046}");

        public delegate bool EnumChildCallback(int hWnd, ref int lParam);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(int hWnd, uint dwObjectId, byte[] rrid, ref Microsoft.Office.Interop.Word.Window ptr);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(int hWndparent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("user32.dll")]
        public static extern int GetClassName(int hWnd, StringBuilder lpClassname, int mMaxCount);


        private bool EnumChildFunc(int hWndChild, ref int lParam)
        {
            StringBuilder buf = new StringBuilder(128);
            GetClassName(hWndChild, buf, 128);
            if (buf.ToString() == WORD_CLASS_NAME) { lParam = hWndChild; return false; }
            return true;
        }

        public Microsoft.Office.Interop.Word.Application GetOpenWordApplication(Process p)
        {
            Microsoft.Office.Interop.Word.Window ptr = null;
            int hWnd = 0;

            int hWndParent = (int)p.MainWindowHandle;
            if (hWndParent == 0) throw new Exception("Excel Main Window not found.");

            EnumChildWindows(hWndParent, EnumChildFunc, ref hWnd);
            if (hWnd == 0) throw new Exception("Child Window not found.");

            int hr = AccessibleObjectFromWindow(hWnd, DW_OBJECTID, rrid.ToByteArray(), ref ptr);

            return ptr.Application;

        }
    }
}
