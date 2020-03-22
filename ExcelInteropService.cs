using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;

namespace WindowTracker
{
    /// <summary>
    /// https://stackoverflow.com/questions/770173/how-to-get-excel-instance-or-excel-instance-clsid-using-the-process-id
    /// https://pastebin.com/F7gkrAST
    /// 
    /// Note, that in the original code, excel is started if not open, which is not necessary here
    /// </summary>
    public class ExcelInteropService
    {
        // When I tested it with the excel main window handle, I got "XLMAIN" as the 
        // excel class name from buf.ToString(), but the child windows are EXCEL7, not sure what this means
        private const string EXCEL_CLASS_NAME = "EXCEL7";
        private const uint DW_OBJECTID = 0xFFFFFFF0;
        private static Guid rrid = new Guid("{00020400-0000-0000-C000-000000000046}");

        public delegate bool EnumChildCallback(int hWnd, ref int lParam);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(int hWnd, uint dwObjectId, byte[] rrid, ref Microsoft.Office.Interop.Excel.Window ptr);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(int hWndparent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("user32.dll")]
        public static extern int GetClassName(int hWnd, StringBuilder lpClassname, int mMaxCount);


        private bool EnumChildFunc(int hWndChild, ref int lParam)
        {
            StringBuilder buf = new StringBuilder(128);
            GetClassName(hWndChild, buf, 128);
            if (buf.ToString() == "EXCEL7") { lParam = hWndChild; return false; }
            return true;
        }

        public Microsoft.Office.Interop.Excel.Application GetOpenExcelApplication(Process p)
        {
            Microsoft.Office.Interop.Excel.Window ptr = null;
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
