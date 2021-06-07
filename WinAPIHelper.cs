using System.Runtime.InteropServices;
using System.Text;
using HWND = System.IntPtr;

namespace office_rpc {
    public abstract class WinAPIHelper {
        public delegate bool EnumWindowsProc(HWND hWnd, int lParam);
        
        [DllImport("USER32.DLL")]
        public static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        public static extern bool EnumChildWindows(HWND hWnd, EnumWindowsProc enumProc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);
        
        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(HWND hWnd);

        public static string GetWindowTitle(HWND hWnd) {
            int textLength = GetWindowTextLength(hWnd);
            StringBuilder builder = new StringBuilder(textLength);
            GetWindowText(hWnd, builder, textLength + 1);
            return builder.ToString();
        }

        [DllImport("USER32.DLL")]
        private static extern int GetClassName(HWND hWnd, StringBuilder lpClassName, int nMaxCount);

        public static string GetWindowClass(HWND hWnd) {
            int length = 255;
            StringBuilder builder = new StringBuilder(length);
            GetClassName(hWnd, builder, length + 1);
            return builder.ToString();
        }
        
    }
}