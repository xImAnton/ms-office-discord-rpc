using System;
using static office_rpc.WinAPIHelper;
using HWND = System.IntPtr;

namespace office_rpc {
    public abstract class OfficeManager {
        public static OfficeWindow GetWindow() {
            OfficeWindow window = null;
            EnumWindows(delegate(IntPtr wnd, int _) {
                switch (GetWindowClass(wnd)) {
                    case "OpusApp":
                        window = GetWordWindow(wnd);
                        return false;
                    case "XLMAIN":
                        window = GetExcelWindow(wnd);
                        return false;
                    default:
                        return true;
                }
            }, 0);
            return window;
        }

        private static OfficeWindow GetExcelWindow(HWND hExcelWindow) {
            OfficeWindow result = new OfficeWindow(null, OfficeWindowType.EXCEL);
            EnumChildWindows(hExcelWindow, (wnd, _) => {
                if (GetWindowClass(wnd) != "XLDESK") return true;
                EnumChildWindows(wnd, (hWnd, _) => {
                    if (GetWindowClass(hWnd) != "EXCEL7") return true;
                    result = new OfficeWindow(GetWindowTitle(hWnd), OfficeWindowType.EXCEL);
                    return false;
                }, 0);
                return false;
            }, 0);
            return result;
        }

        private static OfficeWindow GetWordWindow(HWND hWordWindow) {
            OfficeWindow result = new OfficeWindow(null, OfficeWindowType.WORD);
            EnumChildWindows(hWordWindow, (wnd, _) => {
                if (GetWindowClass(wnd) != "_WwF") return true;
                EnumChildWindows(wnd, (hWnd, _) => {
                    if (GetWindowClass(hWnd) != "_WwB") return true;
                    result = new OfficeWindow(GetWindowTitle(hWnd), OfficeWindowType.WORD);
                    return false;
                }, 0);
            return false;
            }, 0);
            return result;
        }
    }
}