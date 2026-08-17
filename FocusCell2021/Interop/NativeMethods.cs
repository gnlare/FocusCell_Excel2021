using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FocusCell2021.Interop
{
    internal static class NativeMethods
    {
        internal const int GWL_EXSTYLE = -20;
        internal const int WS_EX_TRANSPARENT = 0x00000020;
        internal const int WS_EX_TOOLWINDOW = 0x00000080;
        internal const int WS_EX_NOACTIVATE = 0x08000000;

        internal const int SW_HIDE = 0;
        internal const int SW_SHOWNOACTIVATE = 4;

        internal const uint EVENT_SYSTEM_MOVESIZESTART = 0x000A;
        internal const uint EVENT_SYSTEM_MOVESIZEEND = 0x000B;
        internal const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        internal const int SM_CXVSCROLL = 2;
        internal const int SM_CYHSCROLL = 3;

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct GUITHREADINFO
        {
            public uint cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        internal delegate void WinEventDelegate(
            IntPtr hWinEventHook,
            uint eventType,
            IntPtr hwnd,
            int idObject,
            int idChild,
            uint dwEventThread,
            uint dwmsEventTime);

        internal delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        internal static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetSystemMetricsForDpi")]
        private static extern int GetSystemMetricsForDpi(int nIndex, uint dpi);

        internal static int GetVerticalScrollBarWidth(IntPtr ownerHwnd)
        {
            return GetSystemMetricForWindow(SM_CXVSCROLL, ownerHwnd);
        }

        internal static int GetHorizontalScrollBarHeight(IntPtr ownerHwnd)
        {
            return GetSystemMetricForWindow(SM_CYHSCROLL, ownerHwnd);
        }

        private static int GetSystemMetricForWindow(int metric, IntPtr ownerHwnd)
        {
            try
            {
                uint dpi = ownerHwnd != IntPtr.Zero ? GetDpiForWindow(ownerHwnd) : 0;
                if (dpi > 0)
                {
                    try
                    {
                        int value = GetSystemMetricsForDpi(metric, dpi);
                        if (value > 0) return value;
                    }
                    catch (EntryPointNotFoundException)
                    {
                        // Pre-Windows 10 fallback below.
                    }
                }
            }
            catch { }

            try { return Math.Max(0, GetSystemMetrics(metric)); }
            catch { return 0; }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetWinEventHook(
            uint eventMin,
            uint eventMax,
            IntPtr hmodWinEventProc,
            WinEventDelegate lpfnWinEventProc,
            uint idProcess,
            uint idThread,
            uint dwFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnhookWinEvent(IntPtr hWinEventHook);
    }
}
