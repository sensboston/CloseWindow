using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;

namespace CloseWindow
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

        // Define the callback delegate's type.
        delegate bool EnumDelegate(IntPtr hWnd, int lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        const UInt32 WM_CLOSE = 0x0010;

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        void CloseWindow(IntPtr hwnd) { SendMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero); }

        readonly DispatcherTimer timer = new DispatcherTimer();
        static readonly StringBuilder sb = new StringBuilder(1024);
        static readonly Dictionary<IntPtr, Tuple<string,string>> allVisibleWindows = new Dictionary<IntPtr, Tuple<string,string>>();

        public MainWindow()
        {
            InitializeComponent();
            timer.Interval = TimeSpan.FromMilliseconds(Properties.Settings.Default.IntervalMsec);
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            allVisibleWindows.Clear();
            EnumDesktopWindows(IntPtr.Zero, FilterCallback, IntPtr.Zero);
            foreach (var wnd in allVisibleWindows)
                if (wnd.Value.Item1.IndexOf(Properties.Settings.Default.WindowCaption, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    wnd.Value.Item1.IndexOf(Properties.Settings.Default.WindowClass, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (Properties.Settings.Default.CloseMethod.Equals("SendReturn"))
                    {
                        SetForegroundWindow(wnd.Key);
                        SendKeys.SendWait("{ENTER}");
                    }
                    else CloseWindow(wnd.Key);
                }
        }

        private static bool FilterCallback(IntPtr hWnd, int lParam)
        {
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();
            GetClassName(hWnd, sb, sb.Capacity);
            string className = sb.ToString();

            if (IsWindowVisible(hWnd) && !(string.IsNullOrEmpty(title) || string.IsNullOrEmpty(className)))
            {
                allVisibleWindows[hWnd] = new Tuple<string, string>(title, className);
            }
            return true;
        }
    }
}
