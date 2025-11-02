using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DebtRaven
{
    internal class LineWindowManager
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;
        private static readonly IntPtr HWND_TOPMOST = (IntPtr)(-1);
        private static readonly IntPtr HWND_NOTOPMOST = (IntPtr)(-2);
        private static readonly IntPtr HWND_TOP = (IntPtr)(0);
        private static readonly IntPtr HWND_BOTTOM = (IntPtr)(1);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_NOZORDER = 0x0004;

        // Lưu handle và process ID của ứng dụng hiện tại
        private static IntPtr _currentAppHandle = IntPtr.Zero;
        private static uint _currentAppProcessId = 0;
        private static uint _currentAppThreadId = 0;
        private static RECT _originalWindowRect;
        private static IntPtr _lineHandle = IntPtr.Zero;

        /// <summary>
        /// Khởi tạo - lưu thông tin của ứng dụng hiện tại
        /// </summary>
        public static void Initialize()
        {
            _currentAppHandle = GetForegroundWindow();
            if (_currentAppHandle != IntPtr.Zero)
            {
                GetWindowThreadProcessId(_currentAppHandle, out _currentAppProcessId);
                _currentAppThreadId = GetCurrentThreadId();

                // Lưu vị trí ban đầu của window
                GetWindowRect(_currentAppHandle, out _originalWindowRect);
            }
        }

        /// <summary>
        /// Mở LINE và đặt tool ở vị trí có thể nhìn thấy phía sau LINE
        /// </summary>
        public static bool OpenLineKeepToolVisible(string searchText = "")
        {
            try
            {
                // Lưu thông tin ứng dụng hiện tại nếu chưa có
                if (_currentAppHandle == IntPtr.Zero)
                {
                    Initialize();
                }

                // Đảm bảo tool hiển thị trước khi mở LINE
                if (_currentAppHandle != IntPtr.Zero)
                {
                    ShowWindow(_currentAppHandle, SW_SHOW);
                }

                // Bước 1: Đưa LINE lên foreground
                if (!BringLineToFront())
                {
                    return false;
                }

                // Lưu handle của LINE để sử dụng sau
                _lineHandle = GetLineHandle();

                // Bước 2: Đợi LINE focus và gửi phím tắt
                Thread.Sleep(500);
                SendSearchShortcut();

                // Bước 3: Xử lý search text
                if (!string.IsNullOrEmpty(searchText))
                {
                    Thread.Sleep(300);
                    SendText(searchText);
                }
                else
                {
                    Thread.Sleep(300);
                    ClearSearchBox();
                }

                // Bước 4: Đảm bảo tool vẫn visible nhưng LINE ở phía trước
                Task.Run(async () =>
                {
                    await Task.Delay(100);
                    await EnsureToolVisibleBehindLine();
                });

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in OpenLineKeepToolVisible: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Đảm bảo tool vẫn visible và có thể click được, nhưng LINE ở phía trước
        /// </summary>
        private static async Task EnsureToolVisibleBehindLine()
        {
            try
            {
                if (_currentAppHandle != IntPtr.Zero)
                {
                    // Đảm bảo tool không bị minimize
                    if (IsIconic(_currentAppHandle))
                    {
                        ShowWindow(_currentAppHandle, SW_RESTORE);
                    }

                    // Hiển thị tool
                    ShowWindow(_currentAppHandle, SW_SHOW);

                    // Nếu có LINE handle, đưa LINE lên phía trước
                    if (_lineHandle != IntPtr.Zero)
                    {
                        SetForegroundWindow(_lineHandle);
                        BringWindowToTop(_lineHandle);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in EnsureToolVisibleBehindLine: {ex.Message}");
            }
        }

        /// <summary>
        /// Đặt tool ở vị trí có thể nhìn thấy nhưng LINE ở phía trước
        /// </summary>
        private static async Task SetToolBehindLine()
        {
            try
            {
                if (_currentAppHandle != IntPtr.Zero && _lineHandle != IntPtr.Zero)
                {
                    // Đảm bảo tool vẫn visible và restore nếu bị minimize
                    if (IsIconic(_currentAppHandle))
                    {
                        ShowWindow(_currentAppHandle, SW_RESTORE);
                    }

                    // Hiển thị tool nhưng không focus
                    ShowWindow(_currentAppHandle, SW_SHOW);

                    // Đưa LINE lên phía trước (tool vẫn visible phía sau)
                    SetForegroundWindow(_lineHandle);
                    BringWindowToTop(_lineHandle);

                    // Đảm bảo LINE ở top của Z-order
                    SetWindowPos(_lineHandle, HWND_TOP, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in SetToolBehindLine: {ex.Message}");
            }
        }

        /// <summary>
        /// Đưa tool lên phía trước khi cần thao tác
        /// </summary>
        public static bool BringToolToFront()
        {
            try
            {
                if (_currentAppHandle == IntPtr.Zero)
                {
                    Initialize();
                }

                if (_currentAppHandle != IntPtr.Zero)
                {
                    return ForceForegroundWindow(_currentAppHandle);
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in BringToolToFront: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Đưa tool về phía sau LINE sau khi thao tác xong
        /// </summary>
        public static bool SendToolBehindLine()
        {
            try
            {
                if (_currentAppHandle != IntPtr.Zero)
                {
                    // Nếu không có LINE handle, tìm lại
                    if (_lineHandle == IntPtr.Zero)
                    {
                        _lineHandle = GetLineHandle();
                    }

                    if (_lineHandle != IntPtr.Zero)
                    {
                        // Đảm bảo tool vẫn hiển thị
                        ShowWindow(_currentAppHandle, SW_SHOW);

                        // Focus vào LINE để đưa LINE lên phía trước
                        SetForegroundWindow(_lineHandle);
                        BringWindowToTop(_lineHandle);

                        // Đảm bảo LINE ở top
                        SetWindowPos(_lineHandle, HWND_TOP, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

                        return true;
                    }
                    else
                    {
                        // Nếu không tìm thấy LINE, vẫn giữ tool hiển thị
                        ShowWindow(_currentAppHandle, SW_SHOW);
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in SendToolBehindLine: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Lấy handle của LINE window
        /// </summary>
        private static IntPtr GetLineHandle()
        {
            try
            {
                Process[] lineProcesses = Process.GetProcessesByName("LINE");
                if (lineProcesses.Length > 0)
                {
                    return lineProcesses[0].MainWindowHandle;
                }
                return IntPtr.Zero;
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// Kiểm tra xem LINE có đang mở không
        /// </summary>
        public static bool IsLineOpen()
        {
            try
            {
                Process[] lineProcesses = Process.GetProcessesByName("LINE");
                return lineProcesses.Length > 0 && lineProcesses[0].MainWindowHandle != IntPtr.Zero;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Đưa LINE window lên foreground
        /// </summary>
        public static bool BringLineToFront()
        {
            try
            {
                Process[] lineProcesses = Process.GetProcessesByName("LINE");

                if (lineProcesses.Length == 0)
                {
                    Process.Start("line://");
                    Thread.Sleep(3000);
                    lineProcesses = Process.GetProcessesByName("LINE");
                    if (lineProcesses.Length == 0)
                    {
                        return false;
                    }
                }

                Process lineProcess = lineProcesses[0];
                IntPtr hWnd = lineProcess.MainWindowHandle;

                if (hWnd == IntPtr.Zero)
                {
                    Thread.Sleep(1000);
                    lineProcess.Refresh();
                    hWnd = lineProcess.MainWindowHandle;
                    if (hWnd == IntPtr.Zero)
                    {
                        return false;
                    }
                }

                _lineHandle = hWnd; // Lưu handle của LINE
                return ForceForegroundWindow(hWnd);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in BringLineToFront: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Phương thức mạnh mẽ hơn để đưa window lên foreground
        /// </summary>
        private static bool ForceForegroundWindow(IntPtr hWnd)
        {
            try
            {
                if (IsIconic(hWnd))
                {
                    ShowWindow(hWnd, SW_RESTORE);
                    Thread.Sleep(200);
                }

                uint targetThreadId = GetWindowThreadProcessId(hWnd, out uint targetProcessId);
                uint currentThreadId = GetCurrentThreadId();

                bool success = false;

                if (targetThreadId != currentThreadId)
                {
                    AttachThreadInput(currentThreadId, targetThreadId, true);
                    success = SetForegroundWindow(hWnd);
                    BringWindowToTop(hWnd);
                    AttachThreadInput(currentThreadId, targetThreadId, false);
                }
                else
                {
                    success = SetForegroundWindow(hWnd);
                    BringWindowToTop(hWnd);
                }

                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ForceForegroundWindow: {ex.Message}");
                try
                {
                    return SetForegroundWindow(hWnd);
                }
                catch
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Gửi phím tắt Ctrl+Shift+F
        /// </summary>
        private static void SendSearchShortcut()
        {
            SendKeys.SendWait("^+f");
        }

        /// <summary>
        /// Clear search box
        /// </summary>
        private static void ClearSearchBox()
        {
            SendKeys.SendWait("^a");
            Thread.Sleep(50);
            SendKeys.SendWait("{DELETE}");
        }

        /// <summary>
        /// Gửi text vào LINE
        /// </summary>
        /// <param name="text">Text cần gửi</param>
        private static void SendText(string text)
        {
            SendKeys.SendWait("^a");
            Thread.Sleep(50);
            SendKeys.SendWait("{DELETE}");
            Thread.Sleep(50);

            if (!string.IsNullOrEmpty(text))
            {
                SendTextViaClipboard(text);
            }
        }

        /// <summary>
        /// Gửi text thông qua clipboard
        /// </summary>
        /// <param name="text">Text cần gửi</param>
        private static void SendTextViaClipboard(string text)
        {
            try
            {
                string originalClipboard = "";
                bool hasOriginalClipboard = false;

                try
                {
                    if (Clipboard.ContainsText())
                    {
                        originalClipboard = Clipboard.GetText();
                        hasOriginalClipboard = true;
                    }
                }
                catch { }

                Clipboard.SetText(text);
                Thread.Sleep(50);
                SendKeys.SendWait("^v");

                Thread.Sleep(100);
                if (hasOriginalClipboard)
                {
                    try
                    {
                        Clipboard.SetText(originalClipboard);
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        Clipboard.Clear();
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in SendTextViaClipboard: {ex.Message}");
                try
                {
                    string escapedText = EscapeTextForSendKeys(text);
                    SendKeys.SendWait(escapedText);
                }
                catch (Exception ex2)
                {
                    Console.WriteLine($"Fallback SendKeys also failed: {ex2.Message}");
                }
            }
        }

        /// <summary>
        /// Escape text cho SendKeys
        /// </summary>
        private static string EscapeTextForSendKeys(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            text = text.Replace("+", "{+}");
            text = text.Replace("^", "{^}");
            text = text.Replace("%", "{%}");
            text = text.Replace("~", "{~}");
            text = text.Replace("(", "{(}");
            text = text.Replace(")", "{)}");
            text = text.Replace("[", "{[}");
            text = text.Replace("]", "{]}");
            text = text.Replace("{", "{{}");
            text = text.Replace("}", "{}}");

            return text;
        }

        // ===== PHƯƠNG THỨC CHÍNH SỬ DỤNG =====

        /// <summary>
        /// Phương thức chính - mở LINE và giữ tool visible phía sau
        /// </summary>
        public static bool OpenLineAndSearch(string searchText = "")
        {
            return OpenLineKeepToolVisible(searchText);
        }

        /// <summary>
        /// Đưa tool lên phía trước để thao tác
        /// </summary>
        public static bool ShowTool()
        {
            return BringToolToFront();
        }

        /// <summary>
        /// Đưa tool về phía sau LINE sau khi thao tác xong
        /// </summary>
        public static bool HideTool()
        {
            return SendToolBehindLine();
        }

        /// <summary>
        /// Toggle tool - nếu tool đang ở trước thì đưa về sau, và ngược lại
        /// </summary>
        public static bool ToggleTool()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();

                if (foregroundWindow == _currentAppHandle)
                {
                    // Tool đang ở phía trước, đưa về phía sau
                    return HideTool();
                }
                else
                {
                    // Tool đang ở phía sau, đưa lên phía trước
                    return ShowTool();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ToggleTool: {ex.Message}");
                return false;
            }
        }
    }
}