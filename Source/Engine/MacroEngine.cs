using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq;

namespace MacroApp.Engine;

public class MacroEngine
{
    // Windows API imports for precise input simulation
    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, IntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    // Mouse event flags
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;
    private const uint MOUSEEVENTF_MOVE = 0x0001;
    private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
    private const int WHEEL_DELTA = 120;

    // Keyboard event flags
    private const uint KEYEVENTF_KEYDOWN = 0x0000;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

    // ShowWindow constants
    private const int SW_MINIMIZE = 6;
    private const int SW_MAXIMIZE = 3;
    private const int SW_RESTORE = 9;
    private const int SW_SHOW = 5;

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isPaused;
    private readonly Stopwatch _executionTimer = new();

    public bool IsRunning => _cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested;
    public bool IsPaused => _isPaused;
    public TimeSpan ElapsedTime => _executionTimer.Elapsed;

    public event EventHandler<string>? LogMessage;
    public event EventHandler<int>? ProgressUpdate;
    public event EventHandler? ExecutionCompleted;
    public event EventHandler<Exception>? ExecutionError;

    public async Task ExecuteMacroAsync(MacroScript script, double speedMultiplier = 1.0)
    {
        _cancellationTokenSource = new CancellationTokenSource();
        _isPaused = false;
        _executionTimer.Restart();

        try
        {
            await Task.Run(async () =>
            {
                for (int i = 0; i < script.Commands.Count; i++)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                        break;

                    while (_isPaused)
                    {
                        await Task.Delay(50, _cancellationTokenSource.Token);
                    }

                    var command = script.Commands[i];
                    await ExecuteCommandAsync(command, speedMultiplier, _cancellationTokenSource.Token);

                    ProgressUpdate?.Invoke(this, (int)((i + 1) / (double)script.Commands.Count * 100));

                    // Apply speed multiplier to delays
                    if (command.DelayMs > 0)
                    {
                        int adjustedDelay = (int)(command.DelayMs / speedMultiplier);
                        await Task.Delay(adjustedDelay, _cancellationTokenSource.Token);
                    }
                }
            }, _cancellationTokenSource.Token);

            _executionTimer.Stop();
            ExecutionCompleted?.Invoke(this, EventArgs.Empty);
        }
        catch (OperationCanceledException)
        {
            LogMessage?.Invoke(this, "Macro execution stopped");
        }
        catch (Exception ex)
        {
            ExecutionError?.Invoke(this, ex);
        }
        finally
        {
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    private async Task ExecuteCommandAsync(MacroCommand command, double speedMultiplier, CancellationToken token)
    {
        try
        {
            switch (command.Type)
            {
                case CommandType.MouseClick:
                    MouseClick(command.Button);
                    LogMessage?.Invoke(this, $"Mouse click: {command.Button}");
                    break;

                case CommandType.MouseMove:
                    MouseMove(command.X, command.Y);
                    LogMessage?.Invoke(this, $"Mouse move: ({command.X}, {command.Y})");
                    break;

                case CommandType.MouseGlide:
                    await MouseGlideAsync(command.X, command.Y, command.ToX, command.ToY, command.GlideDurationMs, speedMultiplier, token);
                    LogMessage?.Invoke(this, $"Mouse glide: ({command.X}, {command.Y}) -> ({command.ToX}, {command.ToY})");
                    break;

                case CommandType.MouseHold:
                    MouseHold(command.Button);
                    LogMessage?.Invoke(this, $"Mouse hold: {command.Button}");
                    break;

                case CommandType.MouseRelease:
                    MouseRelease(command.Button);
                    LogMessage?.Invoke(this, $"Mouse release: {command.Button}");
                    break;

                case CommandType.MouseScroll:
                    MouseScroll(command.ScrollAmount);
                    LogMessage?.Invoke(this, $"Mouse scroll: {command.ScrollAmount}");
                    break;

                case CommandType.KeyboardKey:
                    KeyboardKey(command.Key);
                    LogMessage?.Invoke(this, $"Keyboard key: {command.Key}");
                    break;

                case CommandType.KeyboardButton:
                    KeyboardButton(command.SpecialKey);
                    LogMessage?.Invoke(this, $"Keyboard button: {command.SpecialKey}");
                    break;

                case CommandType.KeyboardToggle:
                    KeyboardToggle(command.SpecialKey, true);
                    LogMessage?.Invoke(this, $"Keyboard toggle ON: {command.SpecialKey}");
                    break;

                case CommandType.KeyboardUntoggle:
                    KeyboardToggle(command.SpecialKey, false);
                    LogMessage?.Invoke(this, $"Keyboard toggle OFF: {command.SpecialKey}");
                    break;

                case CommandType.Wait:
                    LogMessage?.Invoke(this, $"Wait: {command.DelayMs}ms");
                    break;

                case CommandType.WindowOpen:
                    WindowOpen(command.ProcessPath);
                    LogMessage?.Invoke(this, $"Window open: {command.ProcessPath}");
                    break;

                case CommandType.WindowClose:
                    WindowClose(command.WindowTitle);
                    LogMessage?.Invoke(this, $"Window close: {command.WindowTitle}");
                    break;

                case CommandType.WindowMinimize:
                    WindowMinimize(command.WindowTitle);
                    LogMessage?.Invoke(this, $"Window minimize: {command.WindowTitle}");
                    break;

                case CommandType.WindowMaximize:
                    WindowMaximize(command.WindowTitle);
                    LogMessage?.Invoke(this, $"Window maximize: {command.WindowTitle}");
                    break;

                case CommandType.CmdRun:
                    CmdRun(command.ShellCommand);
                    LogMessage?.Invoke(this, $"CMD run: {command.ShellCommand}");
                    break;

                case CommandType.PsRun:
                    PsRun(command.ShellCommand);
                    LogMessage?.Invoke(this, $"PowerShell run: {command.ShellCommand}");
                    break;
            }
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Error executing command: {ex.Message}");
        }
    }

    private void MouseClick(string button)
    {
        switch (button.ToLower())
        {
            case "left":
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
                Thread.Sleep(10); // Minimum realistic click duration
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
                break;
            case "right":
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
                Thread.Sleep(10);
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
                break;
            case "middle":
                mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, IntPtr.Zero);
                Thread.Sleep(10);
                mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, IntPtr.Zero);
                break;
        }
    }

    private void MouseMove(int x, int y)
    {
        SetCursorPos(x, y);
    }

    private async Task MouseGlideAsync(int fromX, int fromY, int toX, int toY, int durationMs, double speedMultiplier, CancellationToken token)
    {
        // Adjust duration by speed multiplier
        int adjustedDuration = (int)(durationMs / speedMultiplier);
        
        // Use at least 10ms for ultra-fast glides
        if (adjustedDuration < 10)
            adjustedDuration = 10;

        // Calculate total distance
        double deltaX = toX - fromX;
        double deltaY = toY - fromY;
        double distance = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);

        // If distance is too small, just move directly
        if (distance < 2)
        {
            SetCursorPos(toX, toY);
            return;
        }

        // Number of steps (aim for smooth movement at ~60fps)
        int steps = Math.Max((int)(adjustedDuration / 16.67), 2);
        int delayPerStep = adjustedDuration / steps;

        // Ensure we have at least 1ms delay per step
        if (delayPerStep < 1)
        {
            delayPerStep = 1;
            steps = adjustedDuration;
        }

        // Perform smooth interpolation
        for (int i = 0; i <= steps; i++)
        {
            if (token.IsCancellationRequested)
                break;

            while (_isPaused)
            {
                await Task.Delay(50, token);
            }

            // Ease-in-out interpolation for more natural movement
            double t = (double)i / steps;
            double eased = EaseInOutCubic(t);

            int currentX = fromX + (int)(deltaX * eased);
            int currentY = fromY + (int)(deltaY * eased);

            SetCursorPos(currentX, currentY);

            if (i < steps)
                await Task.Delay(delayPerStep, token);
        }

        // Ensure we end at exact target position
        SetCursorPos(toX, toY);
    }

    private double EaseInOutCubic(double t)
    {
        // Smooth ease-in-out cubic function
        return t < 0.5 
            ? 4 * t * t * t 
            : 1 - Math.Pow(-2 * t + 2, 3) / 2;
    }

    private void MouseHold(string button)
    {
        switch (button.ToLower())
        {
            case "left":
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
                break;
            case "right":
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
                break;
            case "middle":
                mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, IntPtr.Zero);
                break;
        }
    }

    private void MouseRelease(string button)
    {
        switch (button.ToLower())
        {
            case "left":
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
                break;
            case "right":
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
                break;
            case "middle":
                mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, IntPtr.Zero);
                break;
        }
    }

    private void MouseScroll(int amount)
    {
        // Positive = scroll up, Negative = scroll down
        // Windows expects scroll in multiples of WHEEL_DELTA (120)
        int scrollValue = amount * WHEEL_DELTA;
        mouse_event(MOUSEEVENTF_WHEEL, 0, 0, (uint)scrollValue, IntPtr.Zero);
    }

    private void KeyboardKey(string key)
    {
        if (key.Length == 1)
        {
            char c = key[0];
            bool isUpperCase = char.IsUpper(c);
            
            // Press shift if uppercase letter
            if (isUpperCase)
            {
                keybd_event(0x10, 0, KEYEVENTF_KEYDOWN, IntPtr.Zero); // VK_SHIFT
            }
            
            byte vkCode = VirtualKeyFromChar(c);
            keybd_event(vkCode, 0, KEYEVENTF_KEYDOWN, IntPtr.Zero);
            Thread.Sleep(10);
            keybd_event(vkCode, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
            
            // Release shift if it was pressed
            if (isUpperCase)
            {
                keybd_event(0x10, 0, KEYEVENTF_KEYUP, IntPtr.Zero); // VK_SHIFT
            }
        }
    }

    private void KeyboardButton(string specialKey)
    {
        byte vkCode = GetVirtualKeyCode(specialKey);
        bool isExtended = IsExtendedKey(specialKey);
        
        uint flags = isExtended ? KEYEVENTF_EXTENDEDKEY : 0;
        keybd_event(vkCode, 0, flags | KEYEVENTF_KEYDOWN, IntPtr.Zero);
        Thread.Sleep(10);
        keybd_event(vkCode, 0, flags | KEYEVENTF_KEYUP, IntPtr.Zero);
    }

    private void KeyboardToggle(string specialKey, bool down)
    {
        byte vkCode = GetVirtualKeyCode(specialKey);
        bool isExtended = IsExtendedKey(specialKey);
        uint flags = isExtended ? KEYEVENTF_EXTENDEDKEY : 0;

        if (down)
        {
            keybd_event(vkCode, 0, flags | KEYEVENTF_KEYDOWN, IntPtr.Zero);
        }
        else
        {
            keybd_event(vkCode, 0, flags | KEYEVENTF_KEYUP, IntPtr.Zero);
        }
    }

    public void Pause()
    {
        _isPaused = true;
        _executionTimer.Stop();
    }

    public void Resume()
    {
        _isPaused = false;
        _executionTimer.Start();
    }

    public void Stop()
    {
        _cancellationTokenSource?.Cancel();
        _executionTimer.Stop();
    }

    private byte VirtualKeyFromChar(char c)
    {
        char upper = char.ToUpper(c);
        if (upper >= 'A' && upper <= 'Z')
            return (byte)upper;
        
        return c switch
        {
            '0' => 0x30, '1' => 0x31, '2' => 0x32, '3' => 0x33, '4' => 0x34,
            '5' => 0x35, '6' => 0x36, '7' => 0x37, '8' => 0x38, '9' => 0x39,
            ' ' => 0x20,
            _ => 0
        };
    }

    private byte GetVirtualKeyCode(string key)
    {
        return key.ToLower() switch
        {
            "tab" => 0x09,
            "enter" => 0x0D,
            "shift" => 0x10,
            "ctrl" => 0x11,
            "alt" => 0x12,
            "caps" => 0x14,
            "esc" => 0x1B,
            "space" => 0x20,
            "pageup" => 0x21,
            "pagedown" => 0x22,
            "end" => 0x23,
            "home" => 0x24,
            "left" => 0x25,
            "up" => 0x26,
            "right" => 0x27,
            "down" => 0x28,
            "delete" => 0x2E,
            "f1" => 0x70, "f2" => 0x71, "f3" => 0x72, "f4" => 0x73,
            "f5" => 0x74, "f6" => 0x75, "f7" => 0x76, "f8" => 0x77,
            "f9" => 0x78, "f10" => 0x79, "f11" => 0x7A, "f12" => 0x7B,
            "win" => 0x5B,
            "lshift" => 0xA0,
            "rshift" => 0xA1,
            "lctrl" => 0xA2,
            "rctrl" => 0xA3,
            "lalt" => 0xA4,
            "ralt" => 0xA5,
            _ => 0
        };
    }

    private bool IsExtendedKey(string key)
    {
        return key.ToLower() switch
        {
            "alt" or "ctrl" or "left" or "up" or "right" or "down" or 
            "home" or "end" or "pageup" or "pagedown" or "delete" => true,
            _ => false
        };
    }

    public static POINT GetCurrentMousePosition()
    {
        GetCursorPos(out POINT point);
        return point;
    }

    private void WindowOpen(string processPath)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = processPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Failed to open window: {ex.Message}");
        }
    }

    private void WindowClose(string windowTitle)
    {
        try
        {
            IntPtr hWnd = FindWindowByTitle(windowTitle);
            if (hWnd != IntPtr.Zero)
            {
                GetWindowThreadProcessId(hWnd, out uint processId);
                Process? process = Process.GetProcessById((int)processId);
                process?.CloseMainWindow();
                
                // If CloseMainWindow fails, force kill after a brief delay
                Thread.Sleep(500);
                if (process != null && !process.HasExited)
                {
                    process.Kill();
                }
            }
            else
            {
                LogMessage?.Invoke(this, $"Window not found: {windowTitle}");
            }
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Failed to close window: {ex.Message}");
        }
    }

    private void WindowMinimize(string windowTitle)
    {
        try
        {
            IntPtr hWnd = FindWindowByTitle(windowTitle);
            if (hWnd != IntPtr.Zero)
            {
                ShowWindow(hWnd, SW_MINIMIZE);
            }
            else
            {
                LogMessage?.Invoke(this, $"Window not found: {windowTitle}");
            }
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Failed to minimize window: {ex.Message}");
        }
    }

    private void WindowMaximize(string windowTitle)
    {
        try
        {
            IntPtr hWnd = FindWindowByTitle(windowTitle);
            if (hWnd != IntPtr.Zero)
            {
                ShowWindow(hWnd, SW_MAXIMIZE);
                SetForegroundWindow(hWnd);
            }
            else
            {
                LogMessage?.Invoke(this, $"Window not found: {windowTitle}");
            }
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Failed to maximize window: {ex.Message}");
        }
    }

    private IntPtr FindWindowByTitle(string title)
    {
        IntPtr foundWindow = IntPtr.Zero;
        
        EnumWindows((hWnd, lParam) =>
        {
            if (!IsWindow(hWnd))
                return true;

            var sb = new System.Text.StringBuilder(256);
            GetWindowText(hWnd, sb, sb.Capacity);
            string windowTitle = sb.ToString();

            if (windowTitle.Contains(title, StringComparison.OrdinalIgnoreCase))
            {
                foundWindow = hWnd;
                return false; // Stop enumeration
            }

            return true; // Continue enumeration
        }, IntPtr.Zero);

        return foundWindow;
    }

    private void CmdRun(string command)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c {command}",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Failed to run CMD command: {ex.Message}");
        }
    }

    private void PsRun(string command)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -Command \"{command}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
        }
        catch (Exception ex)
        {
            LogMessage?.Invoke(this, $"Failed to run PowerShell command: {ex.Message}");
        }
    }
}
