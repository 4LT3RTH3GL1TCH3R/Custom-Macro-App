using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Input;

namespace MacroApp.Engine;

public class MacroRecorder
{
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    private const int WH_KEYBOARD_LL = 13;
    private const int WH_MOUSE_LL = 14;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_LBUTTONUP = 0x0202;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_RBUTTONUP = 0x0205;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_MBUTTONUP = 0x0208;
    private const int WM_MOUSEMOVE = 0x0200;
    private const int WM_MOUSEWHEEL = 0x020A;
    private const int WHEEL_DELTA = 120;

    private delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);

    private IntPtr _keyboardHookID = IntPtr.Zero;
    private IntPtr _mouseHookID = IntPtr.Zero;
    private LowLevelProc? _keyboardProc;
    private LowLevelProc? _mouseProc;

    private readonly List<MacroCommand> _recordedCommands = new();
    private readonly Stopwatch _recordingTimer = new();
    private bool _isRecording;
    private MacroEngine.POINT _lastMousePos;
    private readonly HashSet<byte> _heldKeys = new();
    private readonly HashSet<string> _heldMouseButtons = new();

    public bool IsRecording => _isRecording;
    public IReadOnlyList<MacroCommand> RecordedCommands => _recordedCommands.AsReadOnly();

    public event EventHandler<string>? RecordingMessage;

    public void StartRecording()
    {
        if (_isRecording) return;

        _recordedCommands.Clear();
        _heldKeys.Clear();
        _heldMouseButtons.Clear();
        _isRecording = true;
        _recordingTimer.Restart();

        _lastMousePos = MacroEngine.GetCurrentMousePosition();

        _keyboardProc = KeyboardHookCallback;
        _mouseProc = MouseHookCallback;

        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule? curModule = curProcess.MainModule)
        {
            if (curModule != null)
            {
                _keyboardHookID = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, GetModuleHandle(curModule.ModuleName), 0);
                _mouseHookID = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        RecordingMessage?.Invoke(this, "Recording started");
    }

    public MacroScript StopRecording()
    {
        if (!_isRecording) return new MacroScript();

        _isRecording = false;
        _recordingTimer.Stop();

        if (_keyboardHookID != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_keyboardHookID);
            _keyboardHookID = IntPtr.Zero;
        }

        if (_mouseHookID != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_mouseHookID);
            _mouseHookID = IntPtr.Zero;
        }

        // Release any held keys
        foreach (byte key in _heldKeys)
        {
            AddCommand(new MacroCommand
            {
                Type = CommandType.KeyboardUntoggle,
                SpecialKey = GetKeyName(key)
            });
        }
        _heldKeys.Clear();
        
        // Release any held mouse buttons
        foreach (string button in _heldMouseButtons)
        {
            AddCommand(new MacroCommand
            {
                Type = CommandType.MouseRelease,
                Button = button
            });
        }
        _heldMouseButtons.Clear();

        RecordingMessage?.Invoke(this, "Recording stopped");

        return new MacroScript { Commands = new List<MacroCommand>(_recordedCommands) };
    }

    private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _isRecording)
        {
            int vkCode = Marshal.ReadInt32(lParam);
            
            if (wParam == (IntPtr)WM_KEYDOWN)
            {
                if (!_heldKeys.Contains((byte)vkCode))
                {
                    _heldKeys.Add((byte)vkCode);
                    
                    string keyName = GetKeyName((byte)vkCode);
                    if (IsModifierKey((byte)vkCode))
                    {
                        // Use toggle for modifier keys (shift, ctrl, alt)
                        AddCommand(new MacroCommand
                        {
                            Type = CommandType.KeyboardToggle,
                            SpecialKey = keyName
                        });
                    }
                    else if (IsSpecialKey((byte)vkCode))
                    {
                        // Use button for other special keys (enter, tab, etc.)
                        AddCommand(new MacroCommand
                        {
                            Type = CommandType.KeyboardButton,
                            SpecialKey = keyName
                        });
                    }
                    else
                    {
                        // Use key for regular characters
                        AddCommand(new MacroCommand
                        {
                            Type = CommandType.KeyboardKey,
                            Key = keyName
                        });
                    }
                }
            }
            else if (wParam == (IntPtr)WM_KEYUP)
            {
                if (_heldKeys.Contains((byte)vkCode))
                {
                    string keyName = GetKeyName((byte)vkCode);
                    if (IsModifierKey((byte)vkCode))
                    {
                        // Record untoggle for modifier keys
                        AddCommand(new MacroCommand
                        {
                            Type = CommandType.KeyboardUntoggle,
                            SpecialKey = keyName
                        });
                    }
                    _heldKeys.Remove((byte)vkCode);
                }
            }
        }

        return CallNextHookEx(_keyboardHookID, nCode, wParam, lParam);
    }

    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _isRecording)
        {
            var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

            if (wParam == (IntPtr)WM_MOUSEMOVE)
            {
                // Record mouse move if position changed significantly
                if (Math.Abs(hookStruct.pt.X - _lastMousePos.X) > 5 || 
                    Math.Abs(hookStruct.pt.Y - _lastMousePos.Y) > 5)
                {
                    AddCommand(new MacroCommand
                    {
                        Type = CommandType.MouseMove,
                        X = hookStruct.pt.X,
                        Y = hookStruct.pt.Y
                    });
                    _lastMousePos = hookStruct.pt;
                }
            }
            else if (wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                _heldMouseButtons.Add("left");
                AddCommand(new MacroCommand
                {
                    Type = CommandType.MouseHold,
                    Button = "left"
                });
            }
            else if (wParam == (IntPtr)WM_LBUTTONUP)
            {
                _heldMouseButtons.Remove("left");
                AddCommand(new MacroCommand
                {
                    Type = CommandType.MouseRelease,
                    Button = "left"
                });
            }
            else if (wParam == (IntPtr)WM_RBUTTONDOWN)
            {
                _heldMouseButtons.Add("right");
                AddCommand(new MacroCommand
                {
                    Type = CommandType.MouseHold,
                    Button = "right"
                });
            }
            else if (wParam == (IntPtr)WM_RBUTTONUP)
            {
                _heldMouseButtons.Remove("right");
                AddCommand(new MacroCommand
                {
                    Type = CommandType.MouseRelease,
                    Button = "right"
                });
            }
            else if (wParam == (IntPtr)WM_MBUTTONDOWN)
            {
                _heldMouseButtons.Add("middle");
                AddCommand(new MacroCommand
                {
                    Type = CommandType.MouseHold,
                    Button = "middle"
                });
            }
            else if (wParam == (IntPtr)WM_MBUTTONUP)
            {
                _heldMouseButtons.Remove("middle");
                AddCommand(new MacroCommand
                {
                    Type = CommandType.MouseRelease,
                    Button = "middle"
                });
            }
            else if (wParam == (IntPtr)WM_MOUSEWHEEL)
            {
                // Extract scroll delta from mouseData (high-order word)
                short delta = (short)((hookStruct.mouseData >> 16) & 0xFFFF);
                int scrollAmount = delta / WHEEL_DELTA;
                
                if (scrollAmount != 0)
                {
                    AddCommand(new MacroCommand
                    {
                        Type = CommandType.MouseScroll,
                        ScrollAmount = scrollAmount
                    });
                }
            }
        }

        return CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
    }

    private void AddCommand(MacroCommand command)
    {
        // Add timing between commands
        if (_recordedCommands.Count > 0)
        {
            long elapsed = _recordingTimer.ElapsedMilliseconds;
            int delay = (int)elapsed;
            
            if (delay > 50) // Only add wait if delay is significant
            {
                _recordedCommands.Add(new MacroCommand
                {
                    Type = CommandType.Wait,
                    DelayMs = delay
                });
            }
        }

        _recordedCommands.Add(command);
        _recordingTimer.Restart();

        RecordingMessage?.Invoke(this, $"Recorded: {command.Type}");
    }

    private bool IsModifierKey(byte vkCode)
    {
        return vkCode switch
        {
            // Shift, Ctrl, Alt (generic and left/right variants)
            0x10 or 0x11 or 0x12 or 0xA0 or 0xA1 or 0xA2 or 0xA3 or 0xA4 or 0xA5 => true,
            _ => false
        };
    }

    private bool IsSpecialKey(byte vkCode)
    {
        return vkCode switch
        {
            0x09 or 0x0D or 0x10 or 0x11 or 0x12 or 0x14 or 0x1B or 0x20 or
            >= 0x21 and <= 0x28 or 0x2E or >= 0x70 and <= 0x7B or 0x5B or
            >= 0xA0 and <= 0xA5 => true,
            _ => false
        };
    }

    private string GetKeyName(byte vkCode)
    {
        return vkCode switch
        {
            0x09 => "tab", 0x0D => "enter", 0x10 => "shift", 0x11 => "ctrl",
            0x12 => "alt", 0x14 => "caps", 0x1B => "esc", 0x20 => "space",
            0x21 => "pageup", 0x22 => "pagedown", 0x23 => "end", 0x24 => "home",
            0x25 => "left", 0x26 => "up", 0x27 => "right", 0x28 => "down",
            0x2E => "delete",
            >= 0x70 and <= 0x7B => $"f{vkCode - 0x6F}",
            0x5B => "win",
            >= 0x41 and <= 0x5A => ((char)vkCode).ToString().ToLower(),
            >= 0x30 and <= 0x39 => ((char)vkCode).ToString(),
            _ => vkCode.ToString()
        };
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public MacroEngine.POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }
}
