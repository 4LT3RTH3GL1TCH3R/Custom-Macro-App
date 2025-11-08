using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using MacroApp.Engine;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MacroApp;

public partial class MainWindow : Window
{
    // Global hotkey imports
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int VK_F6 = 0x75;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private readonly MacroEngine _engine = new();
    private readonly MacroRecorder _recorder = new();
    private MacroScript? _currentScript;
    private bool _isModified;
    private string? _currentFilePath;
    private IntPtr _hookId = IntPtr.Zero;
    private LowLevelKeyboardProc? _hookProc;

    public MainWindow()
    {
        InitializeComponent();
        InitializeEngine();
        InitializeRecorder();
        SetupGlobalHotkey();
        
        SliderSpeed.ValueChanged += (s, e) => 
        {
            TxtSpeed.Text = $"{SliderSpeed.Value:F1}x";
        };
    }

    private void InitializeEngine()
    {
        _engine.LogMessage += (s, msg) => 
        {
            Dispatcher.Invoke(() => AppendLog(msg));
        };

        _engine.ProgressUpdate += (s, progress) => 
        {
            Dispatcher.Invoke(() => 
            {
                ProgressExecution.Value = progress;
                TxtProgress.Text = $"{progress}%";
            });
        };

        _engine.ExecutionCompleted += (s, e) => 
        {
            Dispatcher.Invoke(() => 
            {
                AppendLog("✓ Execution completed");
                BtnPlay.IsEnabled = true;
                BtnPause.IsEnabled = false;
                BtnStop.IsEnabled = false;
                ProgressExecution.Value = 0;
                TxtProgress.Text = "0%";
            });
        };

        _engine.ExecutionError += (s, ex) => 
        {
            Dispatcher.Invoke(() => 
            {
                AppendLog($"❌ Error: {ex.Message}");
                MessageBox.Show($"Execution error: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            });
        };
    }

    private void InitializeRecorder()
    {
        _recorder.RecordingMessage += (s, msg) => 
        {
            Dispatcher.Invoke(() => AppendLog(msg));
        };
    }

    private void BtnOpen_Click(object sender, RoutedEventArgs e)
    {
        if (_isModified && !PromptSaveChanges())
            return;

        var dialog = new OpenFileDialog
        {
            Filter = "Macro Files (*.mcro)|*.mcro|All Files (*.*)|*.*",
            Title = "Open Macro"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                _currentScript = MacroParser.ParseFile(dialog.FileName);
                _currentFilePath = dialog.FileName;
                LoadScriptToEditor();
                _isModified = false;
                BtnPlay.IsEnabled = true;
                BtnSave.IsEnabled = false;
                TxtFileName.Text = Path.GetFileName(dialog.FileName);
                AppendLog($"Loaded: {dialog.FileName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading macro: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnNew_Click(object sender, RoutedEventArgs e)
    {
        if (_isModified && !PromptSaveChanges())
            return;

        TxtMacroContent.Text = "// New macro script\n// Example:\n// mouse.move(100, 100)\n// wait(500)\n// mouse.click(left)\n";
        _currentScript = null;
        _currentFilePath = null;
        _isModified = false;
        TxtFileName.Text = "Untitled";
        BtnPlay.IsEnabled = false;
        BtnSave.IsEnabled = true;
        AppendLog("New macro created");
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentFilePath))
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Macro Files (*.mcro)|*.mcro",
                Title = "Save Macro",
                DefaultExt = ".mcro"
            };

            if (dialog.ShowDialog() != true)
                return;

            _currentFilePath = dialog.FileName;
        }

        try
        {
            File.WriteAllText(_currentFilePath, TxtMacroContent.Text);
            _currentScript = MacroParser.ParseFile(_currentFilePath);
            _isModified = false;
            BtnSave.IsEnabled = false;
            BtnPlay.IsEnabled = true;
            TxtFileName.Text = Path.GetFileName(_currentFilePath);
            AppendLog($"Saved: {_currentFilePath}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error saving macro: {ex.Message}", "Error", 
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnRecord_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _recorder.StartRecording();
            BtnRecord.IsEnabled = false;
            BtnStopRecord.IsEnabled = true;
            BtnOpen.IsEnabled = false;
            BtnPlay.IsEnabled = false;
            AppendLog("⏺ Recording started - perform your actions...");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error starting recorder: {ex.Message}", "Error", 
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnStopRecord_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _currentScript = _recorder.StopRecording();
            LoadScriptToEditor();
            BtnRecord.IsEnabled = true;
            BtnStopRecord.IsEnabled = false;
            BtnOpen.IsEnabled = true;
            BtnPlay.IsEnabled = true;
            BtnSave.IsEnabled = true;
            _isModified = true;
            TxtFileName.Text = "Recorded Macro (unsaved)";
            AppendLog($"⏹ Recording stopped - {_currentScript.Commands.Count} commands captured");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error stopping recorder: {ex.Message}", "Error", 
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnPlay_Click(object sender, RoutedEventArgs e)
    {
        if (_currentScript == null)
        {
            MessageBox.Show("Please load or create a macro first.", "No Macro", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // Parse current editor content
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, TxtMacroContent.Text);
            _currentScript = MacroParser.ParseFile(tempFile);
            File.Delete(tempFile);

            BtnPlay.IsEnabled = false;
            BtnPause.IsEnabled = true;
            BtnStop.IsEnabled = true;
            BtnOpen.IsEnabled = false;
            BtnRecord.IsEnabled = false;

            AppendLog("▶ Starting execution...");
            await _engine.ExecuteMacroAsync(_currentScript, SliderSpeed.Value);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error executing macro: {ex.Message}", "Error", 
                MessageBoxButton.OK, MessageBoxImage.Error);
            BtnPlay.IsEnabled = true;
            BtnPause.IsEnabled = false;
            BtnStop.IsEnabled = false;
            BtnOpen.IsEnabled = true;
            BtnRecord.IsEnabled = true;
        }
    }

    private void BtnPause_Click(object sender, RoutedEventArgs e)
    {
        if (_engine.IsPaused)
        {
            _engine.Resume();
            BtnPause.Content = "⏸ Pause";
            AppendLog("▶ Resumed");
        }
        else
        {
            _engine.Pause();
            BtnPause.Content = "▶ Resume";
            AppendLog("⏸ Paused");
        }
    }

    private void BtnStop_Click(object sender, RoutedEventArgs e)
    {
        _engine.Stop();
        BtnPlay.IsEnabled = true;
        BtnPause.IsEnabled = false;
        BtnStop.IsEnabled = false;
        BtnOpen.IsEnabled = true;
        BtnRecord.IsEnabled = true;
        BtnPause.Content = "⏸ Pause";
        ProgressExecution.Value = 0;
        TxtProgress.Text = "0%";
        AppendLog("⏹ Stopped");
    }

    private void TxtMacroContent_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_currentScript != null || !string.IsNullOrEmpty(_currentFilePath))
        {
            _isModified = true;
            BtnSave.IsEnabled = true;
        }
    }

    private void BtnClearLog_Click(object sender, RoutedEventArgs e)
    {
        TxtLog.Text = string.Empty;
    }

    private void LoadScriptToEditor()
    {
        if (_currentScript == null) return;

        var lines = new System.Text.StringBuilder();
        foreach (var cmd in _currentScript.Commands)
        {
            string line = cmd.Type switch
            {
                CommandType.MouseClick => $"mouse.click({cmd.Button})",
                CommandType.MouseMove => $"mouse.move({cmd.X}, {cmd.Y})",
                CommandType.MouseGlide => $"mouse.glide({cmd.X}, {cmd.Y}, {cmd.ToX}, {cmd.ToY}, {cmd.GlideDurationMs})",
                CommandType.MouseHold => $"mouse.hold({cmd.Button})",
                CommandType.MouseRelease => $"mouse.release({cmd.Button})",
                CommandType.MouseScroll => $"mouse.scroll({cmd.ScrollAmount})",
                CommandType.KeyboardKey => $"keyboard.key({cmd.Key})",
                CommandType.KeyboardButton => $"keyboard.button({cmd.SpecialKey})",
                CommandType.KeyboardToggle => $"keyboard.toggle({cmd.SpecialKey})",
                CommandType.KeyboardUntoggle => $"keyboard.untoggle({cmd.SpecialKey})",
                CommandType.Wait => $"wait({cmd.DelayMs})",
                CommandType.WindowOpen => $"window.open(\"{cmd.ProcessPath}\")",
                CommandType.WindowClose => $"window.close(\"{cmd.WindowTitle}\")",
                CommandType.WindowMinimize => $"window.minimize(\"{cmd.WindowTitle}\")",
                CommandType.WindowMaximize => $"window.maximize(\"{cmd.WindowTitle}\")",
                CommandType.CmdRun => $"cmd.run(\"{cmd.ShellCommand}\")",
                CommandType.PsRun => $"ps.run(\"{cmd.ShellCommand}\")",
                _ => ""
            };
            lines.AppendLine(line);
        }

        TxtMacroContent.Text = lines.ToString();
    }

    private void AppendLog(string message)
    {
        string timestamp = DateTime.Now.ToString("HH:mm:ss");
        TxtLog.Text += $"[{timestamp}] {message}\n";
    }

    private bool PromptSaveChanges()
    {
        if (!_isModified) return true;

        var result = MessageBox.Show("Save changes to current macro?", "Unsaved Changes", 
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

        if (result == MessageBoxResult.Cancel)
            return false;

        if (result == MessageBoxResult.Yes)
        {
            BtnSave_Click(this, new RoutedEventArgs());
        }

        return true;
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (_isModified && !PromptSaveChanges())
        {
            e.Cancel = true;
            return;
        }

        _engine.Stop();
        
        // Unhook global hotkey
        if (_hookId != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }
        
        base.OnClosing(e);
    }

    private void SetupGlobalHotkey()
    {
        _hookProc = HookCallback;
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule? curModule = curProcess.MainModule)
        {
            if (curModule != null)
            {
                _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
        {
            int vkCode = Marshal.ReadInt32(lParam);
            if (vkCode == VK_F6)
            {
                // Stop macro execution when F6 is pressed
                Dispatcher.Invoke(() =>
                {
                    if (_engine.IsRunning)
                    {
                        BtnStop_Click(this, new RoutedEventArgs());
                    }
                });
            }
        }
        return CallNextHookEx(_hookId, nCode, wParam, lParam);
    }
}
