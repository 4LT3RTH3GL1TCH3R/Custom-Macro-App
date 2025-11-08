using System.Collections.Generic;
using System.IO;

namespace MacroApp.Engine;

public enum CommandType
{
    MouseClick,
    MouseDoubleClick,
    MouseDown,
    MouseUp,
    MouseMove,
    MouseGlide,
    MouseHold,
    MouseRelease,
    MouseScroll,
    KeyPress,
    KeyDown,
    KeyUp,
    KeyboardKey,
    KeyboardButton,
    KeyboardToggle,
    KeyboardUntoggle,
    Delay,
    Wait,
    WindowOpen,
    WindowClose,
    WindowMaximize,
    WindowMinimize,
    CmdRun,
    PsRun
}

public class MacroCommand
{
    public CommandType Type { get; set; }
    public string Button { get; set; } = string.Empty;
    public string Key { get; set; } = string.Empty;
    public string SpecialKey { get; set; } = string.Empty;
    public int X { get; set; }
    public int Y { get; set; }
    public int ToX { get; set; }
    public int ToY { get; set; }
    public int GlideDurationMs { get; set; }
    public int ScrollAmount { get; set; }
    public int DelayMs { get; set; }
    public string OriginalLine { get; set; } = string.Empty;
    public string WindowTitle { get; set; } = string.Empty;
    public string ProcessPath { get; set; } = string.Empty;
    public string ShellCommand { get; set; } = string.Empty;
}

    public class MacroScript
    {
        public string FilePath { get; set; } = string.Empty;
        public List<MacroCommand> Commands { get; set; } = new();
        public string Name => Path.GetFileNameWithoutExtension(FilePath);
    }
