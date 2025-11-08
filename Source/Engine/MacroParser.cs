using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace MacroApp.Engine;

public class MacroParser
{
    public static MacroScript ParseFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Macro file not found: {filePath}");

        var script = new MacroScript { FilePath = filePath };
        var lines = File.ReadAllLines(filePath);

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i].Trim();
            
            // Skip empty lines and comments
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("//") || line.StartsWith("#"))
                continue;

            try
            {
                var command = ParseLine(line);
                script.Commands.Add(command);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error parsing line {i + 1}: {line}\n{ex.Message}");
            }
        }

        return script;
    }

    private static MacroCommand ParseLine(string line)
    {
        var command = new MacroCommand { OriginalLine = line };

        // Parse mouse.click(button)
        var mouseClickMatch = Regex.Match(line, @"mouse\.click\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (mouseClickMatch.Success)
        {
            command.Type = CommandType.MouseClick;
            command.Button = mouseClickMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse mouse.move(x, y)
        var mouseMoveMatch = Regex.Match(line, @"mouse\.move\((\d+)\s*,\s*(\d+)\)", RegexOptions.IgnoreCase);
        if (mouseMoveMatch.Success)
        {
            command.Type = CommandType.MouseMove;
            command.X = int.Parse(mouseMoveMatch.Groups[1].Value);
            command.Y = int.Parse(mouseMoveMatch.Groups[2].Value);
            return command;
        }

        // Parse mouse.glide(fromX, fromY, toX, toY, duration)
        var mouseGlideMatch = Regex.Match(line, @"mouse\.glide\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\)", RegexOptions.IgnoreCase);
        if (mouseGlideMatch.Success)
        {
            command.Type = CommandType.MouseGlide;
            command.X = int.Parse(mouseGlideMatch.Groups[1].Value);
            command.Y = int.Parse(mouseGlideMatch.Groups[2].Value);
            command.ToX = int.Parse(mouseGlideMatch.Groups[3].Value);
            command.ToY = int.Parse(mouseGlideMatch.Groups[4].Value);
            command.GlideDurationMs = int.Parse(mouseGlideMatch.Groups[5].Value);
            return command;
        }

        // Parse mouse.hold(button)
        var mouseHoldMatch = Regex.Match(line, @"mouse\.hold\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (mouseHoldMatch.Success)
        {
            command.Type = CommandType.MouseHold;
            command.Button = mouseHoldMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse mouse.release(button)
        var mouseReleaseMatch = Regex.Match(line, @"mouse\.release\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (mouseReleaseMatch.Success)
        {
            command.Type = CommandType.MouseRelease;
            command.Button = mouseReleaseMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse mouse.scroll(amount) - positive=up, negative=down
        var mouseScrollMatch = Regex.Match(line, @"mouse\.scroll\((-?\d+)\)", RegexOptions.IgnoreCase);
        if (mouseScrollMatch.Success)
        {
            command.Type = CommandType.MouseScroll;
            command.ScrollAmount = int.Parse(mouseScrollMatch.Groups[1].Value);
            return command;
        }

        // Parse keyboard.key(character)
        var keyboardKeyMatch = Regex.Match(line, @"keyboard\.key\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (keyboardKeyMatch.Success)
        {
            command.Type = CommandType.KeyboardKey;
            command.Key = keyboardKeyMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse keyboard.button(special_key)
        var keyboardButtonMatch = Regex.Match(line, @"keyboard\.button\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (keyboardButtonMatch.Success)
        {
            command.Type = CommandType.KeyboardButton;
            command.SpecialKey = keyboardButtonMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse keyboard.toggle(special_key)
        var keyboardToggleMatch = Regex.Match(line, @"keyboard\.toggle\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (keyboardToggleMatch.Success)
        {
            command.Type = CommandType.KeyboardToggle;
            command.SpecialKey = keyboardToggleMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse keyboard.untoggle(special_key)
        var keyboardUntoggleMatch = Regex.Match(line, @"keyboard\.untoggle\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (keyboardUntoggleMatch.Success)
        {
            command.Type = CommandType.KeyboardUntoggle;
            command.SpecialKey = keyboardUntoggleMatch.Groups[1].Value.Trim();
            return command;
        }

        // Parse wait(milliseconds)
        var waitMatch = Regex.Match(line, @"wait\((\d+)\)", RegexOptions.IgnoreCase);
        if (waitMatch.Success)
        {
            command.Type = CommandType.Wait;
            command.DelayMs = int.Parse(waitMatch.Groups[1].Value);
            return command;
        }

        // Parse window.open(path)
        var windowOpenMatch = Regex.Match(line, @"window\.open\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (windowOpenMatch.Success)
        {
            command.Type = CommandType.WindowOpen;
            command.ProcessPath = windowOpenMatch.Groups[1].Value.Trim().Trim('"', '\'');
            return command;
        }

        // Parse window.close(title)
        var windowCloseMatch = Regex.Match(line, @"window\.close\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (windowCloseMatch.Success)
        {
            command.Type = CommandType.WindowClose;
            command.WindowTitle = windowCloseMatch.Groups[1].Value.Trim().Trim('"', '\'');
            return command;
        }

        // Parse window.minimize(title)
        var windowMinimizeMatch = Regex.Match(line, @"window\.minimize\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (windowMinimizeMatch.Success)
        {
            command.Type = CommandType.WindowMinimize;
            command.WindowTitle = windowMinimizeMatch.Groups[1].Value.Trim().Trim('"', '\'');
            return command;
        }

        // Parse window.maximize(title)
        var windowMaximizeMatch = Regex.Match(line, @"window\.maximize\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (windowMaximizeMatch.Success)
        {
            command.Type = CommandType.WindowMaximize;
            command.WindowTitle = windowMaximizeMatch.Groups[1].Value.Trim().Trim('"', '\'');
            return command;
        }

        // Parse cmd.run(command)
        var cmdRunMatch = Regex.Match(line, @"cmd\.run\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (cmdRunMatch.Success)
        {
            command.Type = CommandType.CmdRun;
            command.ShellCommand = cmdRunMatch.Groups[1].Value.Trim().Trim('"', '\'');
            return command;
        }

        // Parse ps.run(command)
        var psRunMatch = Regex.Match(line, @"ps\.run\(([^)]+)\)", RegexOptions.IgnoreCase);
        if (psRunMatch.Success)
        {
            command.Type = CommandType.PsRun;
            command.ShellCommand = psRunMatch.Groups[1].Value.Trim().Trim('"', '\'');
            return command;
        }

        throw new Exception($"Unknown command format: {line}");
    }

    public static void SaveMacro(MacroScript script, string filePath)
    {
        var lines = new List<string>();
        lines.Add("// Macro script generated by MacroApp");
        lines.Add($"// Created: {DateTime.Now}");
        lines.Add("");

        foreach (var cmd in script.Commands)
        {
            lines.Add(GenerateLine(cmd));
        }

        File.WriteAllLines(filePath, lines);
    }

    private static string GenerateLine(MacroCommand command)
    {
        return command.Type switch
        {
            CommandType.MouseClick => $"mouse.click({command.Button})",
            CommandType.MouseMove => $"mouse.move({command.X}, {command.Y})",
            CommandType.MouseGlide => $"mouse.glide({command.X}, {command.Y}, {command.ToX}, {command.ToY}, {command.GlideDurationMs})",
            CommandType.MouseHold => $"mouse.hold({command.Button})",
            CommandType.MouseRelease => $"mouse.release({command.Button})",
            CommandType.MouseScroll => $"mouse.scroll({command.ScrollAmount})",
            CommandType.KeyboardKey => $"keyboard.key({command.Key})",
            CommandType.KeyboardButton => $"keyboard.button({command.SpecialKey})",
            CommandType.KeyboardToggle => $"keyboard.toggle({command.SpecialKey})",
            CommandType.KeyboardUntoggle => $"keyboard.untoggle({command.SpecialKey})",
            CommandType.Wait => $"wait({command.DelayMs})",
            CommandType.WindowOpen => $"window.open(\"{command.ProcessPath}\")",
            CommandType.WindowClose => $"window.close(\"{command.WindowTitle}\")",
            CommandType.WindowMinimize => $"window.minimize(\"{command.WindowTitle}\")",
            CommandType.WindowMaximize => $"window.maximize(\"{command.WindowTitle}\")",
            CommandType.CmdRun => $"cmd.run(\"{command.ShellCommand}\")",
            CommandType.PsRun => $"ps.run(\"{command.ShellCommand}\")",
            _ => ""
        };
    }
}
