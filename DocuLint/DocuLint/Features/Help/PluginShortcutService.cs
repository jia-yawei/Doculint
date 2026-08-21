using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal static class PluginShortcutService
    {
        internal const uint ModAlt = 0x0001;
        internal const uint ModControl = 0x0002;
        internal const uint ModShift = 0x0004;
        internal const uint ModWindows = 0x0008;
        internal const uint ModNoRepeat = 0x4000;

        internal sealed class ShortcutDefinition
        {
            internal uint Modifiers { get; set; }

            internal uint VirtualKey { get; set; }
        }

        internal static bool TryParse(string text, out ShortcutDefinition definition)
        {
            definition = null;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            string[] tokens = text.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim())
                .Where(token => token.Length > 0)
                .ToArray();
            if (tokens.Length == 0)
            {
                return false;
            }

            uint modifiers = 0;
            string keyToken = null;
            foreach (string token in tokens)
            {
                switch (token.ToLowerInvariant())
                {
                    case "ctrl":
                    case "control":
                        modifiers |= ModControl;
                        break;
                    case "alt":
                        modifiers |= ModAlt;
                        break;
                    case "shift":
                        modifiers |= ModShift;
                        break;
                    case "win":
                    case "windows":
                        modifiers |= ModWindows;
                        break;
                    default:
                        if (keyToken != null)
                        {
                            return false;
                        }

                        keyToken = token;
                        break;
                }
            }

            if (string.IsNullOrWhiteSpace(keyToken) || modifiers == 0 || !TryParseKey(keyToken, out Keys key))
            {
                return false;
            }

            definition = new ShortcutDefinition
            {
                Modifiers = modifiers,
                VirtualKey = (uint)key
            };
            return true;
        }

        internal static string Format(KeyEventArgs args)
        {
            if (args == null || IsModifierKey(args.KeyCode))
            {
                return string.Empty;
            }

            List<string> parts = new List<string>();
            if (args.Control)
            {
                parts.Add("Ctrl");
            }

            if (args.Alt)
            {
                parts.Add("Alt");
            }

            if (args.Shift)
            {
                parts.Add("Shift");
            }

            if (args.KeyCode == Keys.LWin || args.KeyCode == Keys.RWin)
            {
                parts.Add("Win");
            }

            if (parts.Count == 0)
            {
                return string.Empty;
            }

            parts.Add(FormatKey(args.KeyCode));
            return string.Join("+", parts);
        }

        internal static string Normalize(string text)
        {
            return TryParse(text, out ShortcutDefinition definition)
                ? Format(definition)
                : string.Empty;
        }

        private static string Format(ShortcutDefinition definition)
        {
            List<string> parts = new List<string>();
            if ((definition.Modifiers & ModControl) != 0)
            {
                parts.Add("Ctrl");
            }

            if ((definition.Modifiers & ModAlt) != 0)
            {
                parts.Add("Alt");
            }

            if ((definition.Modifiers & ModShift) != 0)
            {
                parts.Add("Shift");
            }

            if ((definition.Modifiers & ModWindows) != 0)
            {
                parts.Add("Win");
            }

            parts.Add(FormatKey((Keys)definition.VirtualKey));
            return string.Join("+", parts);
        }

        private static bool TryParseKey(string token, out Keys key)
        {
            key = Keys.None;
            if (string.Equals(token, "Space", StringComparison.OrdinalIgnoreCase))
            {
                key = Keys.Space;
                return true;
            }

            if (string.Equals(token, "Enter", StringComparison.OrdinalIgnoreCase))
            {
                key = Keys.Enter;
                return true;
            }

            if (string.Equals(token, "Esc", StringComparison.OrdinalIgnoreCase))
            {
                key = Keys.Escape;
                return true;
            }

            if (token.Length == 1 && char.IsLetterOrDigit(token[0]))
            {
                char character = char.ToUpperInvariant(token[0]);
                key = character >= 'A' && character <= 'Z'
                    ? (Keys)(character - 'A' + (int)Keys.A)
                    : character >= '0' && character <= '9'
                        ? (Keys)(character - '0' + (int)Keys.D0)
                        : Keys.None;
                return key != Keys.None;
            }

            if (Enum.TryParse(token, true, out Keys parsed) && !IsModifierKey(parsed))
            {
                key = parsed;
                return key != Keys.None;
            }

            return false;
        }

        private static bool IsModifierKey(Keys key)
        {
            return key == Keys.ControlKey
                || key == Keys.Menu
                || key == Keys.ShiftKey
                || key == Keys.LWin
                || key == Keys.RWin;
        }

        private static string FormatKey(Keys key)
        {
            if (key == Keys.Space)
            {
                return "Space";
            }

            if (key == Keys.Enter)
            {
                return "Enter";
            }

            if (key == Keys.Escape)
            {
                return "Esc";
            }

            string name = key.ToString();
            if (name.Length == 2 && name[0] == 'D' && char.IsDigit(name[1]))
            {
                return name.Substring(1);
            }

            return name;
        }
    }
}
