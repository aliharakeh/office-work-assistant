using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.CleanText
{
    public enum TextCase
    {
        Keep,
        Upper,
        Lower,
        Proper,
        Sentence
    }

    // Every ticked option is applied, in the order below, to the text cells of the chosen columns.
    public sealed class CleanTextSettings
    {
        // Excel letters; empty = every column.
        public List<string> Columns { get; set; }
        public bool Trim { get; set; }
        public bool CollapseSpaces { get; set; }
        public bool RemoveLineBreaks { get; set; }
        // Control characters and zero-width marks go; a non-breaking space becomes a space.
        public bool RemoveNonPrinting { get; set; }
        public TextCase Case { get; set; }
        public bool TextToNumber { get; set; }
        public bool EmptyToBlank { get; set; }

        public CleanTextSettings()
        {
            Columns = new List<string>();
            Trim = true;
            CollapseSpaces = true;
            RemoveNonPrinting = true;
        }

        public bool DoesAnything
        {
            get
            {
                return Trim || CollapseSpaces || RemoveLineBreaks || RemoveNonPrinting || Case != TextCase.Keep ||
                    TextToNumber || EmptyToBlank;
            }
        }

        public string Describe()
        {
            var parts = new List<string>();
            if (Trim) parts.Add("trim");
            if (CollapseSpaces) parts.Add("single spaces");
            if (RemoveLineBreaks) parts.Add("no line breaks");
            if (RemoveNonPrinting) parts.Add("no hidden characters");
            if (Case != TextCase.Keep) parts.Add(Case.ToString().ToLowerInvariant() + " case");
            if (TextToNumber) parts.Add("text to number");
            if (EmptyToBlank) parts.Add("empty to blank");
            string cols = Columns == null || Columns.Count == 0 ? "All columns" : string.Join(", ", Columns.ToArray());
            return cols + ": " + (parts.Count == 0 ? "nothing" : string.Join(", ", parts.ToArray()));
        }
    }

    // Pure logic: tidy text cells. Numbers and dates are left as they are. No WPF.
    public static class CleanTextWork
    {
        public static DataTable Run(DataTable input, CleanTextSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null || !s.DoesAnything)
                throw new InvalidOperationException("Tick at least one cleaning option.");

            var cols = ExcelColumns.Indexes(input, s.Columns);
            if (cols.Count == 0)
            {
                for (var i = 0; i < input.Columns.Count; i++)
                    cols.Add(i);
            }
            var clean = new HashSet<int>(cols);

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            foreach (DataColumn col in input.Columns)
                ExcelFile.AddColumn(output, ExcelFile.Header(col));
            int width = input.Columns.Count;
            foreach (DataRow row in input.Rows)
            {
                var dest = output.NewRow();
                for (var c = 0; c < width; c++)
                {
                    var text = row[c] as string;
                    dest[c] = text != null && clean.Contains(c) ? Clean(text, s) : row[c];
                }
                output.Rows.Add(dest);
            }
            return output;
        }

        // A string, a number (TextToNumber) or DBNull (EmptyToBlank).
        public static object Clean(string text, CleanTextSettings s)
        {
            string t = text;
            if (s.RemoveNonPrinting)
                t = StripHidden(t);
            if (s.RemoveLineBreaks)
                t = t.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ');
            if (s.CollapseSpaces)
                t = Collapse(t);
            if (s.Trim)
                t = t.Trim();
            switch (s.Case)
            {
                case TextCase.Upper: t = t.ToUpperInvariant(); break;
                case TextCase.Lower: t = t.ToLowerInvariant(); break;
                case TextCase.Proper: t = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(t.ToLowerInvariant()); break;
                case TextCase.Sentence: t = Sentence(t); break;
            }
            if (s.TextToNumber)
            {
                double d;
                string n = t.Trim();
                if (n.Length > 0 && double.TryParse(n, NumberStyles.Number, CultureInfo.InvariantCulture, out d))
                    return d;
            }
            if (s.EmptyToBlank && t.Trim().Length == 0)
                return DBNull.Value;
            return t;
        }

        static string StripHidden(string s)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (c == ' ' || c == ' ' || c == ' ')
                    sb.Append(' ');
                else if (c == '​' || c == '‌' || c == '‍' || c == '⁠' || c == '﻿')
                    continue;
                else if (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t')
                    continue;
                else
                    sb.Append(c);
            }
            return sb.ToString();
        }

        // Runs of spaces and tabs become one space; line breaks stay.
        static string Collapse(string s)
        {
            var sb = new StringBuilder(s.Length);
            bool space = false;
            foreach (char c in s)
            {
                if (c == ' ' || c == '\t')
                {
                    if (!space)
                        sb.Append(' ');
                    space = true;
                }
                else
                {
                    sb.Append(c);
                    space = false;
                }
            }
            return sb.ToString();
        }

        // Lower case, with a capital at the start and after . ! ?
        static string Sentence(string s)
        {
            var chars = s.ToLowerInvariant().ToCharArray();
            bool start = true;
            for (var i = 0; i < chars.Length; i++)
            {
                if (start && char.IsLetter(chars[i]))
                {
                    chars[i] = char.ToUpperInvariant(chars[i]);
                    start = false;
                }
                else if (chars[i] == '.' || chars[i] == '!' || chars[i] == '?')
                    start = true;
                else if (!char.IsWhiteSpace(chars[i]))
                    start = false;
            }
            return new string(chars);
        }
    }
}
