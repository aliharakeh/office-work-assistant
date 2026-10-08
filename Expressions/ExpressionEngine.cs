using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeWorkAssistant.Expressions
{
    public sealed class VariableHelp
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Example { get; set; }
    }

    public sealed class FunctionHelp
    {
        public string Signature { get; set; }
        public string Description { get; set; }
        public string Example { get; set; }
    }

    // Generic $Name + formula DSL. No DataTable, no WPF, no ClosedXML.
    // Lookup resolves a name: columns, built-ins, or anything else.
    // A $letter that fits the caller wins; otherwise a built-in wins.
    // Examples: $A * $B | ($A + $B) / 2 | $A & " " & $B
    // | $A ^ 2 | $A > 10 && $B == "OK" | IF($A > 10, "Big", "Small")
    // | $Today | ADDDAYS($Today, 7) | FORMAT($Today, "dd/MM/yyyy")
    // | FIRSTWORD($Name) | LASTWORD($Name) | split parts: $1 $2, $A1 $B1
    public static class ExpressionEngine
    {
        static readonly Regex Placeholder = new Regex(@"\$(\w+)", RegexOptions.Compiled);

        static readonly Dictionary<string, Func<object>> Builtins =
            new Dictionary<string, Func<object>>(StringComparer.OrdinalIgnoreCase);

        static ExpressionEngine()
        {
            Builtins["Today"] = delegate { return DateTime.Today; };
            Builtins["Yesterday"] = delegate { return DateTime.Today.AddDays(-1); };
            Builtins["Tomorrow"] = delegate { return DateTime.Today.AddDays(1); };
            Builtins["DayAfterTomorrow"] = delegate { return DateTime.Today.AddDays(2); };
            Builtins["WeekAgo"] = delegate { return DateTime.Today.AddDays(-7); };
            Builtins["WeekLater"] = delegate { return DateTime.Today.AddDays(7); };
            Builtins["MonthAgo"] = delegate { return DateTime.Today.AddMonths(-1); };
            Builtins["MonthLater"] = delegate { return DateTime.Today.AddMonths(1); };
            Builtins["StartOfWeek"] = delegate { return StartOfWeek(DateTime.Today); };
            Builtins["EndOfWeek"] = delegate { return EndOfWeek(DateTime.Today); };
            Builtins["StartOfNextWeek"] = delegate { return StartOfWeek(DateTime.Today).AddDays(7); };
            Builtins["EndOfNextWeek"] = delegate { return EndOfWeek(DateTime.Today).AddDays(7); };
            Builtins["StartOfPrevWeek"] = delegate { return StartOfWeek(DateTime.Today).AddDays(-7); };
            Builtins["EndOfPrevWeek"] = delegate { return EndOfWeek(DateTime.Today).AddDays(-7); };
            Builtins["StartOfMonth"] = delegate { return StartOfMonth(DateTime.Today); };
            Builtins["EndOfMonth"] = delegate { return EndOfMonth(DateTime.Today); };
            Builtins["StartOfNextMonth"] = delegate { return StartOfMonth(DateTime.Today).AddMonths(1); };
            Builtins["EndOfNextMonth"] = delegate { return EndOfMonth(StartOfMonth(DateTime.Today).AddMonths(1)); };
            Builtins["StartOfPrevMonth"] = delegate { return StartOfMonth(DateTime.Today).AddMonths(-1); };
            Builtins["EndOfPrevMonth"] = delegate { return EndOfMonth(StartOfMonth(DateTime.Today).AddMonths(-1)); };
            Builtins["StartOfYear"] = delegate { return new DateTime(DateTime.Today.Year, 1, 1); };
            Builtins["EndOfYear"] = delegate { return new DateTime(DateTime.Today.Year, 12, 31); };
            Builtins["CurrentYear"] = delegate { return (object)DateTime.Today.Year; };
            Builtins["CurrentMonth"] = delegate { return (object)DateTime.Today.Month; };
            Builtins["CurrentDay"] = delegate { return (object)DateTime.Today.Day; };
            Builtins["Now"] = delegate { return DateTime.Now; };
        }

        public static bool TryGetVariable(string name, out object value)
        {
            value = null;
            if (string.IsNullOrWhiteSpace(name))
                return false;
            string key = name.Trim();
            Func<object> provider;
            if (!Builtins.TryGetValue(key, out provider))
                return false;
            value = provider();
            return true;
        }

        public static List<VariableHelp> GetVariableHelp()
        {
            var rows = new List<VariableHelp>();
            AddVar(rows, "Today", "Today's date, no time.");
            AddVar(rows, "Yesterday", "Today minus 1 day.");
            AddVar(rows, "Tomorrow", "Today plus 1 day.");
            AddVar(rows, "DayAfterTomorrow", "Today plus 2 days.");
            AddVar(rows, "WeekAgo", "Today minus 7 days.");
            AddVar(rows, "WeekLater", "Today plus 7 days.");
            AddVar(rows, "MonthAgo", "Same day last month.");
            AddVar(rows, "MonthLater", "Same day next month.");
            AddVar(rows, "StartOfWeek", "Monday of this week.");
            AddVar(rows, "EndOfWeek", "Sunday of this week.");
            AddVar(rows, "StartOfNextWeek", "Monday of next week.");
            AddVar(rows, "EndOfNextWeek", "Sunday of next week.");
            AddVar(rows, "StartOfPrevWeek", "Monday of last week.");
            AddVar(rows, "EndOfPrevWeek", "Sunday of last week.");
            AddVar(rows, "StartOfMonth", "First day of this month.");
            AddVar(rows, "EndOfMonth", "Last day of this month.");
            AddVar(rows, "StartOfNextMonth", "First day of next month.");
            AddVar(rows, "EndOfNextMonth", "Last day of next month.");
            AddVar(rows, "StartOfPrevMonth", "First day of last month.");
            AddVar(rows, "EndOfPrevMonth", "Last day of last month.");
            AddVar(rows, "StartOfYear", "Jan 1st of this year.");
            AddVar(rows, "EndOfYear", "Dec 31st of this year.");
            AddVar(rows, "CurrentYear", "Year number, e.g. 2026.");
            AddVar(rows, "CurrentMonth", "Month number 1-12.");
            AddVar(rows, "CurrentDay", "Day of month 1-31.");
            AddVar(rows, "Now", "Current date plus time.");
            return rows;
        }

        static void AddVar(List<VariableHelp> rows, string name, string description)
        {
            object value;
            string example = "";
            if (TryGetVariable(name, out value))
                example = FormatScalar(value);
            rows.Add(new VariableHelp { Name = "$" + name, Description = description, Example = example });
        }

        public static List<FunctionHelp> GetFunctionHelp()
        {
            return new List<FunctionHelp>
            {
                new FunctionHelp { Signature = "TODAY()", Description = "Today's date, no time.", Example = "TODAY()" },
                new FunctionHelp { Signature = "NOW()", Description = "Current date plus time.", Example = "NOW()" },
                new FunctionHelp { Signature = "DATE(year, month, day)", Description = "Build a date.", Example = "DATE(2026, 1, 31)" },
                new FunctionHelp { Signature = "YEAR(d) MONTH(d) DAY(d)", Description = "Parts of a date.", Example = "YEAR($Today)" },
                new FunctionHelp { Signature = "WEEKDAY(d)", Description = "1=Monday .. 7=Sunday.", Example = "WEEKDAY($Today)" },
                new FunctionHelp { Signature = "ADDDAYS(d, n)", Description = "Add n days (n may be negative).", Example = "ADDDAYS($Today, 7)" },
                new FunctionHelp { Signature = "ADDWEEKS(d, n)", Description = "Add n weeks.", Example = "ADDWEEKS($Today, 1)" },
                new FunctionHelp { Signature = "ADDMONTHS(d, n)", Description = "Add n months.", Example = "ADDMONTHS($Today, 1)" },
                new FunctionHelp { Signature = "ADDYEARS(d, n)", Description = "Add n years.", Example = "ADDYEARS($Today, 1)" },
                new FunctionHelp { Signature = "STARTOFWEEK(d) ENDOFWEEK(d)", Description = "Monday / Sunday of that week.", Example = "STARTOFWEEK($Today)" },
                new FunctionHelp { Signature = "STARTOFMONTH(d) ENDOFMONTH(d)", Description = "First / last day of that month.", Example = "ENDOFMONTH($Today)" },
                new FunctionHelp { Signature = "FORMAT(d, fmt)", Description = ".NET date format to text.", Example = "FORMAT($Today, \"dd/MM/yyyy\")" },
                new FunctionHelp { Signature = "REMOVEDIGITS(text)", Description = "Strip 0-9 from text.", Example = "REMOVEDIGITS($A)" },
                new FunctionHelp { Signature = "CLEARSYMBOLS(text)", Description = "Keep letters, digits, and spaces only.", Example = "CLEARSYMBOLS($A)" },
                new FunctionHelp { Signature = "TRIM(text)", Description = "Strip leading and trailing spaces.", Example = "TRIM($A)" },
                new FunctionHelp { Signature = "FIRSTWORD(text)", Description = "First word of the text.", Example = "FIRSTWORD($Name)" },
                new FunctionHelp { Signature = "LASTWORD(text)", Description = "Last word of the text.", Example = "LASTWORD($Name)" },
                new FunctionHelp { Signature = "CONTAINS(text, needle)", Description = "True if text contains needle (ignore case).", Example = "CONTAINS($Name, \"report\")" },
                new FunctionHelp { Signature = "STARTSWITH(text, prefix)", Description = "True if text starts with prefix (ignore case).", Example = "STARTSWITH($Name, \"INV\")" },
                new FunctionHelp { Signature = "ENDSWITH(text, suffix)", Description = "True if text ends with suffix (ignore case).", Example = "ENDSWITH($Name, \".pdf\")" },
            };
        }

        static DateTime StartOfWeek(DateTime d)
        {
            DateTime date = d.Date;
            int diff = (((int)date.DayOfWeek) + 6) % 7;
            return date.AddDays(-diff);
        }

        static DateTime EndOfWeek(DateTime d)
        {
            return StartOfWeek(d).AddDays(6);
        }

        static DateTime StartOfMonth(DateTime d)
        {
            return new DateTime(d.Year, d.Month, 1);
        }

        static DateTime EndOfMonth(DateTime d)
        {
            return StartOfMonth(d).AddMonths(1).AddDays(-1);
        }

        // Fixed text: only $BuiltIn names expand, unknown stays as-is.
        public static string ExpandVariables(string text)
        {
            if (text == null)
                return "";
            return Placeholder.Replace(text, delegate(Match m)
            {
                string key = m.Groups[1].Value != null ? m.Groups[1].Value.Trim() : "";
                object v;
                if (TryGetVariable(key, out v))
                    return FormatScalar(v);
                return m.Value;
            });
        }

        // Pattern text: every $name goes through lookup, missing is "".
        public static string ExpandPlaceholders(string pattern, Func<string, object> lookup)
        {
            if (pattern == null)
                return "";
            return Placeholder.Replace(pattern, delegate(Match m)
            {
                string key = m.Groups[1].Value != null ? m.Groups[1].Value.Trim() : "";
                object v = null;
                if (lookup != null)
                {
                    try { v = lookup(key); }
                    catch { v = null; }
                }
                if (v == null || v == DBNull.Value)
                    return "";
                return FormatScalar(v);
            });
        }

        public static string FormatScalar(object value)
        {
            if (value == null || value == DBNull.Value)
                return "";
            if (value is DateTime)
            {
                DateTime dt = (DateTime)value;
                if (dt.TimeOfDay == TimeSpan.Zero)
                    return dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                return dt.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        static bool TryParseDateString(string s, out DateTime result)
        {
            result = default(DateTime);
            if (string.IsNullOrEmpty(s))
                return false;
            string[] formats = new string[]
            {
                "dd/MM/yyyy HH:mm:ss", "d/M/yyyy HH:mm:ss",
                "dd/MM/yyyy", "d/M/yyyy",
                "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd",
            };
            if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return true;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out result))
                return true;
            return DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
        }

        public static bool TryToDate(object value, out DateTime result)
        {
            result = default(DateTime);
            if (value == null || value == DBNull.Value)
                return false;
            if (value is DateTime)
            {
                result = (DateTime)value;
                return true;
            }
            if (value is double)
            {
                try
                {
                    result = DateTime.FromOADate((double)value);
                    return true;
                }
                catch { return false; }
            }
            if (value is float || value is int || value is long || value is decimal)
            {
                try
                {
                    result = DateTime.FromOADate(Convert.ToDouble(value, CultureInfo.InvariantCulture));
                    return true;
                }
                catch { return false; }
            }
            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (s != null)
                s = s.Trim();
            if (string.IsNullOrEmpty(s))
                return false;
            return TryParseDateString(s, out result);
        }

        public static bool IsDateLike(object value)
        {
            if (value == null || value == DBNull.Value)
                return false;
            if (value is DateTime)
                return true;
            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (s != null)
                s = s.Trim();
            if (string.IsNullOrEmpty(s))
                return false;
            if (s.IndexOf('-') < 0 && s.IndexOf('/') < 0)
                return false;
            DateTime dt;
            return TryParseDateString(s, out dt);
        }

        // One formula box that mixes $names, numbers, quoted text
        // and + - * / ^ with parentheses. Pure math returns a number;
        // any text part (or & operator) switches to a string.
        // Comparisons (== != > >= < <=), logic (! && ||) and IF(cond, a, b)
        // return true/false or pick a branch.
        // Date helpers: TODAY() NOW() DATE(y,m,d) YEAR(d) MONTH(d) DAY(d)
        // WEEKDAY(d) ADDDAYS(d,n) ADDWEEKS(d,n) ADDMONTHS(d,n) ADDYEARS(d,n)
        // STARTOFWEEK(d) ENDOFWEEK(d) STARTOFMONTH(d) ENDOFMONTH(d)
        // FORMAT(d, "dd/MM/yyyy"). Dates also support + / - days, e.g.
        // $Today + 1, and date comparisons, e.g. $Due > $Today.
        // Text helpers: TRIM REMOVEDIGITS CLEARSYMBOLS FIRSTWORD LASTWORD.
        public static string Validate(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return "enter a formula like $A * $B.";
            try
            {
                var parser = new ExpressionParser(expression);
                parser.Parse();
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static object Evaluate(string expression, Func<string, object> lookup)
        {
            var parser = new ExpressionParser(expression);
            ExprNode root = parser.Parse();
            return root.Evaluate(lookup);
        }

        // Parses once; the result evaluates many times. For a formula run per file or per row pair.
        public static Func<Func<string, object>, object> Compile(string expression)
        {
            var parser = new ExpressionParser(expression);
            ExprNode root = parser.Parse();
            return root.Evaluate;
        }

        public static bool TryToNumber(object value, out double result)
        {
            result = 0;
            if (value == null || value == DBNull.Value)
                return false;
            if (value is double)
            {
                result = (double)value;
                return true;
            }
            if (value is float)
            {
                result = (float)value;
                return true;
            }
            if (value is int)
            {
                result = (int)value;
                return true;
            }
            if (value is long)
            {
                result = (long)value;
                return true;
            }
            if (value is decimal)
            {
                result = (double)(decimal)value;
                return true;
            }
            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (s != null)
                s = s.Trim();
            if (string.IsNullOrEmpty(s))
                return false;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                return true;
            return double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out result);
        }

        public static string ToText(object value)
        {
            return FormatScalar(value);
        }

        public static bool ToBool(object value)
        {
            if (value == null || value == DBNull.Value)
                return false;
            if (value is bool)
                return (bool)value;
            double d;
            if (TryToNumber(value, out d))
                return d != 0;
            string s = ToText(value).Trim();
            if (s.Length == 0)
                return false;
            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase))
                return false;
            return true;
        }

        public static bool Compare(string op, object a, object b)
        {
            double x;
            double y;
            if (TryToNumber(a, out x) && TryToNumber(b, out y))
            {
                if (op == "==") return x == y;
                if (op == "!=") return x != y;
                if (op == ">") return x > y;
                if (op == "<") return x < y;
                if (op == ">=") return x >= y;
                return x <= y;
            }
            DateTime da;
            DateTime db;
            if ((IsDateLike(a) || IsDateLike(b)) && TryToDate(a, out da) && TryToDate(b, out db))
            {
                int cmpDate = DateTime.Compare(da, db);
                if (op == "==") return cmpDate == 0;
                if (op == "!=") return cmpDate != 0;
                if (op == ">") return cmpDate > 0;
                if (op == "<") return cmpDate < 0;
                if (op == ">=") return cmpDate >= 0;
                return cmpDate <= 0;
            }
            string sa = ToText(a);
            string sb = ToText(b);
            if (op == "==") return string.Equals(sa, sb, StringComparison.OrdinalIgnoreCase);
            if (op == "!=") return !string.Equals(sa, sb, StringComparison.OrdinalIgnoreCase);
            int cmp = string.Compare(sa, sb, StringComparison.OrdinalIgnoreCase);
            if (op == ">") return cmp > 0;
            if (op == "<") return cmp < 0;
            if (op == ">=") return cmp >= 0;
            return cmp <= 0;
        }

        static object EvaluateFunction(string name, object[] args)
        {
            string fn = name != null ? name.Trim().ToUpperInvariant() : "";
            if (fn == "TODAY")
            {
                if (args.Length != 0)
                    throw new InvalidOperationException("TODAY() takes no arguments.");
                return DateTime.Today;
            }
            if (fn == "NOW")
            {
                if (args.Length != 0)
                    throw new InvalidOperationException("NOW() takes no arguments.");
                return DateTime.Now;
            }
            if (fn == "DATE")
            {
                if (args.Length != 3)
                    throw new InvalidOperationException("DATE needs (year, month, day).");
                double y;
                double m;
                double d;
                if (!TryToNumber(args[0], out y) || !TryToNumber(args[1], out m) || !TryToNumber(args[2], out d))
                    throw new InvalidOperationException("DATE needs (year, month, day) numbers.");
                try { return new DateTime((int)y, (int)m, (int)d); }
                catch { throw new InvalidOperationException("DATE got an invalid year/month/day."); }
            }
            if (fn == "YEAR" || fn == "MONTH" || fn == "DAY" || fn == "WEEKDAY")
            {
                if (args.Length != 1)
                    throw new InvalidOperationException(fn + " needs (date).");
                DateTime dt;
                if (!TryToDate(args[0], out dt))
                    throw new InvalidOperationException(fn + " needs a date.");
                if (fn == "YEAR") return (double)dt.Year;
                if (fn == "MONTH") return (double)dt.Month;
                if (fn == "DAY") return (double)dt.Day;
                return (double)((((int)dt.DayOfWeek) + 6) % 7 + 1);
            }
            if (fn == "ADDDAYS" || fn == "ADDWEEKS" || fn == "ADDMONTHS" || fn == "ADDYEARS")
            {
                if (args.Length != 2)
                    throw new InvalidOperationException(fn + " needs (date, n).");
                DateTime dt;
                double n;
                if (!TryToDate(args[0], out dt))
                    throw new InvalidOperationException(fn + " needs a date first.");
                if (!TryToNumber(args[1], out n))
                    throw new InvalidOperationException(fn + " needs a number second.");
                if (fn == "ADDDAYS") return dt.AddDays(n);
                if (fn == "ADDWEEKS") return dt.AddDays(n * 7);
                if (fn == "ADDMONTHS") return dt.AddMonths((int)n);
                return dt.AddYears((int)n);
            }
            if (fn == "STARTOFWEEK" || fn == "ENDOFWEEK" || fn == "STARTOFMONTH" || fn == "ENDOFMONTH")
            {
                if (args.Length != 1)
                    throw new InvalidOperationException(fn + " needs (date).");
                DateTime dt;
                if (!TryToDate(args[0], out dt))
                    throw new InvalidOperationException(fn + " needs a date.");
                if (fn == "STARTOFWEEK") return StartOfWeek(dt);
                if (fn == "ENDOFWEEK") return EndOfWeek(dt);
                if (fn == "STARTOFMONTH") return StartOfMonth(dt);
                return EndOfMonth(dt);
            }
            if (fn == "FORMAT")
            {
                if (args.Length != 2)
                    throw new InvalidOperationException("FORMAT needs (date, \"fmt\").");
                string fmt = ToText(args[1]);
                if (string.IsNullOrEmpty(fmt))
                    throw new InvalidOperationException("FORMAT needs (date, \"fmt\").");
                DateTime dt;
                if (TryToDate(args[0], out dt))
                {
                    try { return dt.ToString(fmt, CultureInfo.InvariantCulture); }
                    catch { throw new InvalidOperationException("FORMAT got a bad pattern."); }
                }
                double num;
                if (TryToNumber(args[0], out num))
                {
                    try { return num.ToString(fmt, CultureInfo.InvariantCulture); }
                    catch { throw new InvalidOperationException("FORMAT got a bad pattern."); }
                }
                return ToText(args[0]);
            }
            if (fn == "REMOVEDIGITS" || fn == "CLEARSYMBOLS")
            {
                if (args.Length != 1)
                    throw new InvalidOperationException(fn + " needs (text).");
                string s = ToText(args[0]);
                var sb = new StringBuilder();
                for (var i = 0; i < s.Length; i++)
                {
                    char c = s[i];
                    if (fn == "REMOVEDIGITS")
                    {
                        if (!char.IsDigit(c))
                            sb.Append(c);
                    }
                    else if (char.IsLetterOrDigit(c) || char.IsWhiteSpace(c))
                        sb.Append(c);
                }
                return sb.ToString();
            }
            if (fn == "TRIM")
            {
                if (args.Length != 1)
                    throw new InvalidOperationException("TRIM needs (text).");
                return ToText(args[0]).Trim();
            }
            if (fn == "FIRSTWORD" || fn == "LASTWORD")
            {
                if (args.Length != 1)
                    throw new InvalidOperationException(fn + " needs (text).");
                string text = ToText(args[0]).Trim();
                if (text.Length == 0)
                    return "";
                var words = text.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length == 0)
                    return "";
                return fn == "FIRSTWORD" ? words[0] : words[words.Length - 1];
            }
            if (fn == "CONTAINS" || fn == "STARTSWITH" || fn == "ENDSWITH")
            {
                if (args.Length != 2)
                    throw new InvalidOperationException(fn + " needs (text, needle).");
                string hay = ToText(args[0]);
                string needle = ToText(args[1]);
                if (fn == "CONTAINS")
                    return hay.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
                if (fn == "STARTSWITH")
                    return hay.StartsWith(needle, StringComparison.OrdinalIgnoreCase);
                return hay.EndsWith(needle, StringComparison.OrdinalIgnoreCase);
            }
            throw new InvalidOperationException("Unknown function '" + name + "'.");
        }

        static void CheckArity(string name, int count)
        {
            string fn = name != null ? name.Trim().ToUpperInvariant() : "";
            string need = null;
            if (fn == "TODAY" || fn == "NOW")
            {
                if (count != 0) need = "takes no arguments";
            }
            else if (fn == "DATE")
            {
                if (count != 3) need = "needs (year, month, day)";
            }
            else if (fn == "YEAR" || fn == "MONTH" || fn == "DAY" || fn == "WEEKDAY" ||
                fn == "STARTOFWEEK" || fn == "ENDOFWEEK" || fn == "STARTOFMONTH" || fn == "ENDOFMONTH")
            {
                if (count != 1) need = "needs (date)";
            }
            else if (fn == "REMOVEDIGITS" || fn == "CLEARSYMBOLS" || fn == "TRIM" || fn == "FIRSTWORD" || fn == "LASTWORD")
            {
                if (count != 1) need = "needs (text)";
            }
            else if (fn == "CONTAINS" || fn == "STARTSWITH" || fn == "ENDSWITH")
            {
                if (count != 2) need = "needs (text, needle)";
            }
            else if (fn == "ADDDAYS" || fn == "ADDWEEKS" || fn == "ADDMONTHS" || fn == "ADDYEARS" || fn == "FORMAT")
            {
                if (count != 2) need = fn == "FORMAT" ? "needs (date, \"fmt\")" : "needs (date, n)";
            }
            if (need != null)
                throw new InvalidOperationException(fn + " " + need + ".");
        }

        abstract class ExprNode
        {
            public abstract object Evaluate(Func<string, object> lookup);
        }

        sealed class NumberNode : ExprNode
        {
            readonly double _value;
            public NumberNode(double value) { _value = value; }
            public override object Evaluate(Func<string, object> lookup) { return _value; }
        }

        sealed class TextNode : ExprNode
        {
            readonly string _text;
            public TextNode(string text) { _text = text != null ? text : ""; }
            public override object Evaluate(Func<string, object> lookup) { return _text; }
        }

        sealed class ColumnNode : ExprNode
        {
            readonly string _name;
            public ColumnNode(string name) { _name = name != null ? name.Trim() : ""; }
            public override object Evaluate(Func<string, object> lookup)
            {
                if (lookup == null)
                    return null;
                return lookup(_name);
            }
        }

        sealed class UnaryNode : ExprNode
        {
            readonly char _op;
            readonly ExprNode _inner;
            public UnaryNode(char op, ExprNode inner) { _op = op; _inner = inner; }
            public override object Evaluate(Func<string, object> lookup)
            {
                object v = _inner.Evaluate(lookup);
                double d;
                if (!TryToNumber(v, out d))
                    return DBNull.Value;
                return _op == '-' ? -d : d;
            }
        }

        sealed class BinaryNode : ExprNode
        {
            readonly char _op;
            readonly ExprNode _left;
            readonly ExprNode _right;
            public BinaryNode(char op, ExprNode left, ExprNode right)
            {
                _op = op;
                _left = left;
                _right = right;
            }
            public override object Evaluate(Func<string, object> lookup)
            {
                object a = _left.Evaluate(lookup);
                object b = _right.Evaluate(lookup);
                if (_op == '&')
                    return ToText(a) + ToText(b);

                if (_op == '+' || _op == '-')
                {
                    DateTime da;
                    double n;
                    if (IsDateLike(a) && TryToDate(a, out da) && TryToNumber(b, out n))
                    {
                        if (_op == '+')
                            return da.AddDays(n);
                        return da.AddDays(-n);
                    }
                    DateTime db;
                    if (_op == '-' && IsDateLike(a) && IsDateLike(b) &&
                        TryToDate(a, out da) && TryToDate(b, out db))
                        return (da - db).TotalDays;
                }

                double x;
                double y;
                if (!TryToNumber(a, out x) || !TryToNumber(b, out y))
                {
                    // + also concatenates text so "Vol: " + $Qty works.
                    if (_op == '+')
                        return ToText(a) + ToText(b);
                    return DBNull.Value;
                }
                if (_op == '+')
                    return x + y;
                if (_op == '-')
                    return x - y;
                if (_op == '*')
                    return x * y;
                if (_op == '/')
                {
                    if (y == 0)
                        return DBNull.Value;
                    return x / y;
                }
                if (_op == '^')
                    return Math.Pow(x, y);
                return DBNull.Value;
            }
        }

        sealed class BoolNode : ExprNode
        {
            readonly bool _value;
            public BoolNode(bool value) { _value = value; }
            public override object Evaluate(Func<string, object> lookup) { return _value; }
        }

        sealed class ComparisonNode : ExprNode
        {
            readonly string _op;
            readonly ExprNode _left;
            readonly ExprNode _right;
            public ComparisonNode(string op, ExprNode left, ExprNode right)
            {
                _op = op;
                _left = left;
                _right = right;
            }
            public override object Evaluate(Func<string, object> lookup)
            {
                object a = _left.Evaluate(lookup);
                object b = _right.Evaluate(lookup);
                return Compare(_op, a, b);
            }
        }

        sealed class LogicalNode : ExprNode
        {
            readonly string _op;
            readonly ExprNode _left;
            readonly ExprNode _right;
            public LogicalNode(string op, ExprNode left, ExprNode right)
            {
                _op = op;
                _left = left;
                _right = right;
            }
            public override object Evaluate(Func<string, object> lookup)
            {
                if (_op == "&&")
                {
                    if (!ToBool(_left.Evaluate(lookup)))
                        return false;
                    return ToBool(_right.Evaluate(lookup));
                }
                if (!ToBool(_left.Evaluate(lookup)))
                    return ToBool(_right.Evaluate(lookup));
                return true;
            }
        }

        sealed class NotNode : ExprNode
        {
            readonly ExprNode _inner;
            public NotNode(ExprNode inner) { _inner = inner; }
            public override object Evaluate(Func<string, object> lookup)
            {
                return !ToBool(_inner.Evaluate(lookup));
            }
        }

        sealed class IfNode : ExprNode
        {
            readonly ExprNode _cond;
            readonly ExprNode _then;
            readonly ExprNode _else;
            public IfNode(ExprNode cond, ExprNode then, ExprNode otherwise)
            {
                _cond = cond;
                _then = then;
                _else = otherwise;
            }
            public override object Evaluate(Func<string, object> lookup)
            {
                if (ToBool(_cond.Evaluate(lookup)))
                    return _then.Evaluate(lookup);
                return _else.Evaluate(lookup);
            }
        }

        sealed class FunctionNode : ExprNode
        {
            readonly string _name;
            readonly List<ExprNode> _args;
            public FunctionNode(string name, List<ExprNode> args)
            {
                _name = name;
                _args = args != null ? args : new List<ExprNode>();
            }
            public override object Evaluate(Func<string, object> lookup)
            {
                var values = new object[_args.Count];
                for (var i = 0; i < _args.Count; i++)
                    values[i] = _args[i].Evaluate(lookup);
                return EvaluateFunction(_name, values);
            }
        }

        sealed class ExpressionParser
        {
            readonly string _text;
            int _pos;

            public ExpressionParser(string text)
            {
                _text = text != null ? text : "";
                _pos = 0;
            }

            public ExprNode Parse()
            {
                ExprNode node = ParseOr();
                SkipSpaces();
                if (_pos != _text.Length)
                    throw new InvalidOperationException("Unexpected '" + _text[_pos] + "' at position " + (_pos + 1) + ".");
                return node;
            }

            // Lowest precedence: || then && then comparisons, then & text glue.
            ExprNode ParseOr()
            {
                ExprNode left = ParseAnd();
                for (;;)
                {
                    SkipSpaces();
                    if (Match("||"))
                    {
                        ExprNode right = ParseAnd();
                        left = new LogicalNode("||", left, right);
                        continue;
                    }
                    break;
                }
                return left;
            }

            ExprNode ParseAnd()
            {
                ExprNode left = ParseComparison();
                for (;;)
                {
                    SkipSpaces();
                    if (Match("&&"))
                    {
                        ExprNode right = ParseComparison();
                        left = new LogicalNode("&&", left, right);
                        continue;
                    }
                    break;
                }
                return left;
            }

            ExprNode ParseComparison()
            {
                ExprNode left = ParseConcat();
                for (;;)
                {
                    SkipSpaces();
                    string op = null;
                    if (Match("==")) op = "==";
                    else if (Match("!=")) op = "!=";
                    else if (Match(">=")) op = ">=";
                    else if (Match("<=")) op = "<=";
                    else if (Match(">")) op = ">";
                    else if (Match("<")) op = "<";
                    if (op == null)
                        break;
                    ExprNode right = ParseConcat();
                    left = new ComparisonNode(op, left, right);
                }
                return left;
            }

            // Lowest precedence: & and text glue (+ with text also concatenates).
            ExprNode ParseConcat()
            {
                ExprNode left = ParseAdd();
                for (;;)
                {
                    SkipSpaces();
                    if (_pos < _text.Length && _text[_pos] == '&' && !Peek("&&"))
                    {
                        _pos++;
                        ExprNode right = ParseAdd();
                        left = new BinaryNode('&', left, right);
                        continue;
                    }
                    break;
                }
                return left;
            }

            ExprNode ParseAdd()
            {
                ExprNode left = ParseMul();
                for (;;)
                {
                    SkipSpaces();
                    if (_pos < _text.Length && (_text[_pos] == '+' || _text[_pos] == '-'))
                    {
                        char op = _text[_pos];
                        _pos++;
                        ExprNode right = ParseMul();
                        left = new BinaryNode(op, left, right);
                        continue;
                    }
                    break;
                }
                return left;
            }

            ExprNode ParseMul()
            {
                ExprNode left = ParsePower();
                for (;;)
                {
                    SkipSpaces();
                    if (_pos < _text.Length && (_text[_pos] == '*' || _text[_pos] == '/'))
                    {
                        char op = _text[_pos];
                        _pos++;
                        ExprNode right = ParsePower();
                        left = new BinaryNode(op, left, right);
                        continue;
                    }
                    break;
                }
                return left;
            }

            // ^ is right-associative: 2 ^ 3 ^ 2 = 2 ^ (3 ^ 2).
            ExprNode ParsePower()
            {
                ExprNode left = ParseUnary();
                SkipSpaces();
                if (_pos < _text.Length && _text[_pos] == '^')
                {
                    _pos++;
                    ExprNode right = ParsePower();
                    return new BinaryNode('^', left, right);
                }
                return left;
            }

            ExprNode ParseUnary()
            {
                SkipSpaces();
                if (_pos < _text.Length && (_text[_pos] == '+' || _text[_pos] == '-'))
                {
                    char op = _text[_pos];
                    _pos++;
                    return new UnaryNode(op, ParseUnary());
                }
                if (_pos < _text.Length && _text[_pos] == '!' && !Peek("!="))
                {
                    _pos++;
                    return new NotNode(ParseUnary());
                }
                return ParsePrimary();
            }

            ExprNode ParsePrimary()
            {
                SkipSpaces();
                if (_pos >= _text.Length)
                    throw new InvalidOperationException("Unexpected end of formula.");
                char c = _text[_pos];
                if (c == '(')
                {
                    _pos++;
                    ExprNode inner = ParseOr();
                    SkipSpaces();
                    if (_pos >= _text.Length || _text[_pos] != ')')
                        throw new InvalidOperationException("Missing ')'.");
                    _pos++;
                    return inner;
                }
                if (c == '{' || c == '}')
                {
                    throw new InvalidOperationException(
                        "Curly braces are no longer used. Write $Column for columns, $Today for built-ins, $1 $2 for split parts.");
                }
                if (c == '$')
                {
                    _pos++;
                    int start = _pos;
                    while (_pos < _text.Length && (char.IsLetterOrDigit(_text[_pos]) || _text[_pos] == '_'))
                        _pos++;
                    if (_pos == start)
                        throw new InvalidOperationException("Missing name after '$'. Write $Name, $A or $1.");
                    return new ColumnNode(_text.Substring(start, _pos - start));
                }
                if (c == '"' || c == '\'')
                {
                    char quote = c;
                    _pos++;
                    int start = _pos;
                    while (_pos < _text.Length && _text[_pos] != quote)
                        _pos++;
                    if (_pos >= _text.Length)
                        throw new InvalidOperationException("Missing closing quote.");
                    string text = _text.Substring(start, _pos - start);
                    _pos++;
                    return new TextNode(text);
                }
                if (char.IsDigit(c) || c == '.')
                {
                    int start = _pos;
                    bool dot = false;
                    while (_pos < _text.Length && (char.IsDigit(_text[_pos]) || _text[_pos] == '.'))
                    {
                        if (_text[_pos] == '.')
                        {
                            if (dot)
                                break;
                            dot = true;
                        }
                        _pos++;
                    }
                    string token = _text.Substring(start, _pos - start);
                    double value;
                    if (!double.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                        throw new InvalidOperationException("Bad number '" + token + "'.");
                    return new NumberNode(value);
                }
                if (char.IsLetter(c) || c == '_')
                {
                    int start = _pos;
                    while (_pos < _text.Length && (char.IsLetterOrDigit(_text[_pos]) || _text[_pos] == '_' || _text[_pos] == ' '))
                        _pos++;
                    string name = _text.Substring(start, _pos - start).Trim();
                    if (name.Length == 0)
                        throw new InvalidOperationException("Unexpected '" + c + "'. Use $Column for columns.");
                    if (string.Equals(name, "TRUE", StringComparison.OrdinalIgnoreCase))
                        return new BoolNode(true);
                    if (string.Equals(name, "FALSE", StringComparison.OrdinalIgnoreCase))
                        return new BoolNode(false);
                    if (string.Equals(name, "IF", StringComparison.OrdinalIgnoreCase))
                    {
                        SkipSpaces();
                        if (_pos >= _text.Length || _text[_pos] != '(')
                            throw new InvalidOperationException("IF needs (condition, then, else).");
                        _pos++;
                        ExprNode cond = ParseOr();
                        SkipSpaces();
                        if (_pos >= _text.Length || _text[_pos] != ',')
                            throw new InvalidOperationException("IF needs (condition, then, else).");
                        _pos++;
                        ExprNode then = ParseOr();
                        SkipSpaces();
                        if (_pos >= _text.Length || _text[_pos] != ',')
                            throw new InvalidOperationException("IF needs (condition, then, else).");
                        _pos++;
                        ExprNode otherwise = ParseOr();
                        SkipSpaces();
                        if (_pos >= _text.Length || _text[_pos] != ')')
                            throw new InvalidOperationException("Missing ')' for IF.");
                        _pos++;
                        return new IfNode(cond, then, otherwise);
                    }
                    if (name.IndexOf(' ') < 0)
                    {
                        int save = _pos;
                        SkipSpaces();
                        if (_pos < _text.Length && _text[_pos] == '(')
                        {
                            _pos++;
                            var args = new List<ExprNode>();
                            SkipSpaces();
                            if (_pos < _text.Length && _text[_pos] == ')')
                            {
                                _pos++;
                                CheckArity(name, 0);
                                return new FunctionNode(name, args);
                            }
                            for (;;)
                            {
                                args.Add(ParseOr());
                                SkipSpaces();
                                if (_pos < _text.Length && _text[_pos] == ',')
                                {
                                    _pos++;
                                    continue;
                                }
                                if (_pos < _text.Length && _text[_pos] == ')')
                                {
                                    _pos++;
                                    CheckArity(name, args.Count);
                                    return new FunctionNode(name, args);
                                }
                                throw new InvalidOperationException("Missing ')' or ',' in " + name + "().");
                            }
                        }
                        _pos = save;
                    }
                    return new ColumnNode(name);
                }
                throw new InvalidOperationException("Unexpected '" + c + "'. Use $Column, $1, $Today, numbers, quotes, + - * / ^, comparisons, && || ! and functions like TODAY().");
            }

            void SkipSpaces()
            {
                while (_pos < _text.Length && char.IsWhiteSpace(_text[_pos]))
                    _pos++;
            }

            bool Peek(string token)
            {
                if (string.IsNullOrEmpty(token))
                    return false;
                if (_pos + token.Length > _text.Length)
                    return false;
                for (var i = 0; i < token.Length; i++)
                {
                    if (_text[_pos + i] != token[i])
                        return false;
                }
                return true;
            }

            bool Match(string token)
            {
                if (!Peek(token))
                    return false;
                _pos += token.Length;
                return true;
            }
        }
    }
}
