using System;
using System.Collections.Generic;
using WorkAssistant.Expressions;

namespace WorkAssistant.Features.FormulaGuide
{
    public sealed class SampleValue
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public sealed class GuideStep
    {
        public string Number { get; set; }
        public string Formula { get; set; }
        public string Result { get; set; }
        public string Note { get; set; }
    }

    public sealed class GuideLesson
    {
        public string Title { get; set; }
        public string Intro { get; set; }
        public string Tip { get; set; }
        public List<GuideStep> Steps { get; set; }
    }

    // Lesson text plus a live run of every example through ExpressionEngine.
    // Nothing here touches WPF, so the check harness can exercise it.
    public static class FormulaGuideWork
    {
        public static List<SampleValue> DefaultSamples()
        {
            return new List<SampleValue>
            {
                new SampleValue { Name = "$A", Value = "12" },
                new SampleValue { Name = "$B", Value = "5" },
                new SampleValue { Name = "$Name", Value = "Acme Corp" },
                new SampleValue { Name = "$Code", Value = "AB-1234/XY" },
                new SampleValue { Name = "$Messy", Value = "  Inv 0042 !!" },
                new SampleValue { Name = "$Price", Value = "19.99" },
                new SampleValue { Name = "$Due", Value = ExpressionEngine.FormatScalar(DateTime.Today.AddDays(30)) },
            };
        }

        public static List<GuideLesson> BuildLessons(IList<SampleValue> samples)
        {
            var lessons = new List<GuideLesson>();

            lessons.Add(Lesson("Plain numbers",
                "Arithmetic reads like you would write it. + - * / work on numbers, ^ raises to a power, and parentheses group. A number can come from a name or sit right in the formula.",
                null,
                Step("$A + $B", "Adds the two number values."),
                Step("$A * $B", "Multiplies them."),
                Step("($A + $B) / 2", "Parentheses run first."),
                Step("$A ^ 2", "12 squared. The power operator is ^."),
                Step("$A / 0", "Dividing by zero gives a blank result instead of an error."),
                Step("-$B", "A leading minus flips the sign.")));

            lessons.Add(Lesson("Text side by side",
                "& glues values into one piece of text. Names, numbers and quoted text mix freely; the quotes and the spaces are yours to place.",
                "Use & for text. + stays number-only when both sides are numbers, and glues otherwise.",
                Step("$Name & \" x\" & $A", "Fixed text sits between the values."),
                Step("\"INV-\" & $A & \"-\" & $B", "A code built from fixed and moving parts."),
                Step("$A & \" pcs at \" & $Price", "Numbers turn into text inside a glue chain."),
                Step("\"Total: \" + $Price", "+ glues when one side is text, so a formula that expects text still works.")));

            lessons.Add(Lesson("Compare and choose",
                "A comparison gives true or false: == != > >= < <=. Join them with && (and), || (or) and ! (not). IF(condition, then, else) picks between two values.",
                null,
                Step("$A > $B", "12 is bigger than 5."),
                Step("$Name == \"Acme Corp\"", "Text compares ignore case."),
                Step("$A > 10 && $Price < 20", "Both sides must pass for &&."),
                Step("$A > 100 || $B > 100", "One passing side is enough for ||."),
                Step("!($A > $B)", "! flips a result."),
                Step("IF($A > 10, \"Bulk order\", \"Single\")", "The first value when the condition is true, the second when it is false."),
                Step("IF($Price >= 20, $Price * 0.9, $Price) & \" each\"", "Branches can hold math. 19.99 is under 20, so the plain price passes through.")));

            lessons.Add(Lesson("Dates",
                "A date is a value like any other. $Today comes from the engine and $Due sits in the sample values. Add or subtract a number to move by days; subtract one date from another to count the days between them.",
                "Date patterns follow .NET format strings: dd day, MM month, yyyy year, ddd short weekday name.",
                Step("$Today", "Today with no time part."),
                Step("$Today + 7", "A week ahead."),
                Step("$Due - $Today", "$Due is today plus 30 days in the sample values. A negative result means the due date is already past."),
                Step("$Due > $Today", "Dates compare in order."),
                Step("ADDDAYS($Due, -3)", "The function form of subtracting 3 days."),
                Step("FORMAT($Due, \"ddd dd MMM yyyy\")", "Turns a date into text using a pattern you choose."),
                Step("IF($Due < $Today, \"Late\", \"On time\")", "A date check drives a decision.")));

            lessons.Add(Lesson("Built-in names",
                "The engine ships names for common dates. A name is read fresh every time a formula runs, so the same formula gives a new answer tomorrow.",
                "The Variables button in the top bar lists every built-in name with the value it has right now.",
                Step("$StartOfWeek", "Monday of this week."),
                Step("$EndOfMonth", "Last day of this month."),
                Step("$MonthAgo", "Same day last month."),
                Step("FORMAT($EndOfYear, \"dd/MM/yyyy\")", "A built-in works as a function argument."),
                Step("$CurrentYear & \"-\" & $B", "The year glued to a value.")));

            lessons.Add(Lesson("Clean up text",
                "These helpers tighten text before it lands in a file name or a cell.",
                "The value of $Messy starts with spaces and carries a symbol, so the helpers have something to chew on.",
                Step("TRIM($Messy)", "Strips the spaces at both ends."),
                Step("REMOVEDIGITS($Messy)", "Drops 0-9 and keeps everything else."),
                Step("CLEARSYMBOLS($Messy)", "Keeps letters, digits and spaces only."),
                Step("FIRSTWORD($Name)", "The piece up to the first space."),
                Step("LASTWORD($Name)", "The piece after the last space."),
                Step("CONTAINS($Name, \"acme\")", "Looks anywhere in the text and ignores case."),
                Step("STARTSWITH($Code, \"AB\")", "Checks the start."),
                Step("ENDSWITH($Code, \"XY\")", "Checks the end.")));

            lessons.Add(Lesson("Build one formula in steps",
                "A worked example that turns a count and a price into a finished line of text. Each step adds one piece to the formula.",
                null,
                Step("$A", "Start with the count."),
                Step("$A * $Price", "Multiply by the unit price."),
                Step("IF($A > 10, $A * $Price * 0.95, $A * $Price)", "Orders over 10 get 5% off. Nothing rounds here, so the decimals run deep."),
                Step("IF($A > 10, $A * $Price * 0.95, $A * $Price) & \" for \" & $Name", "Glue a label and the customer name onto the number."),
                Step("\"Line \" & $B & \": \" & IF($A > 10, $A * $Price * 0.95, $A * $Price) & \" for \" & $Name & \" (due \" & FORMAT($Due, \"dd MMM\") & \")\"", "The finished line. A long formula still reads left to right.")));

            lessons.Add(Lesson("Shapes to reuse",
                "Small formulas worth carrying into the other tools. Drop one in a Formula box and change the names to fit.",
                null,
                Step("\"INV-\" & $CurrentYear & \"-\" & $A", "A reference number from a fixed prefix, the year and a counter."),
                Step("IF($Due < $Today, \"LATE: \" & $Name, \"OK\")", "A status flag that only shouts when it needs to."),
                Step("FIRSTWORD($Name) & \"-\" & $B", "A short code from a word and a number."),
                Step("FORMAT($StartOfNextMonth, \"yyyy-MM\")", "A month key for grouping rows."),
                Step("CONTAINS($Name, \"corp\") && $A > 10", "A ready-made row filter, for a Formula box on the Filter & Sort page.")));

            // Results come from the engine, so the guide cannot drift from real behaviour.
            foreach (var lesson in lessons)
            {
                for (var i = 0; i < lesson.Steps.Count; i++)
                {
                    GuideStep step = lesson.Steps[i];
                    step.Number = (i + 1) + ".";
                    bool error;
                    string result = Run(step.Formula, samples, out error);
                    step.Result = error ? "problem: " + result : result;
                }
            }
            return lessons;
        }

        static GuideLesson Lesson(string title, string intro, string tip, params GuideStep[] steps)
        {
            return new GuideLesson { Title = title, Intro = intro, Tip = tip, Steps = new List<GuideStep>(steps) };
        }

        static GuideStep Step(string formula, string note)
        {
            return new GuideStep { Formula = formula, Note = note };
        }

        // Run one formula the same way the tools do and return display text.
        public static string Run(string formula, IList<SampleValue> samples, out bool isError)
        {
            isError = false;
            if (string.IsNullOrWhiteSpace(formula))
            {
                isError = true;
                return "Type a formula first.";
            }
            try
            {
                string text = ExpressionEngine.FormatScalar(ExpressionEngine.Evaluate(formula, Lookup(samples)));
                return text.Length == 0 ? "(blank)" : text;
            }
            catch (Exception ex)
            {
                isError = true;
                return ex.Message;
            }
        }

        // Sample values win, then built-ins. Anything else gets a clear message,
        // unlike the tools where a missing column simply reads as blank.
        public static Func<string, object> Lookup(IList<SampleValue> samples)
        {
            return delegate(string name)
            {
                if (samples != null)
                {
                    for (var i = 0; i < samples.Count; i++)
                    {
                        SampleValue row = samples[i];
                        if (row != null && Names(row.Name, name))
                            return row.Value;
                    }
                }
                object builtin;
                if (ExpressionEngine.TryGetVariable(name, out builtin))
                    return builtin;
                throw new InvalidOperationException("Nothing is named $" + name + ". Add it to the sample values or use a built-in such as $Today.");
            };
        }

        static bool Names(string left, string right)
        {
            return string.Equals(Strip(left), Strip(right), StringComparison.OrdinalIgnoreCase);
        }

        // "$A" and "A" mean the same thing.
        static string Strip(string name)
        {
            name = name != null ? name.Trim() : "";
            return name.StartsWith("$") ? name.Substring(1) : name;
        }
    }
}
