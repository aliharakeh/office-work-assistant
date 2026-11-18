using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.Highlight;
using System.Data;
using System.Globalization;
using System.IO;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;
using OfficeWorkAssistant.Features.CleanText;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.MergeDuplicates;
using OfficeWorkAssistant.Features.RemoveDuplicates;
using OfficeWorkAssistant.Features.SaveText;
using OfficeWorkAssistant.Features.Pdf;
using OfficeWorkAssistant.Features.Templates;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // What kind of file or folder a person running a saved pipeline has to give.
    public enum RunnerInputKind
    {
        ExcelFile,
        Folder,
        ExcelOutput,
        TextOutput,
        PdfFile,
        PdfOutput
    }

    // One thing the person running a saved pipeline must fill in: a file or folder the pipeline
    // reads or writes. Optional match rules are exposed separately by RuleGroups. Get and Set change the
    // pipeline in memory only; the saved pipeline file is never written.
    public sealed class RunnerInput
    {
        public readonly string NodeId;
        public readonly string StepTitle;
        public readonly RunnerInputKind Kind;
        // "The Excel file to read"
        public readonly string Label;
        // "Step 1 - reads one sheet of it"
        public readonly string Note;
        // Where to save: ".xlsx", ".txt" or ".json". Blank when it does not apply.
        public string DefaultExtension = "";
        // Text output only: the file holds JSON, not plain text.
        public bool IsJson;

        readonly Func<string> _get;
        readonly Action<string> _set;
        readonly Func<string> _getSheet;
        readonly Action<string> _setSheet;
        readonly Func<string> _check;
        string _sheetsFor;
        string[] _sheets;
        string _sheetProblem;

        public RunnerInput(string nodeId, string stepTitle, RunnerInputKind kind, string label, string note,
            Func<string> get, Action<string> set)
            : this(nodeId, stepTitle, kind, label, note, get, set, null, null, null)
        {
        }

        public RunnerInput(string nodeId, string stepTitle, RunnerInputKind kind, string label, string note,
            Func<string> get, Action<string> set, Func<string> check)
            : this(nodeId, stepTitle, kind, label, note, get, set, check, null, null)
        {
        }

        public RunnerInput(string nodeId, string stepTitle, RunnerInputKind kind, string label, string note,
            Func<string> get, Action<string> set, Func<string> check, Func<string> getSheet, Action<string> setSheet)
        {
            NodeId = nodeId;
            StepTitle = stepTitle;
            Kind = kind;
            Label = label;
            Note = note;
            _get = get;
            _set = set;
            _check = check;
            _getSheet = getSheet;
            _setSheet = setSheet;
        }

        public bool IsOutput
        {
            get { return Kind == RunnerInputKind.ExcelOutput || Kind == RunnerInputKind.TextOutput || Kind == RunnerInputKind.PdfOutput; }
        }

        public string Value
        {
            get { return _get() ?? ""; }
            set { _set(value ?? ""); }
        }

        // Load file steps: which sheet to read. Blank means the first sheet.
        public bool HasSheet
        {
            get { return _getSheet != null; }
        }

        public string Sheet
        {
            get { return _getSheet != null ? _getSheet() ?? "" : ""; }
            set { if (_setSheet != null) _setSheet(value ?? ""); }
        }

        // The sheet names of the chosen file, or null when it cannot be read yet.
        // The answer is kept until another file is chosen.
        public string[] SheetNames()
        {
            var path = Value.Trim();
            if (path.Length == 0 || !File.Exists(path))
            {
                _sheetsFor = null;
                _sheets = null;
                _sheetProblem = null;
                return null;
            }
            if (string.Equals(_sheetsFor, path, StringComparison.OrdinalIgnoreCase) && _sheets != null)
                return _sheets;
            _sheetsFor = path;
            _sheetProblem = null;
            try
            {
                _sheets = ExcelFile.Load(path, null, 1).Sheets;
            }
            catch (Exception ex)
            {
                _sheets = null;
                _sheetProblem = ex.Message;
            }
            return _sheets;
        }

        public string SheetProblem
        {
            get { return _sheetProblem; }
        }

        // A problem with this value that only the step it belongs to can tell, or null.
        public string Problem()
        {
            return _check != null ? _check() : null;
        }
    }

    public sealed class RunnerDuplicateRules
    {
        readonly RemoveDuplicatesSettings _settings;
        public RunnerDuplicateRules(RemoveDuplicatesSettings settings) { _settings = settings; }
        public bool UseFormula { get { return _settings.UseFormula; } set { _settings.UseFormula = value; } }
        public string KeyFormula { get { return _settings.KeyFormula; } set { _settings.KeyFormula = value; } }
        public string Columns
        {
            get { return string.Join(", ", (_settings.Keys ?? new List<string>()).ToArray()); }
            set
            {
                var keys = new List<string>();
                foreach (var key in (value ?? "").Split(','))
                    if (key.Trim().Length > 0) keys.Add(key.Trim());
                _settings.Keys = keys;
            }
        }
    }

    // Only these explicitly selected properties may be edited during a run.
    public sealed class RunnerRuleGroup
    {
        public string Label;
        public string HelpText;
        public List<VariableHelp> Variables;
        public IList Rows;
        public string[] Fields;
        public bool CanAdd;
        public Func<bool> GetMatchAll;
        public Action<bool> SetMatchAll;

        public RunnerRuleGroup(string label, IList rows, bool canAdd, params string[] fields)
        {
            Label = label;
            Rows = rows;
            CanAdd = canAdd;
            Fields = fields;
        }
    }

    // One line of the plain-language list of what a saved pipeline does.
    public sealed class RunnerStepNote
    {
        public PipelineStepKind Kind;
        public string Label;
        public string Text;
        public string Hint;
    }

    // User mode: what a saved pipeline asks the person running it, what it does in plain words,
    // and what happened. Nothing here writes the pipeline file or touches the disk.
    public static class RunnerWork
    {
        // ---------- what the user has to fill in ----------

        // The file and folder choices of a pipeline, in running order.
        public static List<RunnerInput> Inputs(PipelineDefinition def)
        {
            var found = new List<RunnerInput>();
            var order = PipelineWork.Order(def);
            for (var i = 0; i < order.Count; i++)
            {
                var node = order[i];
                var step = "Step " + (i + 1).ToString(CultureInfo.InvariantCulture);
                switch (node.Kind)
                {
                    case PipelineStepKind.Load:
                    {
                        var s = node.Settings as LoadStepSettings;
                        if (s == null)
                            break;
                        var load = s;
                        found.Add(new RunnerInput(node.Id, node.Title, RunnerInputKind.ExcelFile,
                            "The Excel file to read", step + " - reads one sheet of it",
                            () => load.Path, v => load.Path = v, null, () => load.Sheet, v => load.Sheet = v));
                        break;
                    }
                    case PipelineStepKind.LoadPdf:
                    {
                        var load = node.Settings as LoadPdfSettings;
                        if (load == null || PipelineWork.InputLink(def, node.Id, PipelineWork.PortIn) != null)
                            break;
                        found.Add(new RunnerInput(node.Id, node.Title, RunnerInputKind.PdfFile,
                            "The PDF file to read", step + " - reads the document's pages",
                            () => load.Path, v => load.Path = v));
                        break;
                    }
                    case PipelineStepKind.SavePdf:
                    {
                        var save = node.Settings as SavePdfSettings;
                        if (save == null)
                            break;
                        var output = new RunnerInput(node.Id, node.Title, RunnerInputKind.PdfOutput,
                            "Where to save the PDF result", step + " - adds numbers to the filenames when saving several PDFs",
                            () => save.Path, v => save.Path = v);
                        output.DefaultExtension = ".pdf";
                        found.Add(output);
                        break;
                    }
                    case PipelineStepKind.ListFolder:
                    {
                        var s = node.Settings as FolderScanSettings;
                        if (s == null)
                            break;
                        var scan = s;
                        found.Add(new RunnerInput(node.Id, node.Title, RunnerInputKind.Folder,
                            "The folder to look in", step + " - " + DescribeScan(scan),
                            () => scan.Folder, v => scan.Folder = v));
                        break;
                    }
                    case PipelineStepKind.FindFiles:
                    {
                        var s = node.Settings as FindFilesSettings;
                        if (s == null || s.Scan == null)
                            break;
                        var scan = s.Scan;
                        found.Add(new RunnerInput(node.Id, node.Title, RunnerInputKind.Folder,
                            "The folder to search", step + " - looks for the listed files" +
                                (scan.Recursive ? " there and in its subfolders" : " directly in this folder"),
                            () => scan.Folder, v => scan.Folder = v));
                        break;
                    }
                    case PipelineStepKind.MergeFolders:
                    {
                        var s = node.Settings as MergeFoldersSettings;
                        if (s == null)
                            break;
                        var merge = s;
                        found.Add(new RunnerInput(node.Id, node.Title, RunnerInputKind.Folder,
                            "The folder to merge", step + " - merges the subfolders in it that belong together",
                            () => merge.Folder, v => merge.Folder = v));
                        break;
                    }
                    case PipelineStepKind.FileAction:
                    {
                        var s = node.Settings as FileActionSettings;
                        if (s == null || s.Action == FileActionKind.Delete)
                            break;
                        var action = s;
                        found.Add(new RunnerInput(node.Id, node.Title, RunnerInputKind.Folder,
                            action.Action == FileActionKind.Copy ? "Where the copies go" : "Where the files go",
                            step + " - " + DescribeConflict(action),
                            () => action.DestFolder, v => action.DestFolder = v));
                        break;
                    }
                    case PipelineStepKind.Save:
                    {
                        var s = node.Settings as SaveStepSettings;
                        if (s == null)
                            break;
                        var save = s;
                        var output = new RunnerInput(node.Id, node.Title, RunnerInputKind.ExcelOutput,
                            "Where to save the Excel result",
                            step + " - " + (save.Mode == SaveStepMode.NewSheet
                                ? "adds a new sheet named \"" + (save.SheetName ?? "") + "\" to this workbook"
                                : "replaces this workbook when it already exists"),
                            () => save.Path, v => save.Path = v, () => PipelineWork.CheckSave(save));
                        output.DefaultExtension = ".xlsx";
                        found.Add(output);
                        break;
                    }
                    case PipelineStepKind.SaveText:
                    {
                        var s = node.Settings as SaveTextSettings;
                        if (s == null)
                            break;
                        var text = s;
                        var output = new RunnerInput(node.Id, node.Title, RunnerInputKind.TextOutput,
                            SaveTextWork.IsJson(text.Format) ? "Where to save the JSON result" : "Where to save the text result",
                            step + " - " + SaveTextWork.FormatLabel(text.Format),
                            () => text.Path, v => text.Path = v, () => CheckTextOutput(text));
                        output.IsJson = SaveTextWork.IsJson(text.Format);
                        output.DefaultExtension = output.IsJson ? ".json" : ".txt";
                        found.Add(output);
                        break;
                    }
                }
            }
            return found;
        }

        public static bool HasMatchRules(PipelineStepKind kind)
        {
            switch (kind)
            {
                case PipelineStepKind.Compare:
                case PipelineStepKind.FillColumns:
                case PipelineStepKind.FilterSort:
                case PipelineStepKind.Highlight:
                case PipelineStepKind.Templates:
                case PipelineStepKind.ListFolder:
                case PipelineStepKind.FindFiles:
                case PipelineStepKind.MergeFolders:
                case PipelineStepKind.RemoveDuplicates:
                    return true;
                default: return false;
            }
        }

        public static List<RunnerRuleGroup> RuleGroups(PipelineNode node)
        {
            var groups = new List<RunnerRuleGroup>();
            if (!node.EditMatchRulesInRunMode || node.Settings == null || !HasMatchRules(node.Kind))
                return groups;
            var compare = node.Settings as ExcelProcessingSettings;
            if (compare != null)
            {
                if (compare.Matches == null) compare.Matches = new List<ExcelProcessingMatch>();
                var group = new RunnerRuleGroup("Match rules", new BindingList<ExcelProcessingMatch>(compare.Matches), true,
                    "ColumnA", "ExpressionA", "SplitA", "Operator", "ColumnB", "ExpressionB", "SplitB");
                group.GetMatchAll = () => compare.MatchAll;
                group.SetMatchAll = v => compare.MatchAll = v;
                groups.Add(group);
            }
            var fill = node.Settings as FillColumnsSettings;
            if (fill != null)
            {
                if (fill.Keys == null) fill.Keys = new List<FillColumnsKey>();
                var group = new RunnerRuleGroup("Match rows", new BindingList<FillColumnsKey>(fill.Keys), true, "FormulaA", "FormulaB");
                group.GetMatchAll = () => fill.MatchAll;
                group.SetMatchAll = v => fill.MatchAll = v;
                groups.Add(group);
                foreach (var rule in fill.Rules ?? new List<FillColumnsRule>())
                    groups.Add(new RunnerRuleGroup("Fill condition for " + rule.Target, new[] { rule }, false, "When"));
            }
            var filter = node.Settings as FilterSortSettings;
            if (filter != null)
            {
                if (filter.Conditions == null) filter.Conditions = new List<FilterCondition>();
                var group = new RunnerRuleGroup("Filter rules", new BindingList<FilterCondition>(filter.Conditions), true,
                    "Kind", "Column", "Operator", "Value", "Expression");
                group.GetMatchAll = () => filter.MatchAll;
                group.SetMatchAll = v => filter.MatchAll = v;
                groups.Add(group);
            }
            var highlight = node.Settings as HighlightSettings;
            if (highlight != null)
                foreach (var rule in highlight.Rules ?? new List<HighlightRule>())
                    groups.Add(new RunnerRuleGroup(rule.ToString(), new[] { rule }, false, "Condition"));
            var template = node.Settings as TemplateDefinition;
            if (template != null)
                foreach (var column in template.Columns ?? new List<TemplateColumn>())
                    if (column.Kind == TemplateColumnKind.Conditional)
                        groups.Add(new RunnerRuleGroup("Condition for " + column.Name, new[] { column }, false, "Condition"));
            var scan = node.Settings as FolderScanSettings;
            var find = node.Settings as FindFilesSettings;
            if (find != null) scan = find.Scan;
            if (scan != null)
                groups.Add(new RunnerRuleGroup("Folder scan filters (blank means all)", new[] { scan }, false, "FileFilter", "FolderFilter"));
            if (find != null)
                groups.Add(new RunnerRuleGroup("File matching", new[] { find }, false, "Match", "FileFormula", "FolderFormula"));
            var merge = node.Settings as MergeFoldersSettings;
            if (merge != null)
                groups.Add(new RunnerRuleGroup("Folder match", new[] { merge }, false, "Condition"));
            var dedupe = node.Settings as RemoveDuplicatesSettings;
            if (dedupe != null)
                groups.Add(new RunnerRuleGroup("Duplicate key (column headers separated by commas; blank means all columns)",
                    new[] { new RunnerDuplicateRules(dedupe) }, false, "UseFormula", "KeyFormula", "Columns"));
            foreach (var group in groups)
                AddRuleHelp(group);
            return groups;
        }

        static void AddRuleHelp(RunnerRuleGroup group)
        {
            group.HelpText = "Use $A for the first column, $B for the second, or a header such as $Qty. " +
                "Compare with == != > >= < <=. Join conditions with && (AND) or || (OR). Put text in double quotes, for example $Status == \"Late\". " +
                "Built-in variables and functions are listed below.";
            group.Variables = new List<VariableHelp>
            {
                new VariableHelp { Name = "$A, $B, ...", Description = "Columns of the current row by Excel letter (A = first column)." },
                new VariableHelp { Name = "$Qty", Description = "Example of a column by header name; use your own header. Excel letters take precedence over headers." }
            };
            if (group.Rows is BindingList<ExcelProcessingMatch>)
            {
                group.HelpText = "Expression A reads only the A row; Expression B reads only the B row. Column A / Column B choose the column used by $Value. " +
                    "Split A / Split B set the separator for $1, $2, ... on that side. Example: TRIM($Value) on both sides with operator ==. " +
                    "Leave one expression blank to test the other side as a true/false formula. At least one rule must use both sides.";
                group.Variables.Add(new VariableHelp { Name = "$Value", Description = "Value of the chosen column on this side." });
                group.Variables.Add(new VariableHelp { Name = "$1", Description = "First part of the chosen column after splitting with this side's separator.", Example = "Acme" });
                group.Variables.Add(new VariableHelp { Name = "$2", Description = "Second split part. Use $3, $4, ... for later parts.", Example = "2024" });
            }
            else if (group.Rows is BindingList<FillColumnsKey>)
            {
                group.HelpText = "Formula A reads only the A row; Formula B reads only the B row. Their results are compared without case or surrounding spaces. " +
                    "Blank keys do not match. Example: TRIM($Code) on both sides. The first matching A row supplies the fill values.";
                group.Variables.Add(new VariableHelp { Name = "$A_Code / $B_Code", Description = "In Formula A, $A_Code refers to A's Code column. In Formula B, $B_Code refers to B's Code column. Each formula sees only its own side." });
            }
            else if (group.Rows is FillColumnsRule[])
            {
                group.HelpText = "When is true/false; blank means always. Fill conditions can read both rows and earlier fill results. " +
                    "Example: $Matched && $Value == \"\". An unprefixed column reads B first, then A if B has no such header.";
                group.Variables.Add(new VariableHelp { Name = "$A_Name / $A_C", Description = "Column by header or letter of the matched A row; blank when no row matched." });
                group.Variables.Add(new VariableHelp { Name = "$B_Name / $B_C", Description = "Column of the B row being filled, including earlier fill results." });
                group.Variables.Add(new VariableHelp { Name = "$Value", Description = "Current value of the target cell." });
                group.Variables.Add(new VariableHelp { Name = "$Matched", Description = "TRUE when an A row matched this B row." });
            }
            else if (group.Rows is FolderScanSettings[] || group.Rows is FindFilesSettings[])
            {
                var matching = group.Rows is FindFilesSettings[];
                group.HelpText = matching
                    ? "In Formula match mode, File Formula and Folder Formula return true when an item belongs to the input key. Example: $1 == $Key && $Ext == \".pdf\"."
                    : "File Filter and Folder Filter test items while scanning, before key matching. Blank means keep all. Example: $Ext == \".pdf\". Key variables are unavailable in scan filters.";
                group.Variables = FileOpsWork.FileVariableHelp(matching, false);
            }
            else if (group.Rows is MergeFoldersSettings[])
            {
                group.HelpText = "Compare two subfolders A and B. A true condition groups them for merging. Example: $A1 == $B1. Split parts use the separator saved by the builder.";
                group.Variables = new List<VariableHelp>
                {
                    new VariableHelp { Name = "$A / $B", Description = "Name of folder A / folder B.", Example = "Acme_2024" },
                    new VariableHelp { Name = "$A1, $A2, ...", Description = "Split parts of folder A's name.", Example = "Acme" },
                    new VariableHelp { Name = "$B1, $B2, ...", Description = "Split parts of folder B's name.", Example = "Acme" }
                };
            }
            else if (group.Rows is HighlightRule[])
                group.Variables.Add(new VariableHelp { Name = "$Value", Description = "Current cell, only for rules that test each cell separately. Row and column rules use row columns instead." });
            else if (group.Rows is RunnerDuplicateRules[])
                group.HelpText += " With Use Formula enabled, rows with the same Key Formula result are duplicates; otherwise Columns lists header names without $.";
            else if (group.Rows is BindingList<FilterCondition>)
                group.HelpText += " For Kind = Formula, use Expression. For Kind = Column, use the saved header in Column (without $), Operator and Value; Value also accepts built-in variables.";
        }

        static string DescribeScan(FolderScanSettings scan)
        {
            var what = scan.Include == FileInclude.Files ? "lists the files"
                : scan.Include == FileInclude.Folders ? "lists the folders"
                : "lists the files and folders";
            return what + (scan.Recursive ? " there and in its subfolders" : " directly in this folder");
        }

        static string DescribeConflict(FileActionSettings action)
        {
            var verb = action.Action == FileActionKind.Copy ? "copies" : "moves";
            var what = verb + " the listed files and folders there" + (action.KeepStructure ? ", keeping the folder structure" : "");
            if (action.Conflict == FileConflict.Skip)
                return what + ". Files already there are left alone.";
            if (action.Conflict == FileConflict.Overwrite)
                return what + ". Files already there are replaced.";
            return what + ". Files already there keep both, the new one with a number.";
        }

        static string CheckTextOutput(SaveTextSettings s)
        {
            try
            {
                SaveTextWork.OutputPath(s);
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // What is still wrong with one choice, in plain words, or null when it is ready.
        public static string InputProblem(RunnerInput input)
        {
            try
            {
                return InputProblemChecked(input);
            }
            catch (Exception ex)
            {
                // A path with characters Windows does not allow can come from Builder mode.
                return "This path cannot be used: " + ex.Message;
            }
        }

        static string InputProblemChecked(RunnerInput input)
        {
            var path = input.Value.Trim();
            if (path.Length == 0)
                return "Fill in \"" + input.Label + "\" for " + input.StepTitle + ".";
            if (input.Kind == RunnerInputKind.ExcelFile)
            {
                if (!File.Exists(path))
                    return "The file does not exist: " + path;
                // A sheet the pipeline remembers from another workbook is caught before the run.
                if (input.HasSheet && !string.IsNullOrEmpty(input.Sheet))
                {
                    var names = input.SheetNames();
                    if (!HasName(names, input.Sheet))
                        return "The file has no sheet named \"" + input.Sheet + "\". It has: " + Join(names) + ".";
                }
            }
            else if (input.Kind == RunnerInputKind.PdfFile)
                PdfWork.InputPath(path);
            else if (input.Kind == RunnerInputKind.PdfOutput)
                PdfWork.OutputPath(new SavePdfSettings { Path = path });
            else if (input.Kind == RunnerInputKind.Folder)
            {
                if (!Directory.Exists(path))
                    return "The folder does not exist: " + path;
            }
            else
            {
                var folder = Path.GetDirectoryName(path);
                if (string.IsNullOrEmpty(folder))
                    return "Use a full path, including the drive, for: " + path;
                if (!Directory.Exists(folder))
                    return "The folder does not exist: " + folder;
                if (input.Kind == RunnerInputKind.ExcelOutput &&
                    !string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    return "The Excel result must be saved as a .xlsx file: " + path;
            }
            return input.Problem();
        }

        // The first thing the user still has to fix, or null when every choice is ready.
        public static string CheckInputs(IEnumerable<RunnerInput> inputs)
        {
            foreach (var input in inputs)
            {
                var problem = InputProblem(input);
                if (problem != null)
                    return problem;
            }
            return null;
        }

        // ---------- what the pipeline does, in plain words ----------

        public static List<RunnerStepNote> Describe(PipelineDefinition def)
        {
            var notes = new List<RunnerStepNote>();
            foreach (var node in PipelineWork.Order(def))
            {
                notes.Add(new RunnerStepNote
                {
                    Kind = node.Kind,
                    Label = PipelineWork.KindLabel(node.Kind),
                    Text = DescribeNode(node),
                    Hint = PipelineWork.Summary(node)
                });
            }
            return notes;
        }

        static string DescribeNode(PipelineNode node)
        {
            var s = node.Settings;
            switch (node.Kind)
            {
                case PipelineStepKind.Load:
                    return "Reads one sheet of an Excel file.";
                case PipelineStepKind.Save:
                {
                    var save = s as SaveStepSettings;
                    return save != null && save.Mode == SaveStepMode.NewSheet
                        ? "Adds the result to a workbook as a new sheet."
                        : "Writes the result to an Excel file.";
                }
                case PipelineStepKind.SaveText:
                {
                    var text = s as SaveTextSettings;
                    if (text == null)
                        return "Writes the result to a text or JSON file.";
                    if (text.Format == SaveTextFormat.Text)
                        return "Writes the result to a text file, one row per line.";
                    if (text.Format == SaveTextFormat.TextMap)
                        return "Writes the result to a text file, one key and value per line.";
                    return SaveTextWork.IsMap(text.Format)
                        ? "Writes the result to a JSON object, one key and value per entry."
                        : "Writes the result to a JSON array, one value per row.";
                }
                case PipelineStepKind.FilterSort:
                {
                    var filter = s as FilterSortSettings;
                    if (filter == null)
                        return "Filters and sorts the rows.";
                    var count = filter.Conditions == null ? 0 : filter.Conditions.Count;
                    var sorts = filter.Sorts == null ? 0 : filter.Sorts.Count;
                    var line = count > 0 && sorts > 0 ? "Keeps only the rows that match, then sorts them."
                        : count > 0 ? "Keeps only the rows that match."
                        : sorts > 0 ? "Sorts the rows."
                        : "Leaves the rows in their order.";
                    return line + (filter.KeepAll ? "" : " Keeps only the chosen columns.");
                }
                case PipelineStepKind.Templates:
                {
                    var template = s as TemplateDefinition;
                    var columns = template == null || template.Columns == null ? 0 : template.Columns.Count;
                    return "Adds " + Count(columns, "new column") + " built from formulas." +
                        (template != null && !template.KeepSourceColumns ? " Keeps only the chosen source columns." : "");
                }
                case PipelineStepKind.Compare:
                {
                    var compare = s as ExcelProcessingSettings;
                    var op = compare == null ? "Common" : compare.Operation;
                    if (op == "AOnly")
                        return "Compares two tables and keeps the rows only in the first one.";
                    if (op == "BOnly")
                        return "Compares two tables and keeps the rows only in the second one.";
                    return "Compares two tables and keeps the rows they both have.";
                }
                case PipelineStepKind.FillColumns:
                    return "Copies values from the first table into the matching rows of the second.";
                case PipelineStepKind.ValueList:
                {
                    var list = s as ValueListSettings;
                    return list != null && list.Source == ValueListSource.Stored
                        ? "Uses the key list saved inside the pipeline."
                        : "Builds the key list from the rows before it.";
                }
                case PipelineStepKind.ListFolder:
                    return "Lists the files and folders of a folder.";
                case PipelineStepKind.FindFiles:
                    return "Finds the files and folders whose names match the keys.";
                case PipelineStepKind.FileAction:
                {
                    var action = s as FileActionSettings;
                    if (action == null)
                        return "Copies, moves or deletes the listed files and folders.";
                    if (action.Action == FileActionKind.Copy)
                        return "Copies the listed files and folders.";
                    if (action.Action == FileActionKind.Move)
                        return "Moves the listed files and folders.";
                    return action.Permanent
                        ? "Deletes the listed files and folders for good."
                        : "Sends the listed files and folders to the Recycle Bin.";
                }
                case PipelineStepKind.MergeFolders:
                    return "Merges subfolders that belong together.";
                case PipelineStepKind.CleanText:
                {
                    var clean = s as CleanTextSettings;
                    return clean == null ? "Tidies the text cells." : "Tidies the text: " + clean.Describe() + ".";
                }
                case PipelineStepKind.SplitColumn:
                    return "Splits one column into several new ones.";
                case PipelineStepKind.CombineColumns:
                    return "Joins several columns into one.";
                case PipelineStepKind.RemoveDuplicates:
                {
                    var dedupe = s as RemoveDuplicatesSettings;
                    return dedupe == null
                        ? "Removes duplicate rows."
                        : "Removes duplicate rows, " + LowerFirst(RemoveDuplicatesWork.KeepLabel(dedupe.Keep)) + ".";
                }
                case PipelineStepKind.ArrangeColumns:
                    return "Picks, renames and reorders the columns.";
                case PipelineStepKind.Append:
                    return "Stacks the rows of the tables before it into one table.";
                case PipelineStepKind.Highlight:
                    return "Colours the cells or rows that match the saved rules.";
                case PipelineStepKind.FormatSheet:
                    return "Sets the sheet's look: header style, widths, freeze and number formats.";
                default:
                    return PipelineWork.KindDescription(node.Kind);
            }
        }

        // "5 steps - 2 inputs - 1 output - changes files"
        public static string Headline(PipelineDefinition def)
        {
            var order = PipelineWork.Order(def);
            var reads = 0;
            var writes = 0;
            var actions = 0;
            foreach (var node in order)
            {
                if (HasSourceInput(def, node))
                    reads++;
                if (PipelineWork.IsSaveKind(node.Kind))
                    writes++;
                if (PipelineWork.IsActionKind(node.Kind))
                    actions++;
            }
            var parts = new List<string>();
            parts.Add(Count(order.Count, "step"));
            if (reads > 0)
                parts.Add(Count(reads, "input"));
            if (writes > 0)
                parts.Add(Count(writes, "output"));
            if (actions > 0)
                parts.Add("changes files");
            return string.Join(" - ", parts.ToArray());
        }

        // A step whose file or folder the user chooses (not one that only reads the row it is given).
        static bool HasSourceInput(PipelineDefinition def, PipelineNode node)
        {
            switch (node.Kind)
            {
                case PipelineStepKind.LoadPdf:
                    return node.Settings != null && PipelineWork.InputLink(def, node.Id, PipelineWork.PortIn) == null;
                case PipelineStepKind.Load:
                case PipelineStepKind.ListFolder:
                case PipelineStepKind.FindFiles:
                case PipelineStepKind.MergeFolders:
                case PipelineStepKind.FileAction:
                    return node.Settings != null &&
                        (node.Kind != PipelineStepKind.FileAction || ((FileActionSettings)node.Settings).Action != FileActionKind.Delete);
                default:
                    return false;
            }
        }

        // True when the pipeline writes or changes something on disk.
        public static bool WritesAnything(PipelineDefinition def)
        {
            foreach (var node in def.Nodes)
            {
                if (PipelineWork.IsSaveKind(node.Kind) || PipelineWork.IsActionKind(node.Kind))
                    return true;
            }
            return false;
        }

        // ---------- checking, running, reporting ----------

        // "Step 3 (Filter & Sort)"
        public static string StepName(PipelineDefinition def, PipelineNode node)
        {
            var order = PipelineWork.Order(def);
            for (var i = 0; i < order.Count; i++)
            {
                if (order[i].Id == node.Id)
                    return "Step " + (i + 1).ToString(CultureInfo.InvariantCulture) + " (" + PipelineWork.KindLabel(node.Kind) + ")";
            }
            return PipelineWork.KindLabel(node.Kind);
        }

        // What the run would write, one line per output, for the Check button.
        public static List<string> OutputLines(PipelineDefinition def, IDictionary<string, DataTable> cache)
        {
            var lines = new List<string>();
            foreach (var node in PipelineWork.Order(def))
            {
                if (!PipelineWork.IsSaveKind(node.Kind))
                    continue;
                try
                {
                    DataTable table = null;
                    if (cache != null)
                        cache.TryGetValue(node.Id, out table);
                    if (node.Kind == PipelineStepKind.SavePdf)
                    {
                        if (table != null)
                            foreach (var item in PipelineWork.ResultItems(table))
                                foreach (DataRow output in item.Table.Rows)
                                    lines.Add("Writes " + output["Path"] + " - " + Count((int)output["Pages"], "page") +
                                        ((string)output["Status"] == "Overwrite" ? ", replacing the file." : "."));
                        continue;
                    }

                    var text = node.Settings as SaveTextSettings;
                    var save = node.Settings as SaveStepSettings;
                    var items = table == null ? null : PipelineWork.ResultItems(table);
                    int index = 0;
                    foreach (string path in PipelineWork.SaveTargetPaths(node, cache))
                    {
                        var rows = items == null ? "" : RowsOf(items[index++].Table, text != null);
                        if (text != null)
                        {
                            lines.Add("Writes " + path + rows + " as " + SaveTextWork.FormatLabel(text.Format) + ".");
                            continue;
                        }
                        lines.Add("Writes " + path + rows +
                            (save != null && save.Mode == SaveStepMode.NewSheet
                                ? ", added as a new sheet named \"" + (save.SheetName ?? "") + "\"" +
                                    (File.Exists(path) ? " (the workbook must not have that sheet yet)." : ".")
                                : File.Exists(path) ? ", replacing the file." : ", as a new file."));
                    }
                }
                catch (Exception ex)
                {
                    lines.Add("Could not prepare " + StepName(def, node) + ": " + ex.Message);
                }
            }
            return lines;
        }

        static string RowsOf(DataTable table, bool textPlan)
        {
            var rows = table.Rows.Count;
            var column = textPlan ? ExcelFile.FindColumn(table, "Rows") : null;
            if (column != null && table.Rows.Count > 0)
            {
                try
                {
                    rows = Convert.ToInt32(table.Rows[0][column], CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                }
            }
            return " - " + Count(rows, "row");
        }

        // Files that exist and would be replaced, for the confirmation before a run.
        public static List<string> ReplacedFiles(PipelineDefinition def, IDictionary<string, DataTable> cache = null)
        {
            var files = new List<string>();
            foreach (var node in PipelineWork.Order(def))
            {
                if (!PipelineWork.IsSaveKind(node.Kind))
                    continue;
                try
                {
                    files.AddRange(PipelineWork.ReplacementPaths(node, cache));
                }
                catch (Exception)
                {
                }
            }
            return files;
        }

        // What each step that changes files would do, for the Check button.
        public static List<string> ActionLines(PipelineDefinition def, IDictionary<string, DataTable> cache)
        {
            var lines = new List<string>();
            foreach (var node in PipelineWork.Order(def))
            {
                if (!PipelineWork.IsActionKind(node.Kind))
                    continue;
                DataTable plan = null;
                if (cache != null)
                    cache.TryGetValue(node.Id, out plan);
                lines.Add(StepName(def, node) + ": " + PipelineWork.DescribeAction(node, plan));
            }
            return lines;
        }

        // Why the run stopped, in plain words, or null when it worked.
        public static string FailureText(PipelineDefinition def, PipelineRunResult result)
        {
            if (result == null || result.Error == null)
                return null;
            var text = result.Error;
            if (result.FailedNodeId != null)
            {
                var node = PipelineWork.Find(def, result.FailedNodeId);
                if (node != null)
                    text = StepName(def, node) + " could not finish.\n\n" + result.Error;
            }
            var done = new List<string>();
            foreach (var path in result.Written)
                done.Add("Saved " + path);
            foreach (var action in result.Actions)
                done.Add(action);
            text += done.Count == 0
                ? "\n\nNothing was written and no file was changed."
                : "\n\nThese were already done:\n" + string.Join("\n", done.ToArray());
            return text + "\n\nTell whoever built this pipeline, so they can fix it in Builder mode.";
        }

        // One line per file written and per file action, for the result panel.
        public static List<string> DoneLines(PipelineRunResult result)
        {
            var lines = new List<string>();
            if (result == null)
                return lines;
            foreach (var path in result.Written)
                lines.Add("Saved " + path);
            foreach (var action in result.Actions)
                lines.Add(action);
            return lines;
        }

        public static string DoneHeadline(PipelineRunResult result)
        {
            if (result == null)
                return "";
            if (result.Written.Count == 0 && result.Actions.Count == 0)
                return "Finished. Nothing needed to be written.";
            return "Finished. Saved " + Count(result.Written.Count, "file") + " and ran " +
                Count(result.Actions.Count, "file action") + ".";
        }

        // Folders worth opening after a run: where the results went and where files were moved.
        public static List<string> Folders(PipelineDefinition def, PipelineRunResult result)
        {
            var folders = new List<string>();
            if (result != null)
            {
                foreach (var written in result.Written)
                {
                    var name = written;
                    var mark = name.IndexOf(" (sheet", StringComparison.Ordinal);
                    if (mark >= 0)
                        name = name.Substring(0, mark);
                    AddFolder(folders, Path.GetDirectoryName(name));
                }
            }
            foreach (var node in def.Nodes)
            {
                var action = node.Settings as FileActionSettings;
                if (action != null && action.Action != FileActionKind.Delete)
                    AddFolder(folders, action.DestFolder);
                var merge = node.Settings as MergeFoldersSettings;
                if (merge != null)
                    AddFolder(folders, merge.Folder);
            }
            return folders;
        }

        static void AddFolder(List<string> folders, string folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return;
            foreach (var known in folders)
            {
                if (string.Equals(known, folder, StringComparison.OrdinalIgnoreCase))
                    return;
            }
            folders.Add(folder);
        }

        // True when the sheet list could not be read (the file's own message follows) or holds the name.
        static bool HasName(string[] names, string name)
        {
            if (names == null || names.Length == 0)
                return true;
            foreach (var candidate in names)
            {
                if (string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        static string Join(string[] names)
        {
            return names == null ? "" : string.Join(", ", names);
        }

        static string LowerFirst(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            return char.ToLowerInvariant(text[0]) + text.Substring(1);
        }

        static string Count(int n, string noun)
        {
            return n.ToString(CultureInfo.InvariantCulture) + " " + noun + (n == 1 ? "" : "s");
        }
    }
}
