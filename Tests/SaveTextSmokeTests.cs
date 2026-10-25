using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.Pipeline;
using OfficeWorkAssistant.Features.SaveText;

internal static class SaveTextSmokeTests
{
    static int _checks;
    static readonly string Arabic = "\u0645\u0631\u062d\u0628\u0627 \u0628\u0627\u0644\u0639\u0627\u0644\u0645";

    static void Check(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
        _checks++;
    }

    static void Fails(Action action, string expected)
    {
        try { action(); }
        catch (Exception ex)
        {
            Check(ex.Message.IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0,
                "Expected '" + expected + "', got: " + ex.Message);
            return;
        }
        throw new Exception("Expected failure: " + expected);
    }

    static DataTable Input(params object[] values)
    {
        var table = new DataTable();
        ExcelFile.AddColumn(table, "OCR text");
        foreach (object value in values)
            table.Rows.Add(value ?? DBNull.Value);
        return table;
    }

    static SaveTextSettings Settings(string folder, string name, SaveTextFormat format)
    {
        return new SaveTextSettings { Path = Path.Combine(folder, name), Column = "A", ValueColumn = "B", Format = format };
    }

    static object Value(DataTable table, string name)
    {
        return table.Rows[0][ExcelFile.FindColumn(table, name)];
    }

    static void ContentChecks(string folder)
    {
        var input = Input("first\nsecond", Arabic, "", null);
        var s = Settings(folder, "ocr", SaveTextFormat.Text);
        string path = s.Path + ".txt";
        var plan = SaveTextWork.Run(input, s);
        Check((string)Value(plan, "Path") == path, "Text extension did not default to .txt.");
        Check((string)Value(plan, "Content") == "first\nsecond" + Environment.NewLine + Arabic +
            Environment.NewLine + Environment.NewLine, "Text rows, blanks or line breaks were lost.");
        Check(!File.Exists(path), "Preview wrote a file.");
        Check(input.Columns.Count == 1 && input.Rows.Count == 4, "Preview mutated its input.");
        var saved = SaveTextWork.Write(plan);
        Check((string)Value(saved, "Status") == "Saved" && (string)Value(plan, "Status") == "Ready", "Write mutated its cached plan.");
        Check(File.ReadAllText(path, new UTF8Encoding(false, true)) == (string)Value(plan, "Content"), "UTF-8 text changed on disk.");
        byte[] bytes = File.ReadAllBytes(path);
        Check(bytes.Length > 3 && !(bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf), "Unexpected UTF-8 BOM.");
        var replacement = SaveTextWork.Run(Input("replacement"), s);
        Check((string)Value(replacement, "Status") == "Overwrite", "Overwrite was not identified.");
        SaveTextWork.Write(replacement);
        Check(File.ReadAllText(path) == "replacement", "Existing text was not replaced.");
        Check(Directory.GetFiles(folder, "*.tmp").Length == 0, "Temporary save files were left behind.");

        s.Path = Path.Combine(folder, "custom.md");
        Check(SaveTextWork.OutputPath(s) == s.Path, "Custom extension was changed.");
        var json = new JavaScriptSerializer();
        var list = Input("quote \" slash \\ tab\t new\n", Arabic, 12, true, null, "duplicate", "duplicate");
        s = Settings(folder, "list", SaveTextFormat.JsonList);
        plan = SaveTextWork.Run(list, s);
        Check(((string)Value(plan, "Path")).EndsWith(".json"), "JSON extension did not default to .json.");
        var array = json.DeserializeObject((string)Value(plan, "Content")) as object[];
        Check(array != null && array.Length == 7, "JSON list dropped rows or duplicates.");
        Check((string)array[0] == (string)list.Rows[0][0] && (string)array[1] == Arabic, "JSON did not escape/round-trip text.");
        Check((int)array[2] == 12 && (bool)array[3] && array[4] == null, "JSON primitive types changed.");
        SaveTextWork.Write(plan);
        Check(File.ReadAllText((string)Value(plan, "Path")) == (string)Value(plan, "Content"), "JSON content changed on disk.");
        s.Path = Path.Combine(folder, "json.custom");
        Check(SaveTextWork.OutputPath(s) == s.Path, "Explicit JSON extension was changed.");

        var map = new DataTable();
        ExcelFile.AddColumn(map, "Key");
        ExcelFile.AddColumn(map, "Value");
        map.Rows.Add("one", Arabic);
        map.Rows.Add("two", "line1\nline2");
        s = Settings(folder, "map", SaveTextFormat.TextMap);
        plan = SaveTextWork.Run(map, s);
        Check((string)Value(plan, "Content") == "one\t" + Arabic + Environment.NewLine + "two\tline1\nline2", "Text map format changed.");
        s.Format = SaveTextFormat.JsonMap;
        plan = SaveTextWork.Run(map, s);
        var obj = json.DeserializeObject((string)Value(plan, "Content")) as Dictionary<string, object>;
        Check(obj != null && (string)obj["one"] == Arabic && (string)obj["two"] == "line1\nline2", "JSON map values changed.");
        map.Rows.Add("ONE", 5);
        Check(((Dictionary<string, object>)json.DeserializeObject((string)Value(SaveTextWork.Run(map, s), "Content"))).Count == 3,
            "JSON keys must be case-sensitive.");
        map.Rows.Add("one", "duplicate");
        Fails(() => SaveTextWork.Run(map, s), "duplicate key");
        map.Rows.RemoveAt(3);
        map.Rows.Add(DBNull.Value, "null key");
        Fails(() => SaveTextWork.Run(map, s), "blank/null key");

        var emptyMap = map.Clone();
        Check((string)Value(SaveTextWork.Run(emptyMap, s), "Content") == "{}", "Empty JSON map is not an object.");
        s.Format = SaveTextFormat.JsonList;
        Check((string)Value(SaveTextWork.Run(Input(), s), "Content") == "[]", "Empty JSON list is not an array.");
        s.Format = SaveTextFormat.Text;
        Check((string)Value(SaveTextWork.Run(Input(), s), "Content") == "", "Empty text is not empty.");
        Fails(() => SaveTextWork.Run(null, s), "Connect");
        Fails(() => SaveTextWork.Run(input, null), "Set up");
        Fails(() => SaveTextWork.Run(input, new SaveTextSettings { Path = s.Path, Column = "Z" }), "not in this step");
        Fails(() => SaveTextWork.Run(input, new SaveTextSettings { Column = "A" }), "Choose where");
        Fails(() => SaveTextWork.Run(input, new SaveTextSettings { Column = "A", Path = "relative.txt" }), "full output");
        Fails(() => SaveTextWork.Run(input, new SaveTextSettings { Column = "A", Path = folder }), "not a folder");
        Fails(() => SaveTextWork.Run(input, Settings(folder, "missing-folder\\out.txt", SaveTextFormat.Text)), "does not exist");
        Fails(() => SaveTextWork.Run(input, Settings(folder, "out.", SaveTextFormat.Text)), "ends with a dot");
        Fails(() => SaveTextWork.Run(input, Settings(folder, "out.txt", (SaveTextFormat)99)), "supported");
        Fails(() => SaveTextWork.Run(Input(double.NaN), Settings(folder, "nan", SaveTextFormat.JsonList)), "NaN");
        Fails(() => SaveTextWork.Run(Input(double.PositiveInfinity), Settings(folder, "infinity", SaveTextFormat.JsonList)), "Infinity");
        s = Settings(folder, "invalid-map", SaveTextFormat.JsonMap);
        s.ValueColumn = "A";
        Fails(() => SaveTextWork.Run(input, s), "different key and value");
    }

    static PipelineNode Source(PipelineDefinition definition, ValueListMode mode, params string[] keys)
    {
        var source = PipelineWork.AddNode(definition, PipelineStepKind.ValueList, 0, 0);
        var settings = new ValueListSettings { Source = ValueListSource.Stored, Mode = mode };
        foreach (string key in keys)
            settings.Items.Add(new ValueListItem { Key = key, Value = Arabic });
        source.Settings = settings;
        return source;
    }

    static PipelineNode Save(PipelineDefinition definition, PipelineNode source, string path, SaveTextFormat format)
    {
        var node = PipelineWork.AddNode(definition, PipelineStepKind.SaveText, 200, 0);
        node.Settings = new SaveTextSettings { Path = path, Format = format, Column = "Key", ValueColumn = "Value" };
        definition.Links.Add(new PipelineLink { From = source.Id, To = node.Id, Port = PipelineWork.PortIn });
        return node;
    }

    static void PipelineChecks(string folder)
    {
        var json = new JavaScriptSerializer();
        foreach (var mode in new[] { ValueListMode.List, ValueListMode.Set, ValueListMode.Map })
        {
            var definition = new PipelineDefinition();
            var source = Source(definition, mode, "one", "one", "two");
            var save = Save(definition, source, Path.Combine(folder, "pipeline-" + mode.ToString()),
                mode == ValueListMode.Map ? SaveTextFormat.JsonMap : SaveTextFormat.JsonList);
            var roundTrip = PipelineWork.Clone(definition);
            Check(roundTrip.Nodes[1].Settings is SaveTextSettings && roundTrip.Nodes[1].Kind == PipelineStepKind.SaveText,
                "Save text pipeline XML round-trip failed.");
            var cache = new Dictionary<string, DataTable>();
            var preview = PipelineWork.Run(roundTrip, cache, null, false);
            string path = PipelineWork.SaveTargetPath(save);
            Check(preview.Error == null && !File.Exists(path), "Preview failed or wrote a file: " + preview.Error);
            var run = PipelineWork.Run(roundTrip, cache, null, true);
            Check(run.Error == null && run.Written.Count == 1 && File.Exists(path), "Pipeline save failed: " + run.Error);
            object data = json.DeserializeObject(File.ReadAllText(path));
            int count = data is object[] ? ((object[])data).Length : ((Dictionary<string, object>)data).Count;
            Check(count == (mode == ValueListMode.List ? 3 : 2), "List/set/map semantics changed.");
        }

        var original = Input(Arabic);
        var settings = Settings(folder, "remap", SaveTextFormat.Text);
        var named = PipelineWork.SaveTextToNames(settings, original);
        var moved = new DataTable();
        ExcelFile.AddColumn(moved, "New first column");
        ExcelFile.AddColumn(moved, "OCR text");
        Check(named.Column == "OCR text" && settings.Column == "A", "Column name conversion mutated settings.");
        Check(PipelineWork.SaveTextToLetters(named, moved, true).Column == "B", "Export column did not survive reorder.");
        Fails(() => PipelineWork.SaveTextToLetters(named, new DataTable(), true), "OCR text");
        Check(PipelineWork.SaveTextToLetters(named, new DataTable(), false).Column == "", "Missing export column was retargeted.");
        Check(PipelineWork.IsSaveKind(PipelineStepKind.SaveText) && PipelineWork.IsFileKind(PipelineStepKind.SaveText) &&
            !PipelineWork.IsActionKind(PipelineStepKind.SaveText), "Export is not a terminal save.");

        string collision = Path.Combine(folder, "collision.txt");
        File.WriteAllText(collision, "keep original");
        var conflicting = new PipelineDefinition();
        var input = Source(conflicting, ValueListMode.List, "new contents");
        Save(conflicting, input, collision, SaveTextFormat.Text);
        Save(conflicting, input, Path.Combine(folder, ".\\COLLISION.txt"), SaveTextFormat.Text);
        string actionFile = Path.Combine(folder, "copy-source.txt");
        File.WriteAllText(actionFile, "do not copy on a failed run");
        string destination = Path.Combine(folder, "copies");
        Directory.CreateDirectory(destination);
        var actionSource = Source(conflicting, ValueListMode.List, actionFile);
        var action = PipelineWork.AddNode(conflicting, PipelineStepKind.FileAction, 200, 200);
        action.Settings = new FileActionSettings { PathColumn = "Key", Action = FileActionKind.Copy, DestFolder = destination };
        conflicting.Links.Add(new PipelineLink { From = actionSource.Id, To = action.Id, Port = PipelineWork.PortIn });
        var result = PipelineWork.Run(conflicting, new Dictionary<string, DataTable>(), null, true);
        Check(result.Error != null && result.Error.Contains("both write"), "Conflicting export destinations were accepted.");
        Check(File.ReadAllText(collision) == "keep original", "Conflicting saves replaced a file.");
        Check(!File.Exists(Path.Combine(destination, Path.GetFileName(actionFile))), "A disk action ran before collision validation.");

        var mixed = new PipelineDefinition();
        input = Source(mixed, ValueListMode.List, "one");
        string workbook = Path.Combine(folder, "mixed.xlsx");
        Save(mixed, input, workbook, SaveTextFormat.Text);
        var excel = PipelineWork.AddNode(mixed, PipelineStepKind.Save, 200, 200);
        excel.Settings = new SaveStepSettings { Path = workbook, Mode = SaveStepMode.NewSheet };
        mixed.Links.Add(new PipelineLink { From = input.Id, To = excel.Id, Port = PipelineWork.PortIn });
        result = PipelineWork.Run(mixed, new Dictionary<string, DataTable>(), null, true);
        Check(result.Error != null && result.Error.Contains("both write") && !File.Exists(workbook), "Text and Excel saves collided.");

        var failed = new PipelineDefinition();
        input = Source(failed, ValueListMode.List, "new contents");
        Save(failed, input, collision, SaveTextFormat.Text);
        PipelineWork.AddNode(failed, PipelineStepKind.LoadImage, 0, 100);
        result = PipelineWork.Run(failed, new Dictionary<string, DataTable>(), null, true);
        Check(result.Error != null && File.ReadAllText(collision) == "keep original", "A failed step did not prevent export.");
    }

    static void FailedWriteCheck(string folder)
    {
        string path = Path.Combine(folder, "locked.txt");
        File.WriteAllText(path, "keep this");
        var plan = SaveTextWork.Run(Input("replace this"), Settings(folder, "locked.txt", SaveTextFormat.Text));
        bool failed = false;
        using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            try { SaveTextWork.Write(plan); }
            catch (IOException) { failed = true; }
        }
        Check(failed && File.ReadAllText(path) == "keep this", "A failed save damaged the existing file.");
        Check(Directory.GetFiles(folder, "*.tmp").Length == 0, "Failed save left a temporary file.");
    }

    static int Main()
    {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeWorkAssistant-SaveText-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        bool passed = false;
        try
        {
            ContentChecks(folder);
            PipelineChecks(folder);
            FailedWriteCheck(folder);
            Console.WriteLine("Passed " + _checks + " text / JSON save checks.");
            passed = true;
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Console.Error.WriteLine("Test files kept in: " + folder);
            return 1;
        }
        finally
        {
            if (passed)
                Directory.Delete(folder, true);
        }
    }
}
