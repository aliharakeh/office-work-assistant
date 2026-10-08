using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Xml.Serialization;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.Templates;

namespace OfficeWorkAssistant.Features.Pipeline
{
    public enum PipelineStepKind
    {
        Load,
        FilterSort,
        Templates,
        Compare,
        FillColumns,
        Save
    }

    public enum SaveStepMode
    {
        NewFile,
        NewSheet
    }

    public sealed class LoadStepSettings
    {
        public string Path { get; set; }
        // Blank = the first sheet.
        public string Sheet { get; set; }
    }

    public sealed class SaveStepSettings
    {
        public string Path { get; set; }
        public SaveStepMode Mode { get; set; }
        public string SheetName { get; set; }

        public SaveStepSettings()
        {
            SheetName = "Result";
        }
    }

    // A saved pipeline. Feature settings inside it name columns by header, not letter,
    // so a step keeps working when an earlier step adds, drops or moves columns.
    [XmlRoot("Pipeline")]
    public sealed class PipelineDefinition
    {
        [XmlAttribute]
        public int Version { get; set; }
        public string Name { get; set; }
        public List<PipelineNode> Nodes { get; set; }
        public List<PipelineLink> Links { get; set; }

        public PipelineDefinition()
        {
            Version = 1;
            Name = "";
            Nodes = new List<PipelineNode>();
            Links = new List<PipelineLink>();
        }
    }

    public sealed class PipelineNode
    {
        [XmlAttribute]
        public string Id { get; set; }
        [XmlAttribute]
        public PipelineStepKind Kind { get; set; }
        [XmlAttribute]
        public string Title { get; set; }
        [XmlAttribute]
        public double X { get; set; }
        [XmlAttribute]
        public double Y { get; set; }

        // null = the step is not set up yet.
        [XmlElement("Load", typeof(LoadStepSettings))]
        [XmlElement("FilterSort", typeof(FilterSortSettings))]
        [XmlElement("Template", typeof(TemplateDefinition))]
        [XmlElement("Compare", typeof(ExcelProcessingSettings))]
        [XmlElement("FillColumns", typeof(FillColumnsSettings))]
        [XmlElement("Save", typeof(SaveStepSettings))]
        public object Settings { get; set; }
    }

    // The output of From feeds input Port ("In", "A" or "B") of To.
    public sealed class PipelineLink
    {
        [XmlAttribute]
        public string From { get; set; }
        [XmlAttribute]
        public string To { get; set; }
        [XmlAttribute]
        public string Port { get; set; }
    }

    public sealed class PipelineRunResult
    {
        public string FailedNodeId { get; set; }
        public string Error { get; set; }
        public List<string> Written { get; set; }

        public PipelineRunResult()
        {
            Written = new List<string>();
        }
    }

    // Pure logic: graph, XML, running steps. No WPF.
    public static class PipelineWork
    {
        public const string PortIn = "In";
        public const string PortA = "A";
        public const string PortB = "B";

        static readonly string[] NoPorts = new string[0];
        static readonly string[] OnePort = { PortIn };
        static readonly string[] TwoPorts = { PortA, PortB };

        public static string[] Ports(PipelineStepKind kind)
        {
            if (kind == PipelineStepKind.Load)
                return NoPorts;
            if (kind == PipelineStepKind.Compare || kind == PipelineStepKind.FillColumns)
                return TwoPorts;
            return OnePort;
        }

        public static string KindLabel(PipelineStepKind kind)
        {
            switch (kind)
            {
                case PipelineStepKind.Load: return "Load file";
                case PipelineStepKind.FilterSort: return "Filter & Sort";
                case PipelineStepKind.Templates: return "Template";
                case PipelineStepKind.Compare: return "Compare A/B";
                case PipelineStepKind.FillColumns: return "Fill columns";
                default: return "Save file";
            }
        }

        public static string PortLabel(PipelineStepKind kind, string port)
        {
            if (kind == PipelineStepKind.FillColumns)
                return port == PortA ? "A (take values from)" : "B (fill into)";
            if (kind == PipelineStepKind.Compare)
                return port == PortA ? "File A" : "File B";
            return "input";
        }

        public static PipelineNode Find(PipelineDefinition def, string id)
        {
            foreach (var node in def.Nodes)
            {
                if (node.Id == id)
                    return node;
            }
            return null;
        }

        public static PipelineNode AddNode(PipelineDefinition def, PipelineStepKind kind, double x, double y)
        {
            var count = 1;
            foreach (var n in def.Nodes)
            {
                if (n.Kind == kind)
                    count++;
            }
            var node = new PipelineNode
            {
                Id = NextId(def),
                Kind = kind,
                Title = KindLabel(kind) + " " + count.ToString(CultureInfo.InvariantCulture),
                X = x,
                Y = y
            };
            if (kind == PipelineStepKind.Load)
                node.Settings = new LoadStepSettings();
            else if (kind == PipelineStepKind.Save)
                node.Settings = new SaveStepSettings();
            def.Nodes.Add(node);
            return node;
        }

        static string NextId(PipelineDefinition def)
        {
            var max = 0;
            foreach (var n in def.Nodes)
            {
                int v;
                if (n.Id != null && n.Id.Length > 1 && n.Id[0] == 'n' &&
                    int.TryParse(n.Id.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v > max)
                    max = v;
            }
            return "n" + (max + 1).ToString(CultureInfo.InvariantCulture);
        }

        public static void RemoveNode(PipelineDefinition def, string id)
        {
            def.Nodes.RemoveAll(n => n.Id == id);
            def.Links.RemoveAll(l => l.From == id || l.To == id);
        }

        public static PipelineLink InputLink(PipelineDefinition def, string nodeId, string port)
        {
            foreach (var link in def.Links)
            {
                if (link.To == nodeId && link.Port == port)
                    return link;
            }
            return null;
        }

        // A new link from -> to makes a loop when from is already downstream of to (or is to).
        public static bool WouldCycle(PipelineDefinition def, string from, string to)
        {
            return Downstream(def, to).Contains(from);
        }

        // The step and every step that reads its result, directly or further down.
        public static HashSet<string> Downstream(PipelineDefinition def, string id)
        {
            return Walk(def, id, true);
        }

        // The step and every step it needs a result from.
        public static HashSet<string> Upstream(PipelineDefinition def, string id)
        {
            return Walk(def, id, false);
        }

        static HashSet<string> Walk(PipelineDefinition def, string id, bool down)
        {
            var seen = new HashSet<string>();
            var todo = new Stack<string>();
            todo.Push(id);
            while (todo.Count > 0)
            {
                var at = todo.Pop();
                if (!seen.Add(at))
                    continue;
                foreach (var link in def.Links)
                {
                    if (down && link.From == at)
                        todo.Push(link.To);
                    else if (!down && link.To == at)
                        todo.Push(link.From);
                }
            }
            return seen;
        }

        // Steps in an order where every step comes after the steps it reads from.
        public static List<PipelineNode> Order(PipelineDefinition def)
        {
            var waiting = new Dictionary<string, int>();
            foreach (var node in def.Nodes)
                waiting[node.Id] = 0;
            foreach (var link in def.Links)
            {
                if (waiting.ContainsKey(link.To))
                    waiting[link.To]++;
            }

            var ready = new Queue<PipelineNode>();
            foreach (var node in def.Nodes)
            {
                if (waiting[node.Id] == 0)
                    ready.Enqueue(node);
            }

            var order = new List<PipelineNode>();
            while (ready.Count > 0)
            {
                var node = ready.Dequeue();
                order.Add(node);
                foreach (var link in def.Links)
                {
                    if (link.From != node.Id || !waiting.ContainsKey(link.To))
                        continue;
                    if (--waiting[link.To] == 0)
                        ready.Enqueue(Find(def, link.To));
                }
            }
            if (order.Count != def.Nodes.Count)
                throw new InvalidOperationException("The steps form a loop. Remove one of the links.");
            return order;
        }

        // ---------- files ----------

        public static void SaveFile(PipelineDefinition def, string path)
        {
            var ser = new XmlSerializer(typeof(PipelineDefinition));
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                ser.Serialize(fs, def);
        }

        public static PipelineDefinition LoadFile(string path)
        {
            PipelineDefinition def;
            var ser = new XmlSerializer(typeof(PipelineDefinition));
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                try
                {
                    def = ser.Deserialize(fs) as PipelineDefinition;
                }
                catch (InvalidOperationException ex)
                {
                    throw new InvalidOperationException("Not a pipeline file. " + (ex.InnerException != null ? ex.InnerException.Message : ex.Message), ex);
                }
            }
            if (def == null)
                throw new InvalidOperationException("Not a pipeline file.");
            if (def.Nodes == null)
                def.Nodes = new List<PipelineNode>();
            if (def.Links == null)
                def.Links = new List<PipelineLink>();

            // Drop anything that does not fit, rather than refuse the whole file.
            var ids = new HashSet<string>();
            def.Nodes.RemoveAll(n => n == null || string.IsNullOrEmpty(n.Id) || !ids.Add(n.Id));
            foreach (var node in def.Nodes)
            {
                if (node.Title == null)
                    node.Title = KindLabel(node.Kind);
                if (!SettingsFit(node.Kind, node.Settings))
                    node.Settings = null;
                if (node.Settings == null && node.Kind == PipelineStepKind.Load)
                    node.Settings = new LoadStepSettings();
                if (node.Settings == null && node.Kind == PipelineStepKind.Save)
                    node.Settings = new SaveStepSettings();
            }
            var ports = new HashSet<string>();
            def.Links.RemoveAll(l => l == null || !ids.Contains(l.From) || !ids.Contains(l.To) ||
                Array.IndexOf(Ports(Find(def, l.To).Kind), l.Port) < 0 || !ports.Add(l.To + "|" + l.Port));
            Order(def); // throws on a loop
            return def;
        }

        static bool SettingsFit(PipelineStepKind kind, object settings)
        {
            if (settings == null)
                return true;
            switch (kind)
            {
                case PipelineStepKind.Load: return settings is LoadStepSettings;
                case PipelineStepKind.FilterSort: return settings is FilterSortSettings;
                case PipelineStepKind.Templates: return settings is TemplateDefinition;
                case PipelineStepKind.Compare: return settings is ExcelProcessingSettings;
                case PipelineStepKind.FillColumns: return settings is FillColumnsSettings;
                default: return settings is SaveStepSettings;
            }
        }

        // Deep copy through XML, so an edit that gets cancelled never touches the stored settings.
        public static T Clone<T>(T value) where T : class
        {
            if (value == null)
                return null;
            var ser = new XmlSerializer(value.GetType());
            using (var ms = new MemoryStream())
            {
                ser.Serialize(ms, value);
                ms.Position = 0;
                return (T)ser.Deserialize(ms);
            }
        }

        // One line for the step box on the canvas.
        public static string Summary(PipelineNode node)
        {
            var s = node.Settings;
            if (s == null)
                return "Not set up. Double-click to edit.";
            var load = s as LoadStepSettings;
            if (load != null)
            {
                if (string.IsNullOrWhiteSpace(load.Path))
                    return "Choose a file.";
                return Path.GetFileName(load.Path) + (string.IsNullOrEmpty(load.Sheet) ? "" : " / " + load.Sheet);
            }
            var save = s as SaveStepSettings;
            if (save != null)
            {
                if (string.IsNullOrWhiteSpace(save.Path))
                    return "Choose where to save.";
                return Path.GetFileName(save.Path) + " / " + (save.SheetName ?? "") +
                    (save.Mode == SaveStepMode.NewSheet ? " (new sheet)" : "");
            }
            var filter = s as FilterSortSettings;
            if (filter != null)
                return Count(filter.Conditions == null ? 0 : filter.Conditions.Count, "filter") + ", " +
                    Count(filter.Sorts == null ? 0 : filter.Sorts.Count, "sort") +
                    (filter.KeepAll ? "" : ", " + Count(filter.Keep == null ? 0 : filter.Keep.Count, "column"));
            var template = s as TemplateDefinition;
            if (template != null)
                return Count(template.Columns == null ? 0 : template.Columns.Count, "new column") +
                    (string.IsNullOrWhiteSpace(template.Name) ? "" : " (" + template.Name + ")");
            var compare = s as ExcelProcessingSettings;
            if (compare != null)
            {
                var op = compare.Operation == "AOnly" ? "Only in A" : compare.Operation == "BOnly" ? "Only in B" : "Common rows";
                return op + ", " + Count(compare.Matches == null ? 0 : compare.Matches.Count, "match rule");
            }
            var fill = s as FillColumnsSettings;
            if (fill != null)
                return Count(fill.Keys == null ? 0 : fill.Keys.Count, "match rule") + ", " +
                    Count(fill.Rules == null ? 0 : fill.Rules.Count, "fill rule");
            return "";
        }

        static string Count(int n, string noun)
        {
            return n.ToString(CultureInfo.InvariantCulture) + " " + noun + (n == 1 ? "" : "s");
        }

        // ---------- running ----------

        // Runs every step (targetId null) or just what targetId needs, in order. Results land in
        // cache by step id; steps already in cache are reused.
        // writeFiles: once every step has worked, write the Save steps' files. Nothing is
        // written when a step fails. Otherwise Save steps just pass their input through.
        public static PipelineRunResult Run(PipelineDefinition def, IDictionary<string, DataTable> cache,
            string targetId, bool writeFiles)
        {
            var result = new PipelineRunResult();
            List<PipelineNode> order;
            try
            {
                order = Order(def);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
                return result;
            }

            var wanted = targetId == null ? null : Upstream(def, targetId);
            var saves = new List<PipelineNode>();
            foreach (var node in order)
            {
                if (wanted != null && !wanted.Contains(node.Id))
                    continue;
                if (node.Kind == PipelineStepKind.Save)
                    saves.Add(node);
                if (cache.ContainsKey(node.Id))
                    continue;
                try
                {
                    cache[node.Id] = RunNode(def, node, cache);
                }
                catch (Exception ex)
                {
                    cache.Remove(node.Id);
                    result.FailedNodeId = node.Id;
                    result.Error = ex.Message;
                    return result;
                }
            }

            if (writeFiles)
                WriteSaves(saves, cache, result);
            return result;
        }

        // New files first, then new sheets, so a sheet added to a workbook that another
        // step creates in the same run is not wiped out when that workbook is written.
        static void WriteSaves(List<PipelineNode> saves, IDictionary<string, DataTable> cache, PipelineRunResult result)
        {
            var created = new Dictionary<string, PipelineNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var node in saves)
            {
                var s = (SaveStepSettings)node.Settings;
                if (s.Mode != SaveStepMode.NewFile)
                    continue;
                var path = System.IO.Path.GetFullPath(s.Path);
                PipelineNode other;
                if (created.TryGetValue(path, out other))
                {
                    result.FailedNodeId = node.Id;
                    result.Error = "This step and '" + other.Title + "' both replace " + System.IO.Path.GetFileName(path) +
                        ". Make one of them add a new sheet instead.";
                    return;
                }
                created[path] = node;
            }

            var ordered = new List<PipelineNode>();
            foreach (var node in saves)
            {
                if (((SaveStepSettings)node.Settings).Mode == SaveStepMode.NewFile)
                    ordered.Add(node);
            }
            foreach (var node in saves)
            {
                if (((SaveStepSettings)node.Settings).Mode == SaveStepMode.NewSheet)
                    ordered.Add(node);
            }

            foreach (var node in ordered)
            {
                var s = (SaveStepSettings)node.Settings;
                try
                {
                    // A new sheet needs a workbook to go into; create it the first time.
                    bool append = s.Mode == SaveStepMode.NewSheet && File.Exists(s.Path);
                    ExcelFile.Save(cache[node.Id], s.Path, s.SheetName, append);
                    result.Written.Add(s.Path + " (sheet " + s.SheetName.Trim() + ")");
                }
                catch (Exception ex)
                {
                    result.FailedNodeId = node.Id;
                    result.Error = ex.Message;
                    return;
                }
            }
        }

        static DataTable RunNode(PipelineDefinition def, PipelineNode node, IDictionary<string, DataTable> cache)
        {
            if (node.Settings == null)
                throw new InvalidOperationException("Set up this step first: double-click it.");

            switch (node.Kind)
            {
                case PipelineStepKind.Load:
                {
                    var s = (LoadStepSettings)node.Settings;
                    if (string.IsNullOrWhiteSpace(s.Path))
                        throw new InvalidOperationException("Choose a file to load.");
                    return ExcelFile.Load(s.Path, string.IsNullOrEmpty(s.Sheet) ? null : s.Sheet, 0).Table;
                }
                case PipelineStepKind.FilterSort:
                {
                    var input = Input(def, node, PortIn, cache);
                    return FilterSortWork.Run(input, FilterSortToLetters((FilterSortSettings)node.Settings, input, true));
                }
                case PipelineStepKind.Templates:
                {
                    var input = Input(def, node, PortIn, cache);
                    return TemplatesWork.ApplyTemplate(input, TemplateToLetters((TemplateDefinition)node.Settings, input, true));
                }
                case PipelineStepKind.Compare:
                    return ExcelProcessingWork.Run(Input(def, node, PortA, cache), Input(def, node, PortB, cache),
                        (ExcelProcessingSettings)node.Settings);
                case PipelineStepKind.FillColumns:
                    return FillColumnsWork.Run(Input(def, node, PortA, cache), Input(def, node, PortB, cache),
                        (FillColumnsSettings)node.Settings);
                default:
                {
                    // Files are written after every step worked; see WriteSaves.
                    var input = Input(def, node, PortIn, cache);
                    string error = CheckSave((SaveStepSettings)node.Settings);
                    if (error != null)
                        throw new InvalidOperationException(error);
                    return input;
                }
            }
        }

        public static string CheckSave(SaveStepSettings s)
        {
            if (s == null || string.IsNullOrWhiteSpace(s.Path))
                return "Choose where to save.";
            // Writing a .xlsm through ClosedXML drops its macros, same rule as the save dialog.
            if (!string.Equals(Path.GetExtension(s.Path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                return "Save to a .xlsx file.";
            return ExcelFile.CheckSheetName(s.SheetName);
        }

        static DataTable Input(PipelineDefinition def, PipelineNode node, string port, IDictionary<string, DataTable> cache)
        {
            var link = InputLink(def, node.Id, port);
            if (link == null)
                throw new InvalidOperationException("Connect " + PortLabel(node.Kind, port) + " first.");
            DataTable table;
            if (!cache.TryGetValue(link.From, out table))
                throw new InvalidOperationException("The step before this one has no result yet.");
            return table;
        }

        // ---------- column names <-> letters ----------
        // Pages and Work store letters ("A" = 1st column). A pipeline stores header names instead,
        // and converts at the edges against the step's current input.

        public static FilterSortSettings FilterSortToNames(FilterSortSettings settings, DataTable input)
        {
            var s = Clone(settings);
            for (var i = 0; i < s.Keep.Count; i++)
                s.Keep[i] = NameOf(input, s.Keep[i]);
            foreach (var c in s.Conditions)
            {
                if (c.Kind == FilterConditionKind.Column)
                    c.Column = NameOf(input, c.Column);
            }
            foreach (var k in s.Sorts)
                k.Column = NameOf(input, k.Column);
            return s;
        }

        // strict: a missing column is an error (running). Not strict: it is dropped,
        // or left as a name the page cannot select (editing), so the user can fix it.
        public static FilterSortSettings FilterSortToLetters(FilterSortSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Keep = Letters(input, s.Keep, strict);
            foreach (var c in s.Conditions)
            {
                if (c.Kind == FilterConditionKind.Column)
                    c.Column = LetterOf(input, c.Column, strict) ?? c.Column;
            }
            s.Sorts.RemoveAll(k =>
            {
                var letter = LetterOf(input, k.Column, strict);
                if (letter == null)
                    return true;
                k.Column = letter;
                return false;
            });
            return s;
        }

        public static TemplateDefinition TemplateToNames(TemplateDefinition template, DataTable input)
        {
            var t = Clone(template);
            for (var i = 0; i < t.SourceColumns.Count; i++)
                t.SourceColumns[i] = NameOf(input, t.SourceColumns[i]);
            foreach (var c in t.Columns)
            {
                if (c.Kind == TemplateColumnKind.Copy)
                    c.SourceColumn = NameOf(input, c.SourceColumn);
            }
            return t;
        }

        public static TemplateDefinition TemplateToLetters(TemplateDefinition template, DataTable input, bool strict)
        {
            var t = Clone(template);
            var count = t.SourceColumns.Count;
            t.SourceColumns = Letters(input, t.SourceColumns, strict);
            // Every kept column is gone: keep none, rather than turn into "keep all".
            if (count > 0 && t.SourceColumns.Count == 0)
                t.KeepSourceColumns = false;
            foreach (var c in t.Columns)
            {
                if (c.Kind == TemplateColumnKind.Copy)
                    c.SourceColumn = LetterOf(input, c.SourceColumn, strict) ?? c.SourceColumn;
            }
            return t;
        }

        static List<string> Letters(DataTable input, List<string> names, bool strict)
        {
            var list = new List<string>();
            foreach (var name in names)
            {
                var letter = LetterOf(input, name, strict);
                if (letter != null)
                    list.Add(letter);
            }
            return list;
        }

        static string NameOf(DataTable input, string letter)
        {
            int index;
            if (FilterSortWork.TryParseColumnLetter(FilterSortWork.ExtractLetter(letter), out index) &&
                index < input.Columns.Count)
                return ExcelFile.Header(input.Columns[index]);
            return letter;
        }

        static string LetterOf(DataTable input, string name, bool strict)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;
            var want = name.Trim();
            // The real header wins over the binding id, so "Qty" never lands on "Qty_2".
            for (var i = 0; i < input.Columns.Count; i++)
            {
                if (string.Equals(ExcelFile.Header(input.Columns[i]), want, StringComparison.OrdinalIgnoreCase))
                    return FilterSortWork.ColumnLetter(i);
            }
            var col = ExcelFile.FindColumn(input, want);
            if (col != null)
                return FilterSortWork.ColumnLetter(col.Ordinal);
            if (strict)
                throw new InvalidOperationException("Column '" + want + "' is not in this step's input.");
            return null;
        }
    }
}
