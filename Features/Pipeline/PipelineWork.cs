using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Xml.Serialization;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.Append;
using OfficeWorkAssistant.Features.ArrangeColumns;
using OfficeWorkAssistant.Features.EditHeaders;
using OfficeWorkAssistant.Features.CleanText;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.FormatSheet;
using OfficeWorkAssistant.Features.Highlight;
using OfficeWorkAssistant.Features.MergeDuplicates;
using OfficeWorkAssistant.Features.SaveText;
using OfficeWorkAssistant.Features.Pdf;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.RemoveDuplicates;
using OfficeWorkAssistant.Features.SplitCombine;
using OfficeWorkAssistant.Features.Templates;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // Saved by name, so new kinds are appended. IsFileKind / IsStyleKind list their kinds.
    public enum PipelineStepKind
    {
        Load,
        FilterSort,
        Templates,
        Compare,
        FillColumns,
        Save,
        ValueList,
        ListFolder,
        FindFiles,
        FileAction,
        MergeFolders,
        CleanText,
        SplitColumn,
        CombineColumns,
        RemoveDuplicates,
        ArrangeColumns,
        Append,
        Highlight,
        FormatSheet,
        SaveText,
        EditHeaders,
        LoadPdf,
        ArrangePdf,
        MergePdf,
        SplitPdf,
        SavePdf
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
        public bool FirstRowIsData { get; set; }
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

        [XmlAttribute]
        public bool EditMatchRulesInRunMode { get; set; }

        [XmlAttribute]
        public bool ShowInRunMode { get; set; }

        // null = the step is not set up yet.
        [XmlElement("Load", typeof(LoadStepSettings))]
        [XmlElement("FilterSort", typeof(FilterSortSettings))]
        [XmlElement("Template", typeof(TemplateDefinition))]
        [XmlElement("Compare", typeof(ExcelProcessingSettings))]
        [XmlElement("FillColumns", typeof(FillColumnsSettings))]
        [XmlElement("Save", typeof(SaveStepSettings))]
        [XmlElement("ValueList", typeof(ValueListSettings))]
        [XmlElement("ListFolder", typeof(FolderScanSettings))]
        [XmlElement("FindFiles", typeof(FindFilesSettings))]
        [XmlElement("FileAction", typeof(FileActionSettings))]
        [XmlElement("MergeFolders", typeof(MergeFoldersSettings))]
        [XmlElement("CleanText", typeof(CleanTextSettings))]
        [XmlElement("SplitColumn", typeof(SplitColumnSettings))]
        [XmlElement("CombineColumns", typeof(CombineColumnsSettings))]
        [XmlElement("RemoveDuplicates", typeof(RemoveDuplicatesSettings))]
        [XmlElement("ArrangeColumns", typeof(ArrangeColumnsSettings))]
        [XmlElement("EditHeaders", typeof(EditHeadersSettings))]
        [XmlElement("Append", typeof(AppendSettings))]
        [XmlElement("Highlight", typeof(HighlightSettings))]
        [XmlElement("FormatSheet", typeof(FormatSheetSettings))]
        [XmlElement("SaveText", typeof(SaveTextSettings))]
        [XmlElement("LoadPdf", typeof(LoadPdfSettings))]
        [XmlElement("ArrangePdf", typeof(ArrangePdfSettings))]
        [XmlElement("MergePdf", typeof(MergePdfSettings))]
        [XmlElement("SplitPdf", typeof(SplitPdfSettings))]
        [XmlElement("SavePdf", typeof(SavePdfSettings))]
        public object Settings { get; set; }
    }

        // The output of From feeds one input port of To.
    public sealed class PipelineLink
    {
        [XmlAttribute]
        public string Id { get; set; }
        [XmlAttribute]
        public string From { get; set; }
        [XmlAttribute]
        public string FromPort { get; set; }
        [XmlAttribute]
        public string To { get; set; }
        [XmlAttribute]
        public string Port { get; set; }

        public PipelineLink()
        {
            Id = Guid.NewGuid().ToString("N");
        }
    }

    public sealed class PipelineRunResult
    {
        public string FailedNodeId { get; set; }
        public string Error { get; set; }
        public List<string> Written { get; set; }
        // One line per File action step that ran: what happened to its items.
        public List<string> Actions { get; set; }

        public PipelineRunResult()
        {
            Written = new List<string>();
            Actions = new List<string>();
        }
    }

    // Pure logic: graph, XML, running steps. No WPF.
    public static partial class PipelineWork
    {
        public const string PortIn = "In";
        public const string PortA = "A";
        public const string PortB = "B";

        static readonly string[] NoPorts = new string[0];
        static readonly string[] OnePort = { PortIn };
        static readonly string[] TwoPorts = { PortA, PortB };

        // The ports a kind starts with.
        public static string[] Ports(PipelineStepKind kind)
        {
            if (kind == PipelineStepKind.Load || kind == PipelineStepKind.ListFolder || kind == PipelineStepKind.MergeFolders)
                return NoPorts;
            if (kind == PipelineStepKind.Compare || kind == PipelineStepKind.FillColumns || kind == PipelineStepKind.ArrangePdf)
                return TwoPorts;
            if (CombinesInputs(kind))
                return new[] { AppendPort(1) };
            return OnePort;
        }

        // Variable-input steps have one port per link plus a free one.
        public static string[] InputPorts(PipelineDefinition def, PipelineNode node)
        {
            var roles = Ports(node.Kind);
            if (roles.Length == 0)
                return roles;
            var ports = new List<string>(roles);
            foreach (var role in roles)
            {
                var links = BatchLinks(def, node, role);
                for (var i = 2; i <= links.Count + 1; i++)
                    ports.Add(VariablePort(node.Kind, role, i));
            }
            return ports.ToArray();
        }

        public static bool HasVariableInputs(PipelineStepKind kind)
        {
            return Ports(kind).Length != 0;
        }

        static bool CombinesInputs(PipelineStepKind kind)
        {
            return kind == PipelineStepKind.Append || kind == PipelineStepKind.MergePdf;
        }

        public static bool IsOptionalInput(PipelineStepKind kind, string port)
        {
            return kind == PipelineStepKind.ValueList || CombinesInputs(kind) ||
                (port != PortIn && port != PortA && port != PortB) ||
                kind == PipelineStepKind.LoadPdf || (kind == PipelineStepKind.ArrangePdf && port == PortB);
        }

        static string AppendPort(int n)
        {
            return "In" + n.ToString(CultureInfo.InvariantCulture);
        }

        // 1 for "In1", ...; 0 when the name is not a numbered input.
        public static int AppendPortNumber(string port)
        {
            return NumberedPort(port, PortIn);
        }

        static int NumberedPort(string port, string role)
        {
            int number;
            return port != null && port.StartsWith(role, StringComparison.Ordinal) &&
                int.TryParse(port.Substring(role.Length), NumberStyles.None, CultureInfo.InvariantCulture, out number) &&
                number > 0 ? number : 0;
        }

        // Renumber variable inputs with no gaps, keeping their order.
        // True when a port changed.
        public static bool CompactPorts(PipelineDefinition def, string nodeId)
        {
            var node = Find(def, nodeId);
            if (node == null || !HasVariableInputs(node.Kind))
                return false;
            var changed = false;
            foreach (var role in Ports(node.Kind))
            {
                var links = BatchLinks(def, node, role);
                for (var i = 0; i < links.Count; i++)
                {
                    var port = VariablePort(node.Kind, role, i + 1);
                    if (links[i].Port != port)
                    {
                        links[i].Port = port;
                        changed = true;
                    }
                }
            }
            return changed;
        }

        public static string KindLabel(PipelineStepKind kind)
        {
            switch (kind)
            {
                case PipelineStepKind.Load: return "Load file";
                case PipelineStepKind.FilterSort: return "Filter & Sort";
                case PipelineStepKind.Templates: return "Add columns";
                case PipelineStepKind.Compare: return "Compare A/B";
                case PipelineStepKind.FillColumns: return "Fill columns";
                case PipelineStepKind.ValueList: return "List / set / map";
                case PipelineStepKind.ListFolder: return "List folder";
                case PipelineStepKind.FindFiles: return "Find files";
                case PipelineStepKind.FileAction: return "File action";
                case PipelineStepKind.MergeFolders: return "Merge folders";
                case PipelineStepKind.CleanText: return "Clean text";
                case PipelineStepKind.SplitColumn: return "Split column";
                case PipelineStepKind.CombineColumns: return "Combine columns";
                case PipelineStepKind.RemoveDuplicates: return "Remove duplicates";
                case PipelineStepKind.ArrangeColumns: return "Arrange columns";
                case PipelineStepKind.EditHeaders: return "Edit headers";
                case PipelineStepKind.Append: return "Append tables";
                case PipelineStepKind.Highlight: return "Highlight";
                case PipelineStepKind.FormatSheet: return "Format sheet";
                case PipelineStepKind.SaveText: return "Save text / JSON";
                case PipelineStepKind.LoadPdf: return "Load PDF";
                case PipelineStepKind.ArrangePdf: return "Arrange PDF pages";
                case PipelineStepKind.MergePdf: return "Merge PDFs";
                case PipelineStepKind.SplitPdf: return "Split PDF";
                case PipelineStepKind.SavePdf: return "Save PDF";
                default: return "Save file";
            }
        }

        // One line for the step palette.
        public static string KindDescription(PipelineStepKind kind)
        {
            switch (kind)
            {
                case PipelineStepKind.Load: return "Read a sheet of an Excel file.";
                case PipelineStepKind.FilterSort: return "Keep rows that match, sort, pick columns.";
                case PipelineStepKind.Templates: return "Add columns built from formulas.";
                case PipelineStepKind.Compare: return "Rows only in A, only in B, or in both.";
                case PipelineStepKind.FillColumns: return "Copy values from A into matching rows of B.";
                case PipelineStepKind.ValueList: return "Keys (and values) from rows, or typed in.";
                case PipelineStepKind.ListFolder: return "Every file and/or folder in a folder.";
                case PipelineStepKind.FindFiles: return "Files and folders that match the keys.";
                case PipelineStepKind.FileAction: return "Copy, move or delete the listed paths.";
                case PipelineStepKind.MergeFolders: return "Merge subfolders that belong together.";
                case PipelineStepKind.CleanText: return "Trim spaces, fix case, remove hidden characters.";
                case PipelineStepKind.SplitColumn: return "Cut one column into several.";
                case PipelineStepKind.CombineColumns: return "Join several columns into one.";
                case PipelineStepKind.RemoveDuplicates: return "Keep one row per key, or show the repeats.";
                case PipelineStepKind.ArrangeColumns: return "Pick, reorder and rename columns.";
                case PipelineStepKind.EditHeaders: return "Add, clear or rename headers with text or formulas.";
                case PipelineStepKind.Append: return "Stack the rows of several tables.";
                case PipelineStepKind.Highlight: return "Colour cells or rows that match a condition.";
                case PipelineStepKind.FormatSheet: return "Style the sheet, header, data rows, and columns. Add filters or freeze panes.";
                case PipelineStepKind.SaveText: return "Save column text, lists, sets or maps to a text / JSON file.";
                case PipelineStepKind.LoadPdf: return "Read one PDF file or a column of PDF paths.";
                case PipelineStepKind.ArrangePdf: return "Select, reorder, rotate, duplicate or insert pages.";
                case PipelineStepKind.MergePdf: return "Combine PDF documents in input order or interleave their pages.";
                case PipelineStepKind.SplitPdf: return "Divide PDF documents into groups of pages.";
                case PipelineStepKind.SavePdf: return "Write the PDF documents to disk.";
                default: return "Write the result to an Excel file.";
            }
        }

        // Steps that work on files and folders rather than Excel data.
        public static bool IsFileKind(PipelineStepKind kind)
        {
            switch (kind)
            {
                case PipelineStepKind.ValueList:
                case PipelineStepKind.ListFolder:
                case PipelineStepKind.FindFiles:
                case PipelineStepKind.FileAction:
                case PipelineStepKind.MergeFolders:
                case PipelineStepKind.SaveText:
                case PipelineStepKind.LoadPdf:
                case PipelineStepKind.ArrangePdf:
                case PipelineStepKind.MergePdf:
                case PipelineStepKind.SplitPdf:
                case PipelineStepKind.SavePdf:
                    return true;
                default:
                    return false;
            }
        }

        // Terminal saves write only after all computations and file actions have succeeded.
        public static bool IsSaveKind(PipelineStepKind kind)
        {
            return kind == PipelineStepKind.Save || kind == PipelineStepKind.SaveText || kind == PipelineStepKind.SavePdf;
        }

        // Steps that style the table for Save. Their result keeps the style; every other step drops it.
        public static bool IsStyleKind(PipelineStepKind kind)
        {
            return kind == PipelineStepKind.Highlight || kind == PipelineStepKind.FormatSheet;
        }

        // Steps that change the disk. They only plan until Run and save files.
        public static bool IsActionKind(PipelineStepKind kind)
        {
            return kind == PipelineStepKind.FileAction || kind == PipelineStepKind.MergeFolders;
        }

        // How many items an action step's plan would touch, in words.
        public static string DescribeAction(PipelineNode node, DataTable plan)
        {
            var merge = node.Settings as MergeFoldersSettings;
            if (merge != null)
                return "Merge " + Count(MergeDuplicatesWork.CountReady(plan), "folder") + " in " + merge.Folder;
            var count = 0;
            if (plan != null)
                foreach (var item in ResultItems(plan))
                    count += FileOpsWork.CountReady(item.Table);
            return FileOpsWork.Describe((FileActionSettings)node.Settings, count);
        }

        public static string PortLabel(PipelineStepKind kind, string port)
        {
            if (Ports(kind) == TwoPorts)
            {
                int a = NumberedPort(port, PortA);
                int b = NumberedPort(port, PortB);
                if (a > 1 || b > 1)
                    return PortLabel(kind, a > 1 ? PortA : PortB) + " " + Math.Max(a, b).ToString(CultureInfo.InvariantCulture);
            }
            if (kind == PipelineStepKind.FillColumns)
                return port == PortA ? "A (take values from)" : "B (fill into)";
            if (kind == PipelineStepKind.Compare)
                return port == PortA ? "File A" : "File B";
            if (kind == PipelineStepKind.FindFiles)
                return "keys input";
            if (kind == PipelineStepKind.ArrangePdf)
                return port == PortA ? "A (main PDF)" : "B (insert pages)";
            if (kind == PipelineStepKind.LoadPdf)
                return "optional PDF paths";
            if (kind == PipelineStepKind.MergePdf)
                return "PDF input " + AppendPortNumber(port).ToString(CultureInfo.InvariantCulture);
            if (kind == PipelineStepKind.Append)
                return "table " + AppendPortNumber(port).ToString(CultureInfo.InvariantCulture);
            if (HasVariableInputs(kind) && port != PortIn)
                return "input " + AppendPortNumber(port).ToString(CultureInfo.InvariantCulture);
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
            // The first number not used by a step of this kind, so names stay different after a step is deleted.
            var count = 1;
            while (def.Nodes.Exists(n => n.Kind == kind && n.Title == KindLabel(kind) + " " + count.ToString(CultureInfo.InvariantCulture)))
                count++;
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
            else if (kind == PipelineStepKind.LoadPdf)
                node.Settings = new LoadPdfSettings();
            else if (kind == PipelineStepKind.ArrangePdf)
                node.Settings = new ArrangePdfSettings();
            else if (kind == PipelineStepKind.MergePdf)
                node.Settings = new MergePdfSettings();
            else if (kind == PipelineStepKind.SplitPdf)
                node.Settings = new SplitPdfSettings();
            else if (kind == PipelineStepKind.SavePdf)
                node.Settings = new SavePdfSettings();
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

        // Returns the variable-input steps that lost an input; their ports are numbered again.
        public static List<string> RemoveNode(PipelineDefinition def, string id)
        {
            var fed = new List<string>();
            foreach (var l in def.Links)
            {
                var to = Find(def, l.To);
                if (l.From == id && l.To != id && to != null && HasVariableInputs(to.Kind) && !fed.Contains(l.To))
                    fed.Add(l.To);
            }
            def.Nodes.RemoveAll(n => n.Id == id);
            def.Links.RemoveAll(l => l.From == id || l.To == id);
            foreach (var to in fed)
                CompactPorts(def, to);
            return fed;
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
                !PortFits(Find(def, l.To).Kind, l.Port) || !ports.Add(l.To + "|" + l.Port));
            var linkIds = new HashSet<string>();
            foreach (var link in def.Links)
                if (string.IsNullOrEmpty(link.Id) || link.Id.Contains("/") || !linkIds.Add(link.Id))
                {
                    link.Id = Guid.NewGuid().ToString("N");
                    linkIds.Add(link.Id);
                }
            foreach (var node in def.Nodes)
                CompactPorts(def, node.Id);
            Order(def); // throws on a loop
            return def;
        }

        static bool PortFits(PipelineStepKind kind, string port)
        {
            foreach (var role in Ports(kind))
                if (InputNumber(kind, port, role) > 0)
                    return true;
            return false;
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
                case PipelineStepKind.ValueList: return settings is ValueListSettings;
                case PipelineStepKind.ListFolder: return settings is FolderScanSettings;
                case PipelineStepKind.FindFiles: return settings is FindFilesSettings;
                case PipelineStepKind.FileAction: return settings is FileActionSettings;
                case PipelineStepKind.MergeFolders: return settings is MergeFoldersSettings;
                case PipelineStepKind.CleanText: return settings is CleanTextSettings;
                case PipelineStepKind.SplitColumn: return settings is SplitColumnSettings;
                case PipelineStepKind.CombineColumns: return settings is CombineColumnsSettings;
                case PipelineStepKind.RemoveDuplicates: return settings is RemoveDuplicatesSettings;
                case PipelineStepKind.ArrangeColumns: return settings is ArrangeColumnsSettings;
                case PipelineStepKind.EditHeaders: return settings is EditHeadersSettings;
                case PipelineStepKind.Append: return settings is AppendSettings;
                case PipelineStepKind.Highlight: return settings is HighlightSettings;
                case PipelineStepKind.FormatSheet: return settings is FormatSheetSettings;
                case PipelineStepKind.SaveText: return settings is SaveTextSettings;
                case PipelineStepKind.LoadPdf: return settings is LoadPdfSettings;
                case PipelineStepKind.ArrangePdf: return settings is ArrangePdfSettings;
                case PipelineStepKind.MergePdf: return settings is MergePdfSettings;
                case PipelineStepKind.SplitPdf: return settings is SplitPdfSettings;
                case PipelineStepKind.SavePdf: return settings is SavePdfSettings;
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
            var list = s as ValueListSettings;
            if (list != null)
                return list.Mode + " of " + (list.Source == ValueListSource.Stored
                    ? Count(list.Items == null ? 0 : list.Items.Count, "stored value")
                    : list.KeyFormula);
            var scan = s as FolderScanSettings;
            if (scan != null)
                return ScanSummary(scan);
            var find = s as FindFilesSettings;
            if (find != null)
                return MatchLabel(find.Match) + " - " + ScanSummary(find.Scan);
            var action = s as FileActionSettings;
            if (action != null)
            {
                if (action.Action == FileActionKind.Delete)
                    return action.Permanent ? "Delete permanently" : "Delete to Recycle Bin";
                return action.Action + " to " + action.DestFolder + " - " +
                    (action.Conflict == FileConflict.Skip ? "skip" : action.Conflict == FileConflict.Overwrite ? "overwrite" : "rename") + " if it exists";
            }
            var merge = s as MergeFoldersSettings;
            if (merge != null)
                return string.IsNullOrWhiteSpace(merge.Folder) ? "Choose a folder."
                    : "In " + merge.Folder + " when " + merge.Condition + ", keep " + MergeDuplicatesWork.KeepLabel(merge.Keep);
            var clean = s as CleanTextSettings;
            if (clean != null)
                return clean.Describe();
            var split = s as SplitColumnSettings;
            if (split != null)
                return "Split " + split.Column + (split.Mode == SplitMode.FixedWidths ? " by widths " + split.Widths
                    : split.Mode == SplitMode.Pattern ? " by pattern" : " on \"" + split.Delimiter + "\"");
            var combine = s as CombineColumnsSettings;
            if (combine != null)
                return string.Join(" + ", combine.Columns.ToArray()) + " -> " + combine.Header;
            var dedupe = s as RemoveDuplicatesSettings;
            if (dedupe != null)
                return RemoveDuplicatesWork.KeepLabel(dedupe.Keep) + (dedupe.UseFormula ? " by " + dedupe.KeyFormula
                    : dedupe.Keys.Count == 0 ? " (whole rows)" : " by " + string.Join(", ", dedupe.Keys.ToArray()));
            var headers = s as EditHeadersSettings;
            if (headers != null)
                return headers.Columns.Count + " header edits" + (string.IsNullOrWhiteSpace(headers.Formula) ? "" : ", formula for all headers");
            var arrange = s as ArrangeColumnsSettings;
            if (arrange != null)
                return ArrangeColumnsWork.Summary(arrange);
            var append = s as AppendSettings;
            if (append != null)
                return "Match by " + (append.Match == AppendMatch.ByHeader ? "header" : "position") +
                    (append.CommonOnly ? ", common columns only" : "") +
                    (string.IsNullOrWhiteSpace(append.SourceHeader) ? "" : ", source in " + append.SourceHeader);
            var highlight = s as HighlightSettings;
            if (highlight != null)
                return highlight.Rules.Count == 1 ? highlight.Rules[0].ToString() : Count(highlight.Rules.Count, "rule");
            var format = s as FormatSheetSettings;
            if (format != null)
                return FormatSheetWork.Summary(format);
            var pdfLoad = s as LoadPdfSettings;
            if (pdfLoad != null)
                return string.IsNullOrWhiteSpace(pdfLoad.PathColumn) ? Path.GetFileName(pdfLoad.Path) : pdfLoad.PathColumn;
            var arrangePdf = s as ArrangePdfSettings;
            if (arrangePdf != null)
                return arrangePdf.Pages;
            var mergePdf = s as MergePdfSettings;
            if (mergePdf != null)
                return mergePdf.Interleave ? "Interleave pages" : "In input order";
            var splitPdf = s as SplitPdfSettings;
            if (splitPdf != null)
                return string.IsNullOrWhiteSpace(splitPdf.Ranges) ? splitPdf.PagesPerFile + " pages per file" : splitPdf.Ranges;
            var pdfSave = s as SavePdfSettings;
            if (pdfSave != null)
                return Path.GetFileName(pdfSave.Path);
            var textSave = s as SaveTextSettings;
            if (textSave != null)
                return SaveTextWork.FormatLabel(textSave.Format) + " -> " +
                    (string.IsNullOrWhiteSpace(textSave.Path) ? "Choose a file." : Path.GetFileName(textSave.Path));
            return "";
        }

        static string ScanSummary(FolderScanSettings scan)
        {
            if (scan == null || string.IsNullOrWhiteSpace(scan.Folder))
                return "Choose a folder.";
            var what = scan.Include == FileInclude.Files ? "files" : scan.Include == FileInclude.Folders ? "folders" : "files and folders";
            return what + " in " + scan.Folder + (scan.Recursive ? " and below" : "");
        }

        static string MatchLabel(FileMatchMode mode)
        {
            switch (mode)
            {
                case FileMatchMode.NameEquals: return "Name = key";
                case FileMatchMode.StemEquals: return "Name (no ext) = key";
                case FileMatchMode.Contains: return "Name contains key";
                case FileMatchMode.StartsWith: return "Name starts with key";
                case FileMatchMode.EndsWith: return "Name ends with key";
                default: return "Formula match";
            }
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
            string targetId, bool writeFiles, ICollection<string> confirmedReplacements = null)
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
            var outputs = ResultLayout(def, order);
            var saves = new List<PipelineNode>();
            foreach (var node in order)
            {
                if (wanted != null && !wanted.Contains(node.Id))
                    continue;
                if (IsSaveKind(node.Kind))
                    saves.Add(node);
                if (cache.ContainsKey(node.Id))
                    continue;
                try
                {
                    cache[node.Id] = RunStep(def, node, cache, outputs);
                }
                catch (Exception ex)
                {
                    cache.Remove(node.Id);
                    result.FailedNodeId = node.Id;
                    result.Error = ex.Message;
                    return result;
                }
            }

            var sources = PdfSources(order, wanted, cache);
            if (!CheckSaveTargets(saves, cache, sources, result))
                return result;
            if (!writeFiles)
                return result;

            var replacements = confirmedReplacements == null ? Replacements(saves, cache) :
                new HashSet<string>(confirmedReplacements, StringComparer.OrdinalIgnoreCase);
            if (CheckReplacements(saves, cache, replacements, result) &&
                RunActions(def, order, wanted, cache, outputs, result) &&
                CheckSaveTargets(saves, cache, PdfSources(order, wanted, cache), result) &&
                CheckReplacements(saves, cache, replacements, result))
                WriteSaves(saves, cache, replacements, result);
            return result;
        }

        // Every step has worked as a dry run. Now do the File action steps in order, each planned
        // again against the disk as it is now, and run the steps after them again so a Save of
        // their result reports what really happened. False when a step failed.
        static bool RunActions(PipelineDefinition def, List<PipelineNode> order, HashSet<string> wanted,
            IDictionary<string, DataTable> cache, IDictionary<string, string[]> outputs, PipelineRunResult result)
        {
            var stale = new HashSet<string>();
            foreach (var node in order)
            {
                if (wanted != null && !wanted.Contains(node.Id))
                    continue;
                var action = IsActionKind(node.Kind);
                if (!action && !stale.Contains(node.Id))
                    continue;
                try
                {
                    if (action)
                    {
                        var done = RunStep(def, node, cache, outputs, true);
                        cache[node.Id] = done;
                        var items = ResultItems(done);
                        var outcomes = new List<string>();
                        foreach (var item in items)
                            outcomes.Add((items.Count > 1 ? item.Name + " (" : "") + FileOpsWork.Outcome(item.Table) +
                                (items.Count > 1 ? ")" : ""));
                        result.Actions.Add(node.Title + ": " + string.Join(", ", outcomes.ToArray()));
                        foreach (var d in Downstream(def, node.Id))
                        {
                            if (d != node.Id)
                                stale.Add(d);
                        }
                    }
                    else
                        cache[node.Id] = RunStep(def, node, cache, outputs);
                }
                catch (Exception ex)
                {
                    cache.Remove(node.Id);
                    result.FailedNodeId = node.Id;
                    result.Error = ex.Message + (result.Actions.Count > 0
                        ? "\n\nFile actions that already ran:\n" + string.Join("\n", result.Actions.ToArray())
                        : "");
                    return false;
                }
            }
            return true;
        }

        static HashSet<string> PdfSources(List<PipelineNode> order, HashSet<string> wanted, IDictionary<string, DataTable> cache)
        {
            var sources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var node in order)
            {
                if (node.Kind == PipelineStepKind.LoadPdf && (wanted == null || wanted.Contains(node.Id)))
                    foreach (var item in ResultItems(cache[node.Id]))
                        sources.UnionWith(PdfWork.Sources(item.Table));
            }
            return sources;
        }

        static HashSet<string> Replacements(List<PipelineNode> saves, IDictionary<string, DataTable> cache)
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var node in saves)
                paths.UnionWith(ReplacementPaths(node, cache));
            return paths;
        }

        static bool CheckReplacements(List<PipelineNode> saves, IDictionary<string, DataTable> cache,
            HashSet<string> before, PipelineRunResult result)
        {
            foreach (var node in saves)
                foreach (string path in ReplacementPaths(node, cache))
                {
                    if (before.Contains(path))
                        continue;
                    result.FailedNodeId = node.Id;
                    result.Error = "An output now needs a new replacement confirmation. Run again to review: " + path;
                    return false;
                }
            return true;
        }

        // Check every planned file, including the numbered outputs of Split PDF.
        static bool CheckSaveTargets(List<PipelineNode> saves, IDictionary<string, DataTable> cache,
            HashSet<string> pdfSources, PipelineRunResult result)
        {
            var targets = new Dictionary<string, List<PipelineNode>>(StringComparer.OrdinalIgnoreCase);
            foreach (var node in saves)
            {
                try
                {
                    var local = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string path in SaveTargetPaths(node, cache))
                    {
                        if (!local.Add(path))
                            throw new InvalidOperationException("More than one result writes " + path + ".");
                        if (Directory.Exists(path))
                            throw new InvalidOperationException("A folder occupies the output path: " + path);
                        if (pdfSources.Contains(path))
                            throw new InvalidOperationException("Choose a new output path to keep the source PDF: " + path);
                        List<PipelineNode> others;
                        if (!targets.TryGetValue(path, out others))
                        {
                            others = new List<PipelineNode>();
                            targets.Add(path, others);
                        }
                        var currentExcel = node.Settings as SaveStepSettings;
                        foreach (var other in others)
                        {
                            var otherExcel = other.Settings as SaveStepSettings;
                            if (currentExcel == null || otherExcel == null ||
                                (currentExcel.Mode == SaveStepMode.NewFile && otherExcel.Mode == SaveStepMode.NewFile))
                                throw new InvalidOperationException("This step and '" + other.Title + "' both write " + path +
                                    ". Choose different output files (Excel steps may instead add new sheets).");
                            if (string.Equals(currentExcel.SheetName.Trim(), otherExcel.SheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                                throw new InvalidOperationException("This step and '" + other.Title + "' both write sheet '" +
                                    currentExcel.SheetName.Trim() + "' in " + path + ".");
                        }
                        others.Add(node);
                    }
                }
                catch (Exception ex)
                {
                    result.FailedNodeId = node.Id;
                    result.Error = ex.Message;
                    return false;
                }
            }
            foreach (var target in targets)
            {
                var sheetSteps = target.Value.FindAll(node => IsNewSheet(node));
                if (sheetSteps.Count != target.Value.Count || !File.Exists(target.Key))
                    continue;
                try
                {
                    var existing = new HashSet<string>(ExcelFile.Load(target.Key, null, 1).Sheets, StringComparer.OrdinalIgnoreCase);
                    foreach (var node in sheetSteps)
                    {
                        var name = ((SaveStepSettings)node.Settings).SheetName.Trim();
                        if (!existing.Contains(name))
                            continue;
                        result.FailedNodeId = node.Id;
                        result.Error = "The workbook already has a sheet named '" + name + "': " + target.Key;
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    result.FailedNodeId = sheetSteps[0].Id;
                    result.Error = ex.Message;
                    return false;
                }
            }
            return true;
        }

        static bool IsNewSheet(PipelineNode node)
        {
            var excel = node.Settings as SaveStepSettings;
            return excel != null && excel.Mode == SaveStepMode.NewSheet;
        }

        public static IEnumerable<string> SaveTargetPaths(PipelineNode node, IDictionary<string, DataTable> cache)
        {
            DataTable table;
            if (cache == null || !cache.TryGetValue(node.Id, out table))
            {
                if (node.Kind != PipelineStepKind.SavePdf)
                    yield return SaveTargetPath(node);
                yield break;
            }
            foreach (var item in ResultItems(table))
            {
                if (node.Kind == PipelineStepKind.SavePdf)
                {
                    foreach (var path in PdfWork.Targets(item.Table, false))
                        yield return path;
                }
                else
                    yield return item.TargetPath ?? SaveTargetPath(node);
            }
        }

        public static IEnumerable<string> ReplacementPaths(PipelineNode node, IDictionary<string, DataTable> cache)
        {
            if (IsNewSheet(node))
                yield break;
            foreach (var path in SaveTargetPaths(node, cache))
                if (File.Exists(path))
                    yield return path;
        }

        public static string SaveTargetPath(PipelineNode node)
        {
            var text = node.Settings as SaveTextSettings;
            if (node.Kind == PipelineStepKind.SaveText && text != null)
                return SaveTextWork.OutputPath(text);
            var excel = node.Settings as SaveStepSettings;
            if (node.Kind != PipelineStepKind.Save || excel == null)
                throw new InvalidOperationException("Set up the save step first.");
            string error = CheckSave(excel);
            if (error != null)
                throw new InvalidOperationException(error);
            return Path.GetFullPath(excel.Path);
        }

        // New files first, then new sheets, so a sheet added to a workbook that another
        // step creates in the same run is not wiped out when that workbook is written.
        static void WriteSaves(List<PipelineNode> saves, IDictionary<string, DataTable> cache,
            HashSet<string> replacements, PipelineRunResult result)
        {
            var ordered = new List<PipelineNode>();
            foreach (var node in saves)
            {
                if (!IsNewSheet(node))
                    ordered.Add(node);
            }
            foreach (var node in saves)
            {
                if (IsNewSheet(node))
                    ordered.Add(node);
            }

            foreach (var node in ordered)
            {
                try
                {
                    var items = ResultItems(cache[node.Id]);
                    var written = new List<PipelineItem>();
                    if (node.Kind == PipelineStepKind.SavePdf)
                    {
                        var plans = new List<DataTable>();
                        foreach (var item in items)
                            plans.Add(item.Table);
                        var outputs = PdfWork.WritePlans(plans, result.Written, replacements);
                        for (int i = 0; i < items.Count; i++)
                            written.Add(new PipelineItem(outputs[i], items[i].Name, items[i].TargetPath));
                        cache[node.Id] = Pack(written);
                        continue;
                    }
                    foreach (var item in items)
                    {
                        DataTable table = item.Table;
                        if (node.Kind == PipelineStepKind.SaveText)
                        {
                            table = SaveTextWork.Write(table, replacements.Contains(item.TargetPath));
                            result.Written.Add(item.TargetPath);
                        }
                        else
                        {
                            var settings = (SaveStepSettings)node.Settings;
                            bool append = IsNewSheet(node) && File.Exists(item.TargetPath);
                            ExcelFile.SavePipelineOutput(table, item.TargetPath, settings.SheetName, append,
                                replacements.Contains(item.TargetPath));
                            result.Written.Add(item.TargetPath + " (sheet " + settings.SheetName.Trim() + ")");
                        }
                        written.Add(new PipelineItem(table, item.Name, item.TargetPath));
                    }
                    cache[node.Id] = Pack(written);
                }
                catch (Exception ex)
                {
                    result.FailedNodeId = node.Id;
                    result.Error = ex.Message;
                    return;
                }
            }
        }

        // Execute one item using the original feature logic.
        static DataTable RunSingle(PipelineNode node, DataTable first, DataTable second)
        {
            switch (node.Kind)
            {
                case PipelineStepKind.Load:
                {
                    var s = (LoadStepSettings)node.Settings;
                    if (string.IsNullOrWhiteSpace(s.Path))
                        throw new InvalidOperationException("Choose a file to load.");
                    return ExcelFile.Load(s.Path, string.IsNullOrEmpty(s.Sheet) ? null : s.Sheet, 0, !s.FirstRowIsData).Table;
                }
                case PipelineStepKind.LoadPdf:
                {
                    var input = first;
                    var settings = (LoadPdfSettings)node.Settings;
                    if (input != null)
                    {
                        settings = Clone(settings);
                        string header = string.IsNullOrEmpty(settings.PathColumn) ? "Path" : settings.PathColumn;
                        settings.PathColumn = LetterOf(input, header, true);
                    }
                    return PdfWork.Load(input, settings);
                }
                case PipelineStepKind.ArrangePdf:
                    return PdfWork.Arrange(first, second,
                        (ArrangePdfSettings)node.Settings);
                case PipelineStepKind.SplitPdf:
                    return PdfWork.Split(first, (SplitPdfSettings)node.Settings);
                case PipelineStepKind.FilterSort:
                {
                    var input = first;
                    return FilterSortWork.Run(input, FilterSortToLetters((FilterSortSettings)node.Settings, input, true));
                }
                case PipelineStepKind.Templates:
                {
                    var input = first;
                    return TemplatesWork.ApplyTemplate(input, TemplateToLetters((TemplateDefinition)node.Settings, input, true));
                }
                case PipelineStepKind.Compare:
                    return ExcelProcessingWork.Run(first, second,
                        (ExcelProcessingSettings)node.Settings);
                case PipelineStepKind.FillColumns:
                    return FillColumnsWork.Run(first, second,
                        (FillColumnsSettings)node.Settings);
                case PipelineStepKind.ValueList:
                {
                    var s = (ValueListSettings)node.Settings;
                    return FileOpsWork.BuildList(s.Source == ValueListSource.Stored ? null : first, s);
                }
                case PipelineStepKind.ListFolder:
                    return FileOpsWork.ListFolder((FolderScanSettings)node.Settings);
                case PipelineStepKind.FindFiles:
                    return FileOpsWork.FindFiles(first, (FindFilesSettings)node.Settings);
                case PipelineStepKind.MergeFolders:
                    return MergeDuplicatesWork.Plan((MergeFoldersSettings)node.Settings);
                case PipelineStepKind.CleanText:
                {
                    var input = first;
                    return CleanTextWork.Run(input, CleanTextToLetters((CleanTextSettings)node.Settings, input, true));
                }
                case PipelineStepKind.SplitColumn:
                {
                    var input = first;
                    return SplitCombineWork.Split(input, SplitToLetters((SplitColumnSettings)node.Settings, input, true));
                }
                case PipelineStepKind.CombineColumns:
                {
                    var input = first;
                    return SplitCombineWork.Combine(input, CombineToLetters((CombineColumnsSettings)node.Settings, input, true));
                }
                case PipelineStepKind.RemoveDuplicates:
                {
                    var input = first;
                    return RemoveDuplicatesWork.Run(input, RemoveDuplicatesToLetters((RemoveDuplicatesSettings)node.Settings, input, true));
                }
                case PipelineStepKind.EditHeaders:
                {
                    var input = first;
                    return EditHeadersWork.Run(input, HeadersToLetters((EditHeadersSettings)node.Settings, input, true));
                }
                case PipelineStepKind.ArrangeColumns:
                {
                    var input = first;
                    return ArrangeColumnsWork.Run(input, ArrangeToLetters((ArrangeColumnsSettings)node.Settings, input, true));
                }
                case PipelineStepKind.Highlight:
                {
                    var input = first;
                    return HighlightWork.Run(input, HighlightToLetters((HighlightSettings)node.Settings, input, true));
                }
                case PipelineStepKind.FormatSheet:
                {
                    var input = first;
                    return FormatSheetWork.Run(input, FormatSheetToLetters((FormatSheetSettings)node.Settings, input, true));
                }
                case PipelineStepKind.Save:
                    return first;
                default:
                    throw new InvalidOperationException("This step kind cannot process an individual input: " + node.Kind);
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

        public static LoadPdfSettings LoadPdfToNames(LoadPdfSettings settings, DataTable input)
        {
            var result = Clone(settings);
            result.PathColumn = NameOf(input, result.PathColumn);
            return result;
        }

        public static LoadPdfSettings LoadPdfToLetters(LoadPdfSettings settings, DataTable input)
        {
            var result = Clone(settings ?? new LoadPdfSettings());
            result.PathColumn = LetterOf(input, result.PathColumn, false) ?? result.PathColumn;
            return result;
        }

        public static SaveTextSettings SaveTextToNames(SaveTextSettings settings, DataTable input)
        {
            var s = Clone(settings);
            s.Column = NameOf(input, s.Column);
            s.ValueColumn = string.IsNullOrWhiteSpace(s.ValueColumn) ? "" : NameOf(input, s.ValueColumn);
            return s;
        }

        public static SaveTextSettings SaveTextToLetters(SaveTextSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Column = LetterOf(input, s.Column, strict) ?? "";
            s.ValueColumn = LetterOf(input, s.ValueColumn, strict && SaveTextWork.IsMap(s.Format)) ?? "";
            return s;
        }

        public static CleanTextSettings CleanTextToNames(CleanTextSettings settings, DataTable input)
        {
            var s = Clone(settings);
            Names(input, s.Columns);
            return s;
        }

        public static CleanTextSettings CleanTextToLetters(CleanTextSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            var had = s.Columns.Count;
            s.Columns = Letters(input, s.Columns, strict);
            // Every chosen column is gone: cleaning all columns instead would change cells nobody picked.
            if (strict && had > 0 && s.Columns.Count == 0)
                throw new InvalidOperationException("The columns this step cleans are not in its input any more. Set it up again.");
            return s;
        }

        public static SplitColumnSettings SplitToNames(SplitColumnSettings settings, DataTable input)
        {
            var s = Clone(settings);
            s.Column = NameOf(input, s.Column);
            return s;
        }

        public static SplitColumnSettings SplitToLetters(SplitColumnSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Column = LetterOf(input, s.Column, strict) ?? "";
            return s;
        }

        public static CombineColumnsSettings CombineToNames(CombineColumnsSettings settings, DataTable input)
        {
            var s = Clone(settings);
            Names(input, s.Columns);
            return s;
        }

        public static CombineColumnsSettings CombineToLetters(CombineColumnsSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Columns = Letters(input, s.Columns, strict);
            return s;
        }

        public static RemoveDuplicatesSettings RemoveDuplicatesToNames(RemoveDuplicatesSettings settings, DataTable input)
        {
            var s = Clone(settings);
            Names(input, s.Keys);
            return s;
        }

        public static RemoveDuplicatesSettings RemoveDuplicatesToLetters(RemoveDuplicatesSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Keys = Letters(input, s.Keys, strict);
            return s;
        }

        public static ArrangeColumnsSettings ArrangeToNames(ArrangeColumnsSettings settings, DataTable input)
        {
            var s = Clone(settings);
            foreach (var c in s.Columns)
                c.Source = NameOf(input, c.Source);
            return s;
        }

        public static EditHeadersSettings HeadersToNames(EditHeadersSettings settings, DataTable input)
        {
            var s = Clone(settings);
            foreach (var edit in s.Columns)
            {
                if (string.IsNullOrWhiteSpace(edit.Source)) continue;
                bool isId;
                edit.Source = UnambiguousNameOf(input, edit.Source, out isId);
                edit.SourceIsId = isId;
            }
            return s;
        }

        public static EditHeadersSettings HeadersToLetters(EditHeadersSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Columns.RemoveAll(edit =>
            {
                if (string.IsNullOrWhiteSpace(edit.Source)) return false;
                string letter;
                if (edit.SourceIsId)
                {
                    var column = input.Columns[edit.Source];
                    if (column == null && strict) throw new InvalidOperationException("Header column is missing: " + edit.Source);
                    letter = column == null ? null : ExcelColumns.Letter(column.Ordinal);
                }
                else letter = LetterOf(input, edit.Source, strict);
                if (letter == null) return true;
                edit.Source = letter;
                edit.SourceIsId = false;
                return false;
            });
            return s;
        }

        // A dropped column that is gone upstream is no loss, even when running.
        public static ArrangeColumnsSettings ArrangeToLetters(ArrangeColumnsSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Columns.RemoveAll(c =>
            {
                var letter = LetterOf(input, c.Source, strict && c.Keep);
                if (letter == null)
                    return true;
                c.Source = letter;
                return false;
            });
            return s;
        }

        public static HighlightSettings HighlightToNames(HighlightSettings settings, DataTable input)
        {
            var s = Clone(settings);
            foreach (var r in s.Rules)
                Names(input, r.Columns);
            return s;
        }

        public static HighlightSettings HighlightToLetters(HighlightSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            foreach (var r in s.Rules)
                r.Columns = Letters(input, r.Columns, strict);
            return s;
        }

        public static FormatSheetSettings FormatSheetToNames(FormatSheetSettings settings, DataTable input)
        {
            var s = Clone(settings);
            foreach (var c in s.Columns)
            {
                bool isId;
                c.Column = UnambiguousNameOf(input, c.Column, out isId);
                c.ColumnIsId = isId;
            }
            return s;
        }

        public static FormatSheetSettings FormatSheetToLetters(FormatSheetSettings settings, DataTable input, bool strict)
        {
            var s = Clone(settings);
            s.Columns.RemoveAll(c =>
            {
                string letter;
                if (c.ColumnIsId)
                {
                    var column = input.Columns[c.Column];
                    if (column == null && strict)
                        throw new InvalidOperationException("Column '" + c.Column + "' is not in this step's input.");
                    letter = column == null ? null : ExcelColumns.Letter(column.Ordinal);
                }
                else
                    letter = LetterOf(input, c.Column, strict);
                if (letter == null)
                    return true;
                c.Column = letter;
                c.ColumnIsId = false;
                return false;
            });
            return s;
        }

        static void Names(DataTable input, List<string> letters)
        {
            for (var i = 0; i < letters.Count; i++)
                letters[i] = NameOf(input, letters[i]);
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
                return string.IsNullOrWhiteSpace(ExcelFile.Header(input.Columns[index]))
                    ? input.Columns[index].ColumnName : ExcelFile.Header(input.Columns[index]);
            return letter;
        }

        static string UnambiguousNameOf(DataTable input, string letter, out bool isId)
        {
            var column = input.Columns[ExcelColumns.Index(input, letter)];
            string header = ExcelFile.Header(column);
            int matches = 0;
            foreach (DataColumn candidate in input.Columns)
                if (string.Equals(ExcelFile.Header(candidate), header, StringComparison.OrdinalIgnoreCase)) matches++;
            isId = string.IsNullOrWhiteSpace(header) || matches > 1;
            return isId ? column.ColumnName : header;
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
