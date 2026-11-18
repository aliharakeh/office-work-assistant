using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.Append;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.MergeDuplicates;
using OfficeWorkAssistant.Features.Pdf;
using OfficeWorkAssistant.Features.SaveText;

namespace OfficeWorkAssistant.Features.Pipeline
{
    public sealed class PipelineItem
    {
        public DataTable Table { get; private set; }
        public string Name { get; private set; }
        public string TargetPath { get; private set; }

        internal PipelineItem(DataTable table, string name, string targetPath = null)
        {
            Table = table;
            Name = name;
            TargetPath = targetPath;
        }
    }

    public static partial class PipelineWork
    {
        public const string PortOut = "Out";

        static bool MapsOutputs(PipelineStepKind kind)
        {
            return Ports(kind).Length != 0 && !CombinesInputs(kind) && kind != PipelineStepKind.Compare;
        }

        public static string[] OutputPorts(PipelineNode node, IDictionary<string, string[]> layout)
        {
            return MapsOutputs(node.Kind) ? layout[node.Id] : new[] { PortOut };
        }

        // Each key follows the incoming link that produced the result. Old links may carry
        // several results, so their keys also include the upstream result's identity.
        public static Dictionary<string, string[]> ResultLayout(PipelineDefinition def)
        {
            return ResultLayout(def, Order(def));
        }

        static Dictionary<string, string[]> ResultLayout(PipelineDefinition def, List<PipelineNode> order,
            bool removeDisconnected = false)
        {
            var layout = new Dictionary<string, string[]>();
            foreach (var node in order)
            {
                if (removeDisconnected)
                {
                    var links = def.Links.FindAll(link => link.To == node.Id);
                    foreach (var link in links)
                    {
                        if (string.IsNullOrEmpty(link.FromPort))
                            continue;
                        var source = Find(def, link.From);
                        var keys = OutputPorts(source, layout);
                        if (link.FromPort == PortOut && MapsOutputs(source.Kind))
                            link.FromPort = keys[0];
                        else if (Array.IndexOf(keys, link.FromPort) < 0)
                            def.Links.Remove(link);
                    }
                    CompactPorts(def, node.Id);
                }

                var ports = Ports(node.Kind);
                if (ports.Length == 0 || CombinesInputs(node.Kind) || UsesStoredValues(node))
                {
                    layout.Add(node.Id, new[] { PortOut });
                    continue;
                }

                var first = InputKeys(def, node, ports[0], layout);
                var second = ports.Length == 2 ? InputKeys(def, node, ports[1], layout) : new List<string>();
                var fill = node.Settings as FillColumnsSettings;
                if (fill != null && fill.OutputInput > 0)
                {
                    layout.Add(node.Id, new[] { second.Count == 0 ? PortOut : second[0] });
                    continue;
                }
                var keysForOutputs = JointFill(node) || PreferSecond(node.Kind, first.Count, second.Count) ? second : first;
                layout.Add(node.Id, keysForOutputs.Count == 0 ? new[] { PortOut } : keysForOutputs.ToArray());
            }
            return layout;
        }

        public static void RemoveDisconnectedOutputs(PipelineDefinition def)
        {
            ResultLayout(def, Order(def), true);
        }

        static List<string> InputKeys(PipelineDefinition def, PipelineNode node, string role,
            IDictionary<string, string[]> layout)
        {
            var keys = new List<string>();
            foreach (var link in BatchLinks(def, node, role))
            {
                var source = Find(def, link.From);
                var upstream = layout[link.From];
                int index = OutputIndex(source, link, upstream);
                if (index >= 0)
                    keys.Add(link.Id);
                else if (index == -1)
                    foreach (var key in upstream)
                        keys.Add(key == PortOut ? link.Id : link.Id + "/" + key);
            }
            return keys;
        }

        // -1 carries the entire result, and -2 refers to a missing socket.
        static int OutputIndex(PipelineNode source, PipelineLink link, string[] keys)
        {
            if (string.IsNullOrEmpty(link.FromPort) || (link.FromPort == PortOut && !MapsOutputs(source.Kind)))
                return -1;
            if (!MapsOutputs(source.Kind))
                return -2;
            if (link.FromPort == PortOut)
                return 0;
            int index = Array.IndexOf(keys, link.FromPort);
            return index < 0 ? -2 : index;
        }

        public static DataTable LinkedResult(PipelineDefinition def, PipelineLink link, DataTable table,
            IDictionary<string, string[]> layout)
        {
            var selected = LinkedItems(def, link, table, layout);
            return selected.Count == ResultItems(table).Count ? table : Pack(new List<PipelineItem>(selected));
        }

        static IList<PipelineItem> LinkedItems(PipelineDefinition def, PipelineLink link, DataTable table,
            IDictionary<string, string[]> layout)
        {
            var source = Find(def, link.From);
            int index = OutputIndex(source, link, layout[link.From]);
            var items = ResultItems(table);
            if (index < -1 || index >= items.Count)
                throw new InvalidOperationException("Reconnect the output from '" + source.Title + "'. It no longer has that result.");
            return index == -1 || items.Count == 1 ? items : new[] { items[index] };
        }

        static bool UsesStoredValues(PipelineNode node)
        {
            var list = node.Settings as ValueListSettings;
            return list != null && list.Source == ValueListSource.Stored;
        }

        static bool PreferSecond(PipelineStepKind kind, int firstCount, int secondCount)
        {
            return secondCount > 0 && (firstCount < secondCount ||
                (firstCount == secondCount && kind == PipelineStepKind.FillColumns));
        }

        internal static bool JointFill(PipelineNode node)
        {
            var settings = node.Settings as FillColumnsSettings;
            return settings != null && settings.JointLookups;
        }

        // Keep the existing table cache API. A batch gets a dedicated representative table;
        // the real members keep their own data, styles and PDF document metadata.
        static readonly ConditionalWeakTable<DataTable, IList<PipelineItem>> Items =
            new ConditionalWeakTable<DataTable, IList<PipelineItem>>();

        public static IList<PipelineItem> ResultItems(DataTable table)
        {
            IList<PipelineItem> items;
            return Items.TryGetValue(table, out items) ? items :
                Array.AsReadOnly(new[] { new PipelineItem(table, table.TableName) });
        }

        static DataTable Pack(List<PipelineItem> items)
        {
            var table = items[0].Table;
            IList<PipelineItem> existing;
            if (items.Count > 1 || Items.TryGetValue(table, out existing))
                table = table.Copy();
            Items.Add(table, items.AsReadOnly());
            return table;
        }

        internal static string VariablePort(PipelineStepKind kind, string role, int index)
        {
            if (CombinesInputs(kind))
                return AppendPort(index);
            return index == 1 ? role : role + index.ToString(CultureInfo.InvariantCulture);
        }

        static int InputNumber(PipelineStepKind kind, string port, string role)
        {
            if (CombinesInputs(kind))
                return AppendPortNumber(port);
            if (port == role)
                return 1;
            int number = NumberedPort(port, role);
            return number > 1 ? number : 0;
        }

        static List<PipelineLink> BatchLinks(PipelineDefinition def, PipelineNode node, string role)
        {
            var links = def.Links.FindAll(link => link.To == node.Id && InputNumber(node.Kind, link.Port, role) > 0);
            links.Sort((a, b) => InputNumber(node.Kind, a.Port, role).CompareTo(InputNumber(node.Kind, b.Port, role)));
            return links;
        }

        static void AddInputs(PipelineDefinition def, PipelineLink link, IDictionary<string, DataTable> cache,
            IDictionary<string, string[]> layout, List<PipelineItem> inputs, bool combine)
        {
            DataTable table;
            if (!cache.TryGetValue(link.From, out table))
                throw new InvalidOperationException("The step before this one has no result yet.");
            var items = LinkedItems(def, link, table, layout);
            if (combine && items.Count == 1 && (string.IsNullOrEmpty(link.FromPort) || link.FromPort == PortOut))
                inputs.Add(new PipelineItem(items[0].Table, Find(def, link.From).Title));
            else
                inputs.AddRange(items);
        }

        static List<PipelineItem> ReadPort(PipelineDefinition def, PipelineNode node, string port,
            IDictionary<string, DataTable> cache, IDictionary<string, string[]> layout)
        {
            var items = new List<PipelineItem>();
            foreach (var link in BatchLinks(def, node, port))
                AddInputs(def, link, cache, layout, items, CombinesInputs(node.Kind));
            if (items.Count == 0 && !IsOptionalInput(node.Kind, port))
                throw new InvalidOperationException("Connect " + PortLabel(node.Kind, port) + " first.");
            return items;
        }

        static DataTable RunStep(PipelineDefinition def, PipelineNode node, IDictionary<string, DataTable> cache,
            IDictionary<string, string[]> layout, bool applyActions = false)
        {
            if (node.Settings == null)
                throw new InvalidOperationException("Set up this step first: double-click it.");

            var first = new List<PipelineItem>();
            var second = new List<PipelineItem>();
            var ports = Ports(node.Kind);
            var list = node.Settings as ValueListSettings;
            if (ports.Length == 1 && !UsesStoredValues(node))
            {
                first = ReadPort(def, node, ports[0], cache, layout);
                if (list != null && first.Count == 0)
                    throw new InvalidOperationException("Connect a table to build the list from.");
            }
            else if (ports.Length == 2)
            {
                first = ReadPort(def, node, ports[0], cache, layout);
                second = ReadPort(def, node, ports[1], cache, layout);
                var selectedFill = node.Settings as FillColumnsSettings;
                if (selectedFill != null && selectedFill.OutputInput > 0)
                {
                    var inputs = new List<PipelineItem>(second);
                    inputs.AddRange(first);
                    int selected = selectedFill.OutputInput - 1;
                    if (selected >= inputs.Count)
                        throw new InvalidOperationException("Select a connected sheet to fill.");
                    var target = inputs[selected];
                    inputs.RemoveAt(selected);
                    var result = FillColumnsWork.Fill(inputs.ConvertAll(item => item.Table), target.Table, null,
                        selectedFill.Rules, 0, true, selectedFill.OutputInput, selectedFill.MatchFormula);
                    return Pack(new List<PipelineItem> { OutputItem(node, result.Table, target.Name) });
                }
                if (JointFill(node))
                {
                    var settings = (FillColumnsSettings)node.Settings;
                    var sources = first.ConvertAll(item => item.Table);
                    var filled = new List<PipelineItem>();
                    foreach (var target in second)
                    {
                        var result = FillColumnsWork.Fill(sources, target.Table, settings.Keys, settings.Rules, 0, settings.MatchAll);
                        filled.Add(OutputItem(node, result.Table, target.Name));
                    }
                    return Pack(filled);
                }
                if (first.Count > 1 && second.Count > 1 && first.Count != second.Count)
                    throw new InvalidOperationException("Connect batches of the same size to A and B, or use a single result on one side.");
            }

            if (CombinesInputs(node.Kind))
            {
                var tables = first.ConvertAll(item => item.Table);
                if (node.Kind == PipelineStepKind.Append && tables.Count < 2)
                    throw new InvalidOperationException("Connect at least two tables to append. One connection can carry several tables.");
                var table = node.Kind == PipelineStepKind.Append
                    ? AppendWork.Run(tables, first.ConvertAll(item => item.Name), (AppendSettings)node.Settings)
                    : PdfWork.Merge(tables, (MergePdfSettings)node.Settings);
                return Pack(new List<PipelineItem> { OutputItem(node, table, node.Title) });
            }

            int count = Math.Max(1, Math.Max(first.Count, second.Count));
            var results = new List<PipelineItem>();
            var savedPaths = IsSaveKind(node.Kind) ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;
            var pdfPaths = node.Kind == PipelineStepKind.SavePdf ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;
            var action = node.Settings as FileActionSettings;
            var actionSources = action == null ? null : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var actionTargets = action == null ? null : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string basePath = node.Kind == PipelineStepKind.SavePdf
                ? PdfWork.OutputPath((SavePdfSettings)node.Settings)
                : savedPaths == null ? null : SaveTargetPath(node);

            for (int i = 0; i < count; i++)
            {
                var a = ItemAt(first, i);
                var b = ItemAt(second, i);
                var source = PreferSecond(node.Kind, first.Count, second.Count)
                    ? b : a;
                string name = source == null ? SourceName(node) : source.Name;
                try
                {
                    string path = count > 1 && basePath != null ? BatchPath(basePath, name, savedPaths) : basePath;
                    DataTable table;
                    if (action != null)
                        table = FileOpsWork.PlanActions(a.Table, action, actionSources, actionTargets);
                    else if (node.Kind == PipelineStepKind.MergeFolders && applyActions)
                        table = MergeDuplicatesWork.Apply((MergeFoldersSettings)node.Settings);
                    else if (node.Kind == PipelineStepKind.SaveText)
                    {
                        var settings = SaveTextToLetters((SaveTextSettings)node.Settings, a.Table, true);
                        settings.Path = path;
                        table = SaveTextWork.Run(a.Table, settings);
                    }
                    else if (node.Kind == PipelineStepKind.SavePdf)
                    {
                        while (true)
                        {
                            table = PdfWork.PlanSave(a.Table, new SavePdfSettings { Path = path });
                            var targets = new List<string>(PdfWork.Targets(table, false));
                            if (!pdfPaths.Overlaps(targets))
                            {
                                pdfPaths.UnionWith(targets);
                                break;
                            }
                            path = BatchPath(basePath, name, savedPaths);
                        }
                    }
                    else
                        table = RunSingle(node, a == null ? null : a.Table, b == null ? null : b.Table);
                    results.Add(OutputItem(node, table, name, path));
                }
                catch (Exception ex)
                {
                    if (count == 1)
                        throw;
                    throw new InvalidOperationException("Input " + (i + 1).ToString(CultureInfo.InvariantCulture) +
                        " (" + name + "): " + ex.Message, ex);
                }
            }
            if (action != null && applyActions)
                foreach (var item in results)
                    FileOpsWork.ApplyPlan(item.Table, action);
            return Pack(results);
        }

        static PipelineItem ItemAt(List<PipelineItem> items, int index)
        {
            return items.Count == 0 ? null : items[items.Count == 1 ? 0 : index];
        }

        static string SourceName(PipelineNode node)
        {
            var excel = node.Settings as LoadStepSettings;
            var pdf = node.Settings as LoadPdfSettings;
            string path = excel != null ? excel.Path : pdf == null ? null : pdf.Path;
            return string.IsNullOrWhiteSpace(path) ? node.Title : Path.GetFileNameWithoutExtension(path);
        }

        static string BatchPath(string path, string name, HashSet<string> used)
        {
            name = name ?? "";
            foreach (char invalid in Path.GetInvalidFileNameChars())
                name = name.Replace(invalid, '_');
            name = name.Trim().TrimEnd('.');
            if (name.Length == 0)
                name = "input";
            string stem = Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path) + "-" + name);
            string extension = Path.GetExtension(path);
            string candidate = stem + extension;
            for (int i = 2; !used.Add(candidate); i++)
                candidate = stem + "-" + i.ToString(CultureInfo.InvariantCulture) + extension;
            return candidate;
        }

        static PipelineItem OutputItem(PipelineNode node, DataTable table, string name, string path = null)
        {
            // Save file passes the original table through so it retains its sheet style.
            if (!IsStyleKind(node.Kind) && node.Kind != PipelineStepKind.Save)
                ExcelFile.ClearStyle(table);
            if (node.Kind != PipelineStepKind.Save && node.Kind != PipelineStepKind.LoadPdf &&
                node.Kind != PipelineStepKind.ArrangePdf && node.Kind != PipelineStepKind.MergePdf && node.Kind != PipelineStepKind.SplitPdf)
                PdfWork.ClearDocuments(table);
            return new PipelineItem(table, name, path);
        }
    }
}
