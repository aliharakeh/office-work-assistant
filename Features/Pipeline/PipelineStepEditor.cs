using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.Append;
using OfficeWorkAssistant.Features.ArrangeColumns;
using OfficeWorkAssistant.Features.EditHeaders;
using OfficeWorkAssistant.Features.CleanText;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.FormatSheet;
using OfficeWorkAssistant.Features.Highlight;
using OfficeWorkAssistant.Features.MergeDuplicates;
using OfficeWorkAssistant.Features.SaveText;
using OfficeWorkAssistant.Features.Pdf;
using OfficeWorkAssistant.Features.RemoveDuplicates;
using OfficeWorkAssistant.Features.SplitCombine;
using OfficeWorkAssistant.Features.Templates;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // Builder and Run use the same editors and column-reference conversions.
    public static class PipelineStepEditor
    {
        public static FrameworkElement Prepare(PipelineDefinition def, PipelineNode node, Action<object> use)
        {
            var ports = PipelineWork.InputPorts(def, node);
            var inputs = new DataTable[ports.Length];
            var labels = new string[ports.Length];
            var cache = new Dictionary<string, DataTable>();
            var layout = PipelineWork.ResultLayout(def);
            for (var i = 0; i < ports.Length; i++)
            {
                var link = PipelineWork.InputLink(def, node.Id, ports[i]);
                if (link == null && PipelineWork.IsOptionalInput(node.Kind, ports[i])) continue;
                if (link == null) throw new InvalidOperationException("Connect this step's input in Builder mode first.");
                var result = PipelineWork.Run(def, cache, link.From, false);
                if (result.Error != null) throw new InvalidOperationException(result.Error);
                inputs[i] = PipelineWork.LinkedResult(def, link, cache[link.From], layout);
                labels[i] = "From step: " + PipelineWork.Find(def, link.From).Title;
            }
            return Create(def, node, inputs, labels, use);
        }

        public static FrameworkElement Create(PipelineDefinition def, PipelineNode node,
            DataTable[] inputs, string[] labels, Action<object> use)
        {
            if (node.Kind == PipelineStepKind.Load || node.Kind == PipelineStepKind.Save)
                return FileSettings(node, use);
            if (node.Kind == PipelineStepKind.Append || node.Kind == PipelineStepKind.MergePdf)
                return CombineEditor(def, node, inputs, labels, use);

            var sample = (DataTable[])inputs.Clone();
            var sampleLabels = (string[])labels.Clone();
            var ports = PipelineWork.InputPorts(def, node);
            int firstCount = 0;
            int secondCount = 0;
            var fillTargets = new List<PipelineItem>();
            var fillLookups = new List<PipelineItem>();
            var fillTargetLabels = new List<string>();
            var fillLookupLabels = new List<string>();
            for (int i = 0; i < inputs.Length; i++)
            {
                if (inputs[i] == null)
                    continue;
                var items = PipelineWork.ResultItems(inputs[i]);
                bool isFill = node.Kind == PipelineStepKind.FillColumns;
                var targetRole = isFill ? PipelineWork.PortIn : PipelineWork.PortB;
                bool target = ports[i].StartsWith(targetRole, StringComparison.Ordinal);
                if (target) secondCount += items.Count;
                else firstCount += items.Count;
                if (isFill)
                {
                    var sheets = target ? fillTargets : fillLookups;
                    var sheetLabels = target ? fillTargetLabels : fillLookupLabels;
                    sheets.AddRange(items);
                    foreach (var item in items) sheetLabels.Add(items.Count == 1 ? labels[i] : item.Name);
                }
                sample[i] = items[0].Table;
                if (items.Count > 1)
                    sampleLabels[i] = items[0].Name;
            }
            if (node.Kind == PipelineStepKind.FillColumns)
            {
                if (fillTargets.Count == 0) return Notice("Connect a sheet to fill.");
                fillTargets.AddRange(fillLookups);
                fillTargetLabels.AddRange(fillLookupLabels);
                return new FillColumnsPage(fillTargets.ConvertAll(item => item.Table), fillTargetLabels,
                    PipelineWork.Clone(node.Settings as FillColumnsSettings), r => use(r));
            }
            var editor = CreateEditor(node, sample, sampleLabels, use);
            int count = Math.Max(firstCount, secondCount);
            if (count <= 1)
                return editor;
            var panel = new DockPanel();
            var note = Notice("These settings run on " + count.ToString(CultureInfo.InvariantCulture) +
                " inputs. The editor shows the first input.");
            note.Margin = new Thickness(0, 0, 0, 8);
            DockPanel.SetDock(note, Dock.Top);
            panel.Children.Add(note);
            panel.Children.Add(editor);
            return panel;
        }

        static FrameworkElement CombineEditor(PipelineDefinition def, PipelineNode node,
            DataTable[] inputs, string[] labels, Action<object> use)
        {
            var ports = PipelineWork.InputPorts(def, node);
            var tables = new List<DataTable>();
            var shown = new List<string>();
            var names = new List<string>();
            for (int i = 0; i < inputs.Length; i++)
            {
                if (inputs[i] == null)
                    continue;
                var items = PipelineWork.ResultItems(inputs[i]);
                foreach (var item in items)
                {
                    tables.Add(item.Table);
                    shown.Add(items.Count == 1 ? labels[i] : item.Name);
                    var link = PipelineWork.InputLink(def, node.Id, ports[i]);
                    names.Add(items.Count == 1 && link != null &&
                        (string.IsNullOrEmpty(link.FromPort) || link.FromPort == PipelineWork.PortOut)
                        ? PipelineWork.Find(def, link.From).Title : item.Name);
                }
            }
            if (node.Kind == PipelineStepKind.MergePdf)
                return new MergePdfPage(tables, shown, node.Settings as MergePdfSettings, r => use(r));
            if (tables.Count == 0)
                return Notice("Link the steps whose rows you want to stack: drag from the right dot of each one to this step's left dot.");
            return new AppendPage(tables, shown, names, PipelineWork.Clone(node.Settings as AppendSettings), r => use(r));
        }

        static FrameworkElement CreateEditor(PipelineNode node, DataTable[] inputs, string[] labels, Action<object> use)
        {
            switch (node.Kind)
            {
                case PipelineStepKind.LoadPdf:
                {
                    var input = inputs[0];
                    var settings = node.Settings as LoadPdfSettings;
                    return new LoadPdfPage(input, labels[0], input == null ? settings : PipelineWork.LoadPdfToLetters(settings, input),
                        r => use(input == null ? r : PipelineWork.LoadPdfToNames(r, input)));
                }
                case PipelineStepKind.ArrangePdf:
                    return new ArrangePdfPage(inputs[0], labels[0], inputs[1], labels[1], node.Settings as ArrangePdfSettings, r => use(r));
                case PipelineStepKind.SplitPdf:
                    return new SplitPdfPage(inputs[0], labels[0], node.Settings as SplitPdfSettings, r => use(r));
                case PipelineStepKind.SavePdf:
                    return new SavePdfPage(inputs[0], labels[0], node.Settings as SavePdfSettings, r => use(r));
                case PipelineStepKind.SaveText:
                {
                    var input = inputs[0];
                    var s = node.Settings as SaveTextSettings;
                    return new SaveTextPage(input, labels[0], s == null ? null : PipelineWork.SaveTextToLetters(s, input, false),
                        r => use(PipelineWork.SaveTextToNames(r, input)));
                }
                case PipelineStepKind.FilterSort:
                {
                    var input = inputs[0];
                    var s = node.Settings as FilterSortSettings;
                    return new FilterSortPage(input, labels[0], s == null ? null : PipelineWork.FilterSortToLetters(s, input, false),
                        r => use(PipelineWork.FilterSortToNames(r, input)));
                }
                case PipelineStepKind.Templates:
                {
                    var input = inputs[0];
                    var t = node.Settings as TemplateDefinition;
                    return new TemplatesPage(input, labels[0], t == null ? null : PipelineWork.TemplateToLetters(t, input, false),
                        r => use(PipelineWork.TemplateToNames(r, input)));
                }
                case PipelineStepKind.Compare:
                    return new ExcelProcessingPage(inputs[0], labels[0], inputs[1], labels[1],
                        PipelineWork.Clone(node.Settings as ExcelProcessingSettings), r => use(r));
                case PipelineStepKind.ValueList:
                    return new ValueListPage(inputs[0], labels[0], PipelineWork.Clone(node.Settings as ValueListSettings), r => use(r));
                case PipelineStepKind.ListFolder:
                    return new FindFilesPage(PipelineWork.Clone(node.Settings as FolderScanSettings), r => use(r));
                case PipelineStepKind.FindFiles:
                    return new FindFilesPage(inputs[0], labels[0], PipelineWork.Clone(node.Settings as FindFilesSettings), r => use(r));
                case PipelineStepKind.FileAction:
                    return new FileActionPage(inputs[0], labels[0], PipelineWork.Clone(node.Settings as FileActionSettings), r => use(r));
                case PipelineStepKind.MergeFolders:
                    return new MergeDuplicatesPage(PipelineWork.Clone(node.Settings as MergeFoldersSettings), r => use(r));
                case PipelineStepKind.CleanText:
                {
                    var input = inputs[0];
                    var s = node.Settings as CleanTextSettings;
                    return new CleanTextPage(input, labels[0], s == null ? null : PipelineWork.CleanTextToLetters(s, input, false),
                        r => use(PipelineWork.CleanTextToNames(r, input)));
                }
                case PipelineStepKind.SplitColumn:
                {
                    var input = inputs[0];
                    var s = node.Settings as SplitColumnSettings;
                    return new SplitColumnPage(input, labels[0], s == null ? null : PipelineWork.SplitToLetters(s, input, false),
                        r => use(PipelineWork.SplitToNames(r, input)));
                }
                case PipelineStepKind.CombineColumns:
                {
                    var input = inputs[0];
                    var s = node.Settings as CombineColumnsSettings;
                    return new CombineColumnsPage(input, labels[0], s == null ? null : PipelineWork.CombineToLetters(s, input, false),
                        r => use(PipelineWork.CombineToNames(r, input)));
                }
                case PipelineStepKind.RemoveDuplicates:
                {
                    var input = inputs[0];
                    var s = node.Settings as RemoveDuplicatesSettings;
                    return new RemoveDuplicatesPage(input, labels[0], s == null ? null : PipelineWork.RemoveDuplicatesToLetters(s, input, false),
                        r => use(PipelineWork.RemoveDuplicatesToNames(r, input)));
                }
                case PipelineStepKind.ArrangeColumns:
                {
                    var input = inputs[0];
                    var s = node.Settings as ArrangeColumnsSettings;
                    return new ArrangeColumnsPage(input, labels[0], s == null ? null : PipelineWork.ArrangeToLetters(s, input, false),
                        r => use(PipelineWork.ArrangeToNames(r, input)));
                }
                case PipelineStepKind.EditHeaders:
                {
                    var input = inputs[0];
                    var s = node.Settings as EditHeadersSettings;
                    return new EditHeadersPage(input, labels[0], s == null ? null : PipelineWork.HeadersToLetters(s, input, false),
                        r => use(PipelineWork.HeadersToNames(r, input)));
                }
                case PipelineStepKind.Highlight:
                {
                    var input = inputs[0];
                    var s = node.Settings as HighlightSettings;
                    return new HighlightPage(input, labels[0], s == null ? null : PipelineWork.HighlightToLetters(s, input, false),
                        r => use(PipelineWork.HighlightToNames(r, input)));
                }
                case PipelineStepKind.FormatSheet:
                {
                    var input = inputs[0];
                    var s = node.Settings as FormatSheetSettings;
                    return new FormatSheetPage(input, labels[0], s == null ? null : PipelineWork.FormatSheetToLetters(s, input, false),
                        r => use(PipelineWork.FormatSheetToNames(r, input)));
                }
                default:
                    throw new InvalidOperationException("No editor for " + node.Kind + ".");
            }
        }

        static FrameworkElement FileSettings(PipelineNode node, Action<object> use)
        {
            var panel = new StackPanel();
            panel.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/OfficeWorkAssistant;component/Excel/EditorStyles.xaml", UriKind.Relative) });
            var load = node.Settings as LoadStepSettings;
            var save = node.Settings as SaveStepSettings;
            var path = Field(panel, "File path", load != null ? load.Path : save == null ? "" : save.Path);
            var sheet = Field(panel, load != null ? "Sheet (blank = first sheet)" : "Output sheet name",
                load != null ? load.Sheet : save == null ? "Result" : save.SheetName);
            var option = new CheckBox
            {
                Content = load != null ? "No header row (first row is data)" : "Add a new sheet to an existing workbook",
                IsChecked = load != null ? load.FirstRowIsData : save != null && save.Mode == SaveStepMode.NewSheet,
                Margin = new Thickness(0, 6, 0, 8)
            };
            panel.Children.Add(option);
            if (node.Kind == PipelineStepKind.Save)
                panel.Children.Add(Notice("With several results, the source name is added to each output filename."));
            var apply = new Button { Content = "Apply to step", HorizontalAlignment = HorizontalAlignment.Left, Padding = new Thickness(8, 4, 8, 4) };
            apply.Click += (sender, e) =>
            {
                object settings;
                string error = null;
                if (node.Kind == PipelineStepKind.Load)
                {
                    var s = new LoadStepSettings { Path = path.Text.Trim(), Sheet = sheet.Text.Trim(), FirstRowIsData = option.IsChecked == true };
                    try { ExcelFile.Load(s.Path, s.Sheet, 1, !s.FirstRowIsData); }
                    catch (Exception ex) { error = ex.Message; }
                    settings = s;
                }
                else
                {
                    var s = new SaveStepSettings { Path = path.Text.Trim(), SheetName = sheet.Text.Trim(),
                        Mode = option.IsChecked == true ? SaveStepMode.NewSheet : SaveStepMode.NewFile };
                    error = PipelineWork.CheckSave(s);
                    settings = s;
                }
                if (error != null)
                    Notices.Show(panel, error, node.Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                else use(settings);
            };
            panel.Children.Add(apply);
            return panel;
        }

        static TextBox Field(StackPanel panel, string label, string value)
        {
            panel.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 6, 0, 2) });
            var box = new TextBox { Text = value ?? "" };
            panel.Children.Add(box);
            return box;
        }

        static FrameworkElement Notice(string text)
        {
            return new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap };
        }
    }
}
