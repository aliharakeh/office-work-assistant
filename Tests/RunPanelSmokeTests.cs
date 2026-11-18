using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.EditHeaders;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.Pipeline;

internal static class RunPanelSmokeTests
{
    static int checks;
    static void Check(bool ok, string message)
    {
        if (!ok) throw new Exception(message);
        checks++;
    }

    static T Find<T>(DependencyObject root, Predicate<T> match) where T : DependencyObject
    {
        var value = root as T;
        if (value != null && match(value)) return value;
        foreach (var child in LogicalTreeHelper.GetChildren(root))
        {
            var dependency = child as DependencyObject;
            if (dependency == null) continue;
            var found = Find(dependency, match);
            if (found != null) return found;
        }
        return null;
    }

    static void Link(PipelineDefinition def, PipelineNode from, PipelineNode to, string port)
    {
        def.Links.Add(new PipelineLink { From = from.Id, To = to.Id, Port = port });
    }

    static void CheckLayout(FrameworkElement editor, string label)
    {
        foreach (double width in new[] { 260.0, 320.0, 480.0 })
        {
            editor.Measure(new Size(width, double.PositiveInfinity));
            editor.Arrange(new Rect(0, 0, width, editor.DesiredSize.Height));
            editor.UpdateLayout();
            CheckBounds(editor, editor, width, label);
        }
    }

    static void CheckBounds(DependencyObject element, FrameworkElement root, double width, string label)
    {
        var current = element as FrameworkElement;
        if (current != null && current.Visibility != Visibility.Visible) return;
        if (current != null && (current is TextBox || current is ComboBox || current is CheckBox ||
            current is RadioButton || current is Button || current is TextBlock))
        {
            var bounds = current.TransformToAncestor(root).TransformBounds(new Rect(current.RenderSize));
            Check(bounds.Left >= -1 && bounds.Right <= width + 1,
                label + " at " + width + "px: " + current.GetType().Name + " " + current.Name +
                " exceeds the editor (" + bounds.Left + ".." + bounds.Right + "). Parent: " +
                (VisualTreeHelper.GetParent(current) == null ? "none" : VisualTreeHelper.GetParent(current).GetType().Name) +
                (current is TextBlock ? "; text: " + ((TextBlock)current).Text : ""));
            if (current is TextBox || current is ComboBox)
                Check(current.ActualWidth > 20, label + ": input has no usable width: " + current.Name);
        }
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(element); i++)
            CheckBounds(VisualTreeHelper.GetChild(element, i), root, width, label);
    }

    [STAThread]
    static int Main()
    {
        string folder = Path.Combine(Path.GetTempPath(), "RunPanel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            string source = Path.Combine(folder, "input.xlsx");
            var table = new DataTable();
            ExcelFile.AddColumn(table, "Name");
            ExcelFile.AddColumn(table, "Amount");
            table.Rows.Add("Alice", 12);
            ExcelFile.Save(table, source);

            // Every kind can be shown or hidden and can construct its shared editor.
            foreach (PipelineStepKind kind in Enum.GetValues(typeof(PipelineStepKind)))
            {
                var def = new PipelineDefinition();
                var load = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 0);
                load.Settings = new LoadStepSettings { Path = source };
                var node = PipelineWork.AddNode(def, kind, 100, 0);
                var ports = PipelineWork.InputPorts(def, node);
                foreach (var port in ports) Link(def, load, node, port);
                if (kind == PipelineStepKind.Load) node.Settings = new LoadStepSettings { Path = source };
                Check(!node.ShowInRunMode, kind + " should default to hidden in Run mode.");
                node.ShowInRunMode = true;
                Check(PipelineWork.Clone(def).Nodes[1].ShowInRunMode, kind + " show flag lost in XML.");
                node.ShowInRunMode = false;
                Check(!PipelineWork.Clone(def).Nodes[1].ShowInRunMode, kind + " hide flag lost in XML.");
                var editor = PipelineStepEditor.Prepare(def, node, settings => { throw new Exception("Constructor applied settings."); });
                Check(editor != null && Find<Button>(editor, b => Equals(b.Content, "Apply to step")) != null,
                    kind + " has no usable editor.");
                CheckLayout(editor, kind.ToString());
                var narrowInput = table.Copy();
                narrowInput.Columns[0].Caption = "Customer name with a long descriptive header " + new string('W', 100);
                var narrowInputs = new DataTable[PipelineWork.InputPorts(def, node).Length];
                var narrowLabels = new string[narrowInputs.Length];
                for (var i = 0; i < ports.Length; i++)
                {
                    narrowInputs[i] = narrowInput;
                    narrowLabels[i] = "From a step with a long title " + new string('W', 100);
                }
                var narrowEditor = PipelineStepEditor.Create(def, node, narrowInputs, narrowLabels, settings => { });
                if (kind == PipelineStepKind.FormatSheet)
                {
                    Expander section;
                    while ((section = Find<Expander>(narrowEditor, e => !e.IsExpanded)) != null)
                        section.IsExpanded = true;
                }
                CheckLayout(narrowEditor, kind + " long labels");
            }

            var dialogCases = new object[][]
            {
                new object[] { "ExcelProcessing.MatchRuleWindow", new object[] { null, table, table, null } },
                new object[] { "FillColumns.MatchKeyWindow", new object[] { null, null } },
                new object[] { "FillColumns.FillRuleWindow", new object[] { null, table, null } },
                new object[] { "Templates.TemplateColumnWindow", new object[] { null, table, null,
                    new List<OfficeWorkAssistant.Features.Templates.TemplateColumn>(), -1 } },
                new object[] { "Highlight.HighlightRuleWindow", new object[] { null, table, null, 1 } },
                new object[] { "FormatSheet.ColumnFormatWindow", new object[] { null, table, 0, null } }
            };
            foreach (var dialogCase in dialogCases)
            {
                var type = typeof(PipelineWork).Assembly.GetType("OfficeWorkAssistant.Features." + dialogCase[0]);
                var window = (Window)Activator.CreateInstance(type, BindingFlags.Instance | BindingFlags.NonPublic,
                    null, (object[])dialogCase[1], null);
                CheckLayout((FrameworkElement)window.Content, (string)dialogCase[0]);
                window.Close();
            }

            var pipeline = new PipelineDefinition();
            var input = PipelineWork.AddNode(pipeline, PipelineStepKind.Load, 0, 0);
            input.Settings = new LoadStepSettings { Path = source };
            input.ShowInRunMode = true;
            var edit = PipelineWork.AddNode(pipeline, PipelineStepKind.EditHeaders, 100, 0);
            edit.Settings = new EditHeadersSettings();
            edit.ShowInRunMode = true;
            var save = PipelineWork.AddNode(pipeline, PipelineStepKind.Save, 200, 0);
            string output = Path.Combine(folder, "output.xlsx");
            save.Settings = new SaveStepSettings { Path = output };
            save.ShowInRunMode = false;
            Link(pipeline, input, edit, "In");
            Link(pipeline, edit, save, "In");
            string savedPath = Path.Combine(folder, "pipeline.xml");
            PipelineWork.SaveFile(pipeline, savedPath);
            string before = File.ReadAllText(savedPath);
            var run = PipelineWork.LoadFile(savedPath);
            var runEdit = run.Nodes[1];
            var headerEditor = (EditHeadersPage)PipelineStepEditor.Prepare(run, runEdit, settings => runEdit.Settings = settings);
            Check(!File.Exists(output), "Opening an editor wrote an output.");
            Check(((EditHeadersSettings)runEdit.Settings).Columns.Count == 0, "Opening an editor changed settings.");
            ((TextBox)headerEditor.FindName("AllFormula")).Text = "upper($Value)";
            Check(string.IsNullOrEmpty(((EditHeadersSettings)runEdit.Settings).Formula), "Unapplied changes leaked into the run.");
            Find<Button>(headerEditor, b => Equals(b.Content, "Apply to step")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Check(((EditHeadersSettings)runEdit.Settings).Formula == "upper($Value)", "Apply did not update run settings.");
            Check(File.ReadAllText(savedPath) == before, "Run edits changed the saved pipeline.");
            var converted = PipelineWork.HeadersToLetters((EditHeadersSettings)runEdit.Settings, table, true);
            Check(converted.Formula == "upper($Value)", "Formula lost in conversion.");
            Check(ExcelFile.Header(EditHeadersWork.Run(table, converted).Columns[0]) == "NAME", "Direct header evaluation failed.");
            var cache = new Dictionary<string, DataTable>();
            var result = PipelineWork.Run(run, cache, null, true);
            string actualHeader = File.Exists(output) ? ExcelFile.Header(ExcelFile.Load(output, null).Table.Columns[0]) : "missing output";
            Check(result.Error == null && actualHeader == "NAME",
                "Run did not use edited settings: " + result.Error + "; header: " + actualHeader +
                "; computed: " + ExcelFile.Header(cache[runEdit.Id].Columns[0]) + "; edits: " + ((EditHeadersSettings)runEdit.Settings).Columns.Count);

            // Form visibility does not change execution; hidden outputs remain valid inputs to validation.
            var page = new RunnerPage();
            typeof(RunnerPage).GetField("_def", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(page, run);
            typeof(RunnerPage).GetMethod("BuildForm", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(page, null);
            var form = (StackPanel)page.FindName("Form");
            var tabs = Find<TabControl>(form, t => true);
            var branchPanel = (StackPanel)((TabItem)tabs.Items[0]).Content;
            Check(Find<TextBlock>(branchPanel, t => t.Text == runEdit.Title) != null, "Visible step has no section.");
            Check(Find<TextBlock>(branchPanel, t => t.Text == save.Title) == null, "Hidden step has an editable section.");

            // Branch tabs separate choices while one Run executes every output.
            var secondSave = PipelineWork.AddNode(run, PipelineStepKind.Save, 300, 0);
            secondSave.Title = "Second result";
            secondSave.ShowInRunMode = true;
            string secondOutput = Path.Combine(folder, "second.xlsx");
            secondSave.Settings = new SaveStepSettings { Path = secondOutput };
            Link(run, runEdit, secondSave, "In");
            string firstOutput = Path.Combine(folder, "first-branch.xlsx");
            ((SaveStepSettings)run.Nodes[2].Settings).Path = firstOutput;
            typeof(RunnerPage).GetMethod("BuildForm", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(page, null);
            tabs = Find<TabControl>(form, t => true);
            Check(tabs.Items.Count == 2, "Expected a tab for each branch.");
            var firstTab = (TabItem)tabs.Items[0];
            var secondTab = (TabItem)tabs.Items[1];
            Check((string)firstTab.Tag == run.Nodes[2].Id && (string)secondTab.Tag == secondSave.Id,
                "Branch tabs do not match the Builder's final steps.");
            Check(Find<TextBlock>((DependencyObject)firstTab.Header, t => t.Text == run.Nodes[2].Title) != null &&
                Find<TextBlock>((DependencyObject)secondTab.Header, t => t.Text == secondSave.Title) != null, "Branch titles are missing.");
            Check(Find<TextBlock>((DependencyObject)firstTab.Content, t => t.Text == secondSave.Title) == null &&
                Find<TextBlock>((DependencyObject)secondTab.Content, t => t.Text == secondSave.Title) != null,
                "A branch shows another branch's settings.");
            var footer = (StackPanel)page.FindName("SharedButtons");
            Check(Find<Button>(footer, b => Equals(b.Content, "Run")) != null &&
                Find<Button>((DependencyObject)firstTab.Content, b => Equals(b.Content, "Run")) == null &&
                Find<Button>((DependencyObject)secondTab.Content, b => Equals(b.Content, "Run")) == null,
                "Run must be shared outside the tabs.");
            string otherSource = Path.Combine(folder, "other-input.xlsx");
            ExcelFile.Save(table, otherSource);
            ((LoadStepSettings)run.Nodes[0].Settings).Path = otherSource;
            typeof(RunnerPage).GetMethod("RefreshInputBlocks", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(page, new object[] { run.Nodes[0].Id });
            Check(Find<TextBox>((DependencyObject)firstTab.Content, t => t.Text == otherSource) != null &&
                Find<TextBox>((DependencyObject)secondTab.Content, t => t.Text == otherSource) != null,
                "Shared input choices are not synchronized across tabs.");
            ((SaveStepSettings)secondSave.Settings).Path = "";
            tabs.SelectedItem = firstTab;
            Check(!(bool)typeof(RunnerPage).GetMethod("Validate", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(page, null) &&
                ReferenceEquals(tabs.SelectedItem, secondTab), "Validation did not reveal the branch needing attention.");
            ((SaveStepSettings)secondSave.Settings).Path = secondOutput;
            typeof(RunnerPage).GetMethod("BuildForm", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(page, null);
            tabs = Find<TabControl>(form, t => true);
            Check((string)((TabItem)tabs.SelectedItem).Tag == secondSave.Id, "Rebuilding the form changed branches.");
            Find<Button>(footer, b => Equals(b.Content, "Run")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Check(File.Exists(firstOutput) && File.Exists(secondOutput), "One Run did not write every branch.");
            Check(ExcelFile.Header(ExcelFile.Load(secondOutput, null).Table.Columns[0]) == "NAME",
                "Shared step settings did not reach both results.");
            Check(File.ReadAllText(savedPath) == before, "Branch run changed the saved pipeline.");

            CheckBuilderBatch(source, otherSource);

            // Pipelines without a panel flag default to hidden for every step kind.
            var legacy = new PipelineDefinition();
            var legacyLoad = PipelineWork.AddNode(legacy, PipelineStepKind.Load, 0, 0);
            var legacyCompare = PipelineWork.AddNode(legacy, PipelineStepKind.Compare, 100, 0);
            legacyCompare.EditMatchRulesInRunMode = true;
            var legacyClean = PipelineWork.AddNode(legacy, PipelineStepKind.CleanText, 200, 0);
            PipelineWork.SaveFile(legacy, savedPath);
            File.WriteAllText(savedPath, Regex.Replace(File.ReadAllText(savedPath), " ShowInRunMode=\"(true|false)\"", ""));
            legacy = PipelineWork.LoadFile(savedPath);
            Check(!legacy.Nodes[0].ShowInRunMode && !legacy.Nodes[1].ShowInRunMode && !legacy.Nodes[2].ShowInRunMode,
                "Missing panel flags must default to hidden.");
            legacy.Nodes[1].ShowInRunMode = false;
            Check(!PipelineWork.Clone(legacy).Nodes[1].ShowInRunMode, "Legacy rule permission overrides hiding.");

            // A downstream editor may preview a disk action, but must not execute it.
            var actionDef = new PipelineDefinition();
            var list = PipelineWork.AddNode(actionDef, PipelineStepKind.ListFolder, 0, 0);
            list.Settings = new FolderScanSettings { Folder = folder };
            var action = PipelineWork.AddNode(actionDef, PipelineStepKind.FileAction, 100, 0);
            string copies = Path.Combine(folder, "copies");
            Directory.CreateDirectory(copies);
            action.Settings = new FileActionSettings { Action = FileActionKind.Copy, DestFolder = copies };
            var downstream = PipelineWork.AddNode(actionDef, PipelineStepKind.CleanText, 200, 0);
            Link(actionDef, list, action, "In");
            Link(actionDef, action, downstream, "In");
            PipelineStepEditor.Prepare(actionDef, downstream, settings => { });
            Check(Directory.GetFiles(copies).Length == 0 && File.Exists(source), "Editor preparation executed a file action.");
            Console.WriteLine("Passed " + checks + " Run panel checks.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { Directory.Delete(folder, true); }
    }

    static void CheckBuilderBatch(string firstPath, string secondPath)
    {
        var def = new PipelineDefinition();
        var first = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 0);
        var second = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 120);
        first.Settings = new LoadStepSettings { Path = firstPath };
        second.Settings = new LoadStepSettings { Path = secondPath };
        var edit = PipelineWork.AddNode(def, PipelineStepKind.EditHeaders, 200, 0);
        edit.Settings = new EditHeadersSettings { Formula = "upper($Value)" };
        var tail = PipelineWork.AddNode(def, PipelineStepKind.EditHeaders, 400, 0);
        tail.Settings = new EditHeadersSettings();
        Link(def, first, edit, "In");

        const BindingFlags privateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
        var page = new PipelinePage();
        var type = typeof(PipelinePage);
        type.GetMethod("Reset", privateInstance).Invoke(page, new object[] { def, null });
        type.GetMethod("Connect", privateInstance).Invoke(page, new object[] { second.Id, edit.Id, "In2", PipelineWork.PortOut });
        Check(PipelineWork.InputPorts(def, edit).Length == 3, "Linking another workbook did not add a free port.");
        type.GetMethod("Select", privateInstance, null, new[] { typeof(string) }, null).Invoke(page, new object[] { edit.Id });
        var result = (bool)type.GetMethod("RunPreview", privateInstance).Invoke(page, new object[] { edit.Id });
        var picker = (ComboBox)page.FindName("OutResult");
        var grid = (DataGrid)page.FindName("GridOut");
        var views = (Dictionary<string, PipelineNodeView>)type.GetField("_views", privateInstance).GetValue(page);
        var outputs = PipelineWork.OutputPorts(edit, PipelineWork.ResultLayout(def));
        var initial = grid.ItemsSource;
        Check(result && picker.Visibility == Visibility.Visible && picker.Items.Count == 2 && initial != null &&
            views[edit.Id].Outputs.Count == 2 && views[edit.Id].OutputPoint(outputs[0]) != views[edit.Id].OutputPoint(outputs[1]),
            "The Builder did not show both results after preview.");
        picker.SelectedIndex = 1;
        Check(!ReferenceEquals(initial, grid.ItemsSource) &&
            ((string)picker.SelectedItem).Contains(Path.GetFileNameWithoutExtension(secondPath)),
            "Choosing the second workbook did not update the Builder preview.");
        type.GetMethod("AddStep", privateInstance).Invoke(page, new object[] { PipelineStepKind.Highlight, null });
        var added = def.Nodes[def.Nodes.Count - 1];
        Check(PipelineWork.InputLink(def, added.Id, "In").FromPort == outputs[0] &&
            PipelineWork.InputLink(def, added.Id, "In2").FromPort == outputs[1],
            "Adding a step from the palette did not wire both outputs separately.");
        type.GetMethod("Connect", privateInstance).Invoke(page, new object[] { edit.Id, tail.Id, "In", outputs[1] });
        type.GetMethod("Select", privateInstance, null, new[] { typeof(string) }, null).Invoke(page, new object[] { tail.Id });
        result = (bool)type.GetMethod("RunPreview", privateInstance).Invoke(page, new object[] { tail.Id });
        var cache = (Dictionary<string, DataTable>)type.GetField("_cache", privateInstance).GetValue(page);
        var items = PipelineWork.ResultItems(cache[tail.Id]);
        Check(result && picker.Visibility == Visibility.Collapsed && items.Count == 1 &&
            items[0].Name == Path.GetFileNameWithoutExtension(secondPath),
            "The Builder did not route only the selected output socket.");
    }
}
