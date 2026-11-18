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
    // The right sidebar: one tab for each branch of the pipeline (a last step and every step that feeds it),
    // and in each tab one section for each step with all of that step's settings, in running order.
    public partial class PipelinePage
    {
        // A step that feeds two branches appears in both tabs, each with its own editor.
        sealed class StepTab
        {
            public string SinkId;
            public TabItem Item;
            public TextBlock Header;
            public StackPanel Panel;
            public readonly List<string> NodeIds = new List<string>();
            public readonly List<string> Keys = new List<string>();
            public bool Built;
        }

        sealed class StepSection
        {
            public PipelineNode Node;
            public Border Root;
            public TextBlock Kind;
            public TextBlock Error;
            public Border Body;
            public bool Placeholder;
            public CheckBox RunRulesToggle;
            // Why the editor's settings could not be applied yet (shown instead of a dialog).
            public string Problem;
            public TextBox LoadPath;
            public ComboBox LoadSheet;
            public CheckBox LoadFirstRowIsData;
            public TextBox SavePath;
            public TextBox SaveSheet;
            public RadioButton SaveNewFile;
            public RadioButton SaveNewSheet;
        }

        static readonly Brush SectionBrush = Frozen(Color.FromRgb(0xDD, 0xE3, 0xE8));
        static readonly Brush SectionSelectedBrush = Frozen(Color.FromRgb(0x1E, 0x88, 0xE5));
        static readonly Brush ErrorBrush = Frozen(Color.FromRgb(0xC6, 0x28, 0x28));

        readonly List<StepTab> _tabs = new List<StepTab>();
        readonly Dictionary<string, StepSection> _sections = new Dictionary<string, StepSection>();
        bool _rebuilding;

        // ---------- tabs ----------

        // The graph changed. changed: steps whose result is now different, so the editors of everything after them
        // are made again; every other editor keeps what was typed in it.
        void RebuildSidebar(ICollection<string> changed)
        {
            var previous = CurrentTab();
            var keepSink = previous != null ? previous.SinkId : null;

            var live = new HashSet<string>();
            foreach (var node in _def.Nodes)
                live.Add(node.Id);
            var stale = new HashSet<string>();
            if (changed != null)
            {
                foreach (var id in changed)
                    stale.UnionWith(PipelineWork.Downstream(_def, id));
            }
            var drop = new List<string>();
            foreach (var pair in _sections)
            {
                if (!live.Contains(pair.Value.Node.Id) || stale.Contains(pair.Value.Node.Id))
                    drop.Add(pair.Key);
            }
            foreach (var key in drop)
                _sections.Remove(key);

            _rebuilding = true;
            try
            {
                foreach (var tab in _tabs)
                    tab.Panel.Children.Clear();
                _tabs.Clear();
                StepTabs.Items.Clear();

                var order = PipelineWork.Order(_def);
                var feeds = new HashSet<string>();
                foreach (var link in _def.Links)
                    feeds.Add(link.From);
                var seen = new Dictionary<string, int>();
                foreach (var sink in _def.Nodes)
                {
                    if (feeds.Contains(sink.Id))
                        continue;
                    var up = PipelineWork.Upstream(_def, sink.Id);
                    var tab = new StepTab { SinkId = sink.Id, Panel = new StackPanel { Margin = new Thickness(0, 8, 4, 0) } };
                    foreach (var node in order)
                    {
                        if (!up.Contains(node.Id))
                            continue;
                        int n;
                        seen.TryGetValue(node.Id, out n);
                        seen[node.Id] = n + 1;
                        tab.NodeIds.Add(node.Id);
                        tab.Keys.Add(n == 0 ? node.Id : node.Id + "#" + n.ToString(CultureInfo.InvariantCulture));
                    }
                    tab.Header = new TextBlock { Text = sink.Title, MaxWidth = 140, TextTrimming = TextTrimming.CharacterEllipsis };
                    tab.Item = new TabItem
                    {
                        Header = tab.Header,
                        Tag = tab,
                        ToolTip = "The steps that lead to " + sink.Title,
                        Content = new ScrollViewer
                        {
                            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                            Content = tab.Panel
                        }
                    };
                    _tabs.Add(tab);
                    StepTabs.Items.Add(tab.Item);
                }
            }
            finally
            {
                _rebuilding = false;
            }

            NoSelection.Visibility = _tabs.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            StepTabs.Visibility = _tabs.Count == 0 ? Visibility.Collapsed : Visibility.Visible;

            // Stay on the same branch, or on one that has the selected step.
            StepTab show = null;
            foreach (var tab in _tabs)
            {
                if (tab.SinkId == keepSink)
                    show = tab;
            }
            if (show == null && _selected != null)
                show = TabWith(_selected, null);
            if (show == null && _tabs.Count > 0)
                show = _tabs[0];
            if (show != null)
            {
                StepTabs.SelectedItem = show.Item;
                EnsureBuilt(show);
            }
        }

        StepTab CurrentTab()
        {
            var item = StepTabs.SelectedItem as TabItem;
            return item != null ? item.Tag as StepTab : null;
        }

        // The tab that shows the step, the preferred one when it does.
        StepTab TabWith(string nodeId, StepTab prefer)
        {
            if (prefer != null && prefer.NodeIds.Contains(nodeId))
                return prefer;
            foreach (var tab in _tabs)
            {
                if (tab.NodeIds.Contains(nodeId))
                    return tab;
            }
            return null;
        }

        void StepTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Combo boxes inside the sections raise this too.
            if (_rebuilding || !ReferenceEquals(e.OriginalSource, StepTabs))
                return;
            var tab = CurrentTab();
            if (tab != null)
                EnsureBuilt(tab);
        }

        // A tab's sections are made when it is first shown: every editor needs its step's input, which means running the steps before it.
        void EnsureBuilt(StepTab tab)
        {
            if (tab.Built)
                return;
            tab.Built = true;
            tab.Panel.Children.Clear();
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                for (var i = 0; i < tab.Keys.Count; i++)
                {
                    StepSection section;
                    if (!_sections.TryGetValue(tab.Keys[i], out section))
                    {
                        section = CreateSection(PipelineWork.Find(_def, tab.NodeIds[i]));
                        _sections[tab.Keys[i]] = section;
                    }
                    section.Kind.Text = (i + 1).ToString(CultureInfo.InvariantCulture) + ".  " + section.Node.Title;
                    tab.Panel.Children.Add(section.Root);
                }
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
            UpdateSections();
        }

        // ---------- one step's section ----------

        static bool HasEditor(PipelineStepKind kind)
        {
            return kind != PipelineStepKind.Load && kind != PipelineStepKind.Save;
        }

        StepSection CreateSection(PipelineNode node)
        {
            var section = new StepSection { Node = node };
            var panel = new StackPanel();
            panel.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/OfficeWorkAssistant;component/Excel/EditorStyles.xaml", UriKind.Relative) });

            var top = new StackPanel { Margin = new Thickness(0, 0, 0, 2) };
            var delete = new Button { Content = "Delete step", Padding = new Thickness(8, 1, 8, 1), Margin = new Thickness(0, 4, 0, 4), HorizontalAlignment = HorizontalAlignment.Left };
            delete.Click += (s, e) =>
            {
                Select(node.Id, false);
                DeleteSelected();
            };
            section.Kind = new TextBlock
            {
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = PipelineNodeView.AccentOf(node.Kind),
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            top.Children.Add(section.Kind);
            top.Children.Add(delete);
            panel.Children.Add(top);
            // Every step may expose its settings for a run.
            {
                var allowRules = new CheckBox
                {
                    Content = "Show in Run panel",
                    IsChecked = node.ShowInRunMode,
                    Margin = new Thickness(0, 4, 0, 6)
                };
                section.RunRulesToggle = allowRules;
                allowRules.Click += (s, e) =>
                {
                    node.ShowInRunMode = allowRules.IsChecked == true;
                    foreach (var other in _sections.Values)
                    {
                        var toggle = other.RunRulesToggle;
                        if (other.Node.Id == node.Id && toggle != null)
                            toggle.IsChecked = node.ShowInRunMode;
                    }
                    MarkDirty();
                };
                panel.Children.Add(allowRules);
            }

            panel.Children.Add(new TextBlock
            {
                Text = PipelineWork.KindDescription(node.Kind) +
                    (PipelineWork.IsActionKind(node.Kind) ? " Preview only shows the plan; files change on Run and save files." : ""),
                Foreground = Brushes.Gray,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 6)
            });

            section.Error = new TextBlock { Foreground = ErrorBrush, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 6) };
            if (node.Kind == PipelineStepKind.Load)
                AddLoadFields(section, panel);
            else if (node.Kind == PipelineStepKind.Save)
                AddSaveFields(section, panel);
            else if (node.Kind == PipelineStepKind.SavePdf ||
                (node.Kind == PipelineStepKind.LoadPdf && PipelineWork.InputLink(_def, node.Id, PipelineWork.PortIn) == null))
            {
                var browse = new Button
                {
                    Content = node.Kind == PipelineStepKind.LoadPdf ? "Choose PDF..." : "Choose output PDF...",
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Padding = new Thickness(8, 2, 8, 2),
                    Margin = new Thickness(0, 0, 0, 6)
                };
                browse.Click += (s, e) => BrowsePdf(section);
                panel.Children.Add(browse);
            }
            else if (node.Kind == PipelineStepKind.SaveText)
            {
                var browse = new Button { Content = "Choose output file...", HorizontalAlignment = HorizontalAlignment.Left,
                    Padding = new Thickness(8, 2, 8, 2), Margin = new Thickness(0, 0, 0, 6) };
                browse.Click += (s, e) => BrowseSaveText(node);
                panel.Children.Add(browse);
            }
            panel.Children.Add(section.Error);

            section.Body = new Border();
            panel.Children.Add(section.Body);

            section.Root = new Border
            {
                BorderThickness = new Thickness(2),
                BorderBrush = SectionBrush,
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(8),
                Margin = new Thickness(0, 0, 0, 10),
                Child = panel
            };
            // Clicking anywhere in a section makes it the selected step: its result shows below the canvas.
            section.Root.PreviewMouseLeftButtonDown += (s, e) =>
            {
                if (_selected != node.Id)
                    Select(node.Id, false);
            };

            FillFields(section);
            if (HasEditor(node.Kind))
                RebuildEditor(section);
            return section;
        }

        void RebuildEditor(StepSection section)
        {
            if (!HasEditor(section.Node.Kind))
                return;
            var editor = BuildEditor(section);
            section.Placeholder = editor is TextBlock;
            // A grid-heavy editor needs a height to fill; the others are as tall as they need.
            section.Body.Height = !section.Placeholder && IsTall(section.Node.Kind) ? 560 : double.NaN;
            section.Body.Child = editor;
            section.Problem = null;
            if (!section.Placeholder)
                HookAutoApply(section, editor);
        }

        // There is no Apply button: a change made in the editor is applied by itself a moment after the last edit.
        void HookAutoApply(StepSection section, FrameworkElement editor)
        {
            var apply = FindApplyButton(editor);
            if (apply == null)
                return;
            apply.Visibility = Visibility.Collapsed;

            var armed = false;
            var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(600) };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                if (!ReferenceEquals(section.Body.Child, editor))
                    return;
                Notices.Quiet = true;
                Notices.Last = null;
                try
                {
                    apply.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                }
                finally
                {
                    Notices.Quiet = false;
                }
                section.Problem = Notices.Last;
                UpdateSections();
            };
            Action poke = () =>
            {
                if (!armed)
                    return;
                timer.Stop();
                timer.Start();
            };
            // The editor fills itself in while it loads; only what happens after that is the user's edit.
            editor.Loaded += (s, e) => Dispatcher.BeginInvoke(new Action(() => armed = true), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            editor.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new TextChangedEventHandler((s, e) => poke()));
            editor.AddHandler(System.Windows.Controls.Primitives.Selector.SelectionChangedEvent, new SelectionChangedEventHandler((s, e) => poke()));
            editor.AddHandler(Keyboard.LostKeyboardFocusEvent, new KeyboardFocusChangedEventHandler((s, e) => poke()));
            RoutedEventHandler clicked = (s, e) =>
            {
                if (!ReferenceEquals(e.OriginalSource, apply))
                    poke();
            };
            editor.AddHandler(System.Windows.Controls.Primitives.ButtonBase.ClickEvent, clicked);
            editor.AddHandler(System.Windows.Controls.Primitives.ToggleButton.CheckedEvent, clicked);
            editor.AddHandler(System.Windows.Controls.Primitives.ToggleButton.UncheckedEvent, clicked);
        }

        static Button FindApplyButton(DependencyObject parent)
        {
            foreach (var child in LogicalTreeHelper.GetChildren(parent))
            {
                var button = child as Button;
                if (button != null && "Apply to step".Equals(button.Content as string))
                    return button;
                var inner = child as DependencyObject;
                var found = inner != null ? FindApplyButton(inner) : null;
                if (found != null)
                    return found;
            }
            return null;
        }

        static string SettingsXml(object settings)
        {
            if (settings == null)
                return null;
            using (var writer = new System.IO.StringWriter(CultureInfo.InvariantCulture))
            {
                new System.Xml.Serialization.XmlSerializer(settings.GetType()).Serialize(writer, settings);
                return writer.ToString();
            }
        }

        static bool IsTall(PipelineStepKind kind)
        {
            return kind == PipelineStepKind.ValueList;
        }

        List<StepSection> SectionsOf(string nodeId)
        {
            var found = new List<StepSection>();
            foreach (var section in _sections.Values)
            {
                if (section.Node.Id == nodeId)
                    found.Add(section);
            }
            return found;
        }

        // Highlight the selected step's section and show each step's problem.
        void UpdateSections()
        {
            foreach (var section in _sections.Values)
            {
                section.Root.BorderBrush = section.Node.Id == _selected ? SectionSelectedBrush : SectionBrush;
                PipelineNodeView view;
                var failed = _views.TryGetValue(section.Node.Id, out view) ? view.Box.ToolTip as string : null;
                section.Error.Text = !string.IsNullOrEmpty(section.Problem) ? section.Problem : failed ?? "";
            }
        }

        void RevealSection(string nodeId)
        {
            var tab = TabWith(nodeId, CurrentTab());
            if (tab == null)
                return;
            if (!ReferenceEquals(StepTabs.SelectedItem, tab.Item))
                StepTabs.SelectedItem = tab.Item;
            EnsureBuilt(tab);
            StepSection section;
            if (_sections.TryGetValue(tab.Keys[tab.NodeIds.IndexOf(nodeId)], out section))
                Dispatcher.BeginInvoke(new Action(() => section.Root.BringIntoView()), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        // ---------- Load and Save fields ----------

        void AddLoadFields(StepSection section, StackPanel panel)
        {
            var node = section.Node;
            panel.Children.Add(new TextBlock { Text = "File", Margin = new Thickness(0, 0, 0, 2) });
            var row = new StackPanel { Margin = new Thickness(0, 0, 0, 6) };
            var browse = new Button { Content = "Browse...", Width = 72, Margin = new Thickness(0, 4, 0, 0), HorizontalAlignment = HorizontalAlignment.Left };
            browse.Click += (s, e) => BrowseLoad(node);
            section.LoadPath = new TextBox { IsReadOnly = true };
            row.Children.Add(section.LoadPath);
            row.Children.Add(browse);
            panel.Children.Add(row);

            panel.Children.Add(new TextBlock { Text = "Sheet", Margin = new Thickness(0, 0, 0, 2) });
            section.LoadSheet = new ComboBox { Margin = new Thickness(0, 0, 0, 8) };
            section.LoadSheet.SelectionChanged += (s, e) =>
            {
                var load = node.Settings as LoadStepSettings;
                var sheet = section.LoadSheet.SelectedItem as string;
                if (_filling || load == null || sheet == null || sheet == load.Sheet)
                    return;
                load.Sheet = sheet;
                StepChanged(node, null);
                RefreshFields(node.Id);
            };
            panel.Children.Add(section.LoadSheet);
            section.LoadFirstRowIsData = new CheckBox { Content = "No header row (first row is data)", Margin = new Thickness(0, 0, 0, 8) };
            RoutedEventHandler headersChanged = (s, e) =>
            {
                var load = node.Settings as LoadStepSettings;
                if (_filling || load == null) return;
                load.FirstRowIsData = section.LoadFirstRowIsData.IsChecked == true;
                StepChanged(node, null);
                RefreshFields(node.Id);
            };
            section.LoadFirstRowIsData.Checked += headersChanged;
            section.LoadFirstRowIsData.Unchecked += headersChanged;
            panel.Children.Add(section.LoadFirstRowIsData);
        }

        void AddSaveFields(StepSection section, StackPanel panel)
        {
            var node = section.Node;
            var group = "save" + node.Id;
            section.SaveNewFile = new RadioButton { Content = "New file (replaces it)", GroupName = group, Margin = new Thickness(0, 0, 0, 2) };
            section.SaveNewSheet = new RadioButton { Content = "New sheet in a workbook", GroupName = group, Margin = new Thickness(0, 0, 0, 6) };
            RoutedEventHandler mode = (s, e) =>
            {
                var save = node.Settings as SaveStepSettings;
                if (_filling || save == null)
                    return;
                save.Mode = section.SaveNewSheet.IsChecked == true ? SaveStepMode.NewSheet : SaveStepMode.NewFile;
                StepChanged(node, null);
                RefreshFields(node.Id);
            };
            section.SaveNewFile.Checked += mode;
            section.SaveNewSheet.Checked += mode;
            panel.Children.Add(section.SaveNewFile);
            panel.Children.Add(section.SaveNewSheet);

            panel.Children.Add(new TextBlock { Text = "File (.xlsx)", Margin = new Thickness(0, 0, 0, 2) });
            var row = new StackPanel { Margin = new Thickness(0, 0, 0, 6) };
            var browse = new Button { Content = "Browse...", Width = 72, Margin = new Thickness(0, 4, 0, 0), HorizontalAlignment = HorizontalAlignment.Left };
            browse.Click += (s, e) => BrowseSave(node);
            section.SavePath = new TextBox { IsReadOnly = true };
            row.Children.Add(section.SavePath);
            row.Children.Add(browse);
            panel.Children.Add(row);

            panel.Children.Add(new TextBlock { Text = "Sheet name", Margin = new Thickness(0, 0, 0, 2) });
            section.SaveSheet = new TextBox { Margin = new Thickness(0, 0, 0, 4) };
            section.SaveSheet.TextChanged += (s, e) =>
            {
                var save = node.Settings as SaveStepSettings;
                if (_filling || save == null)
                    return;
                save.SheetName = section.SaveSheet.Text;
                StepChanged(node, null);
                RefreshFields(node.Id);
            };
            panel.Children.Add(section.SaveSheet);
            panel.Children.Add(new TextBlock
            {
                Text = "With several results, result.xlsx becomes result-January.xlsx, result-February.xlsx, and so on. A new sheet must not exist yet.",
                Foreground = Brushes.Gray,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });
        }

        // Put the step's settings in every copy of its Load or Save fields.
        void RefreshFields(string nodeId)
        {
            foreach (var section in SectionsOf(nodeId))
                FillFields(section);
        }

        void FillFields(StepSection section)
        {
            var was = _filling;
            _filling = true;
            try
            {
                var load = section.Node.Settings as LoadStepSettings;
                if (load != null && section.LoadPath != null)
                {
                    section.LoadPath.Text = load.Path ?? "";
                    section.LoadFirstRowIsData.IsChecked = load.FirstRowIsData;
                    var sheets = SheetsOf(load, section);
                    var shown = section.LoadSheet.ItemsSource as string[];
                    if (shown == null || !SameSheets(shown, sheets))
                        section.LoadSheet.ItemsSource = sheets;
                    section.LoadSheet.SelectedItem = load.Sheet;
                }
                var save = section.Node.Settings as SaveStepSettings;
                if (save != null && section.SavePath != null)
                {
                    section.SavePath.Text = save.Path ?? "";
                    if (section.SaveSheet.Text != (save.SheetName ?? ""))
                        section.SaveSheet.Text = save.SheetName ?? "";
                    section.SaveNewFile.IsChecked = save.Mode == SaveStepMode.NewFile;
                    section.SaveNewSheet.IsChecked = save.Mode == SaveStepMode.NewSheet;
                }
            }
            finally
            {
                _filling = was;
            }
        }

        static bool SameSheets(string[] a, string[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (var i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i])
                    return false;
            }
            return true;
        }

        // Sheet names for the combo; read once per file because opening a workbook is slow.
        string[] SheetsOf(LoadStepSettings load, StepSection section)
        {
            if (string.IsNullOrWhiteSpace(load.Path))
                return new string[0];
            string[] sheets;
            if (_sheets.TryGetValue(load.Path, out sheets))
                return sheets;
            try
            {
                sheets = ExcelFile.Load(load.Path, null, 1).Sheets;
                _sheets[load.Path] = sheets;
                return sheets;
            }
            catch (Exception ex)
            {
                section.Error.Text = ex.Message;
                return string.IsNullOrEmpty(load.Sheet) ? new string[0] : new[] { load.Sheet };
            }
        }

        void BrowseLoad(PipelineNode node)
        {
            var load = node != null ? node.Settings as LoadStepSettings : null;
            if (load == null)
                return;
            var dlg = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                Title = "Choose the file for " + node.Title
            };
            if (dlg.ShowDialog() != true)
                return;
            try
            {
                var loaded = ExcelFile.Load(dlg.FileName, null, 1);
                _sheets[dlg.FileName] = loaded.Sheets;
                load.Path = dlg.FileName;
                load.Sheet = loaded.Sheet;
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not open file", MessageBoxImage.Error);
                return;
            }
            StepChanged(node, null);
            RefreshFields(node.Id);
        }

        void BrowsePdf(StepSection section)
        {
            const string filter = "PDF files (*.pdf)|*.pdf";
            var load = section.Body.Child as LoadPdfPage;
            if (load != null)
            {
                var dialog = new OpenFileDialog { Filter = filter, CheckFileExists = true };
                if (dialog.ShowDialog() == true)
                    load.SetPath(dialog.FileName);
                return;
            }
            var save = section.Body.Child as SavePdfPage;
            if (save == null)
            {
                Alert("Connect a working PDF input first.", "Save PDF", MessageBoxImage.Warning);
                return;
            }
            var output = new SaveFileDialog { Filter = filter, DefaultExt = ".pdf", AddExtension = true, OverwritePrompt = false };
            if (output.ShowDialog() == true)
                save.SetPath(output.FileName);
        }

        void BrowseSaveText(PipelineNode node)
        {
            SaveTextPage editor = null;
            foreach (var section in SectionsOf(node.Id))
            {
                editor = section.Body.Child as SaveTextPage;
                if (editor != null)
                    break;
            }
            if (editor == null)
            {
                Alert("Connect a working input first, then choose the output file.", "Save text / JSON", MessageBoxImage.Warning);
                return;
            }
            bool json = SaveTextWork.IsJson(editor.SelectedFormat);
            var settings = node.Settings as SaveTextSettings;
            var dialog = new SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|JSON files (*.json)|*.json|All files (*.*)|*.*",
                FilterIndex = json ? 2 : 1,
                DefaultExt = json ? ".json" : ".txt",
                AddExtension = true,
                OverwritePrompt = false,
                Title = "Where should " + node.Title + " write?",
                FileName = settings == null || string.IsNullOrWhiteSpace(settings.Path)
                    ? (json ? "result.json" : "result.txt") : System.IO.Path.GetFileName(settings.Path)
            };
            if (dialog.ShowDialog() == true)
                editor.SetOutputPath(dialog.FileName);
        }

        void BrowseSave(PipelineNode node)
        {
            var save = node != null ? node.Settings as SaveStepSettings : null;
            if (save == null)
                return;
            var dlg = new SaveFileDialog
            {
                Filter = "Excel workbook (*.xlsx)|*.xlsx",
                Title = "Where should " + node.Title + " write?",
                // Adding a sheet to an existing workbook is not a replace.
                OverwritePrompt = false,
                FileName = string.IsNullOrWhiteSpace(save.Path) ? "result.xlsx" : System.IO.Path.GetFileName(save.Path)
            };
            if (dlg.ShowDialog() != true)
                return;
            save.Path = dlg.FileName;
            StepChanged(node, null);
            RefreshFields(node.Id);
        }

        // A step's settings changed: its result and everything after it are out of date, and the steps after it
        // see a different table, so their editors are made again from it (keep: the editor that was just applied).
        void StepChanged(PipelineNode node, StepSection keep)
        {
            Invalidate(node.Id);
            _views[node.Id].Refresh();
            MarkDirty();
            var down = PipelineWork.Downstream(_def, node.Id);
            if (node.Kind == PipelineStepKind.ValueList)
            {
                PipelineWork.RemoveDisconnectedOutputs(_def);
                if (RefreshPortViews())
                {
                    RedrawWires();
                    RebuildSidebar(down);
                    UpdateSections();
                    return;
                }
            }
            foreach (var section in new List<StepSection>(_sections.Values))
            {
                if (section != keep && down.Contains(section.Node.Id))
                    RebuildEditor(section);
            }
            UpdateSections();
        }

        // The editor's settings, applied by itself after each edit: keep them and show what the step now gives.
        void UseSettings(StepSection section, object settings)
        {
            var node = section.Node;
            Select(node.Id, false);
            // Selecting in a list or leaving a box changes nothing.
            if (node.Settings != null && SettingsXml(node.Settings) == SettingsXml(settings) && _cache.ContainsKey(node.Id))
                return;
            node.Settings = settings;
            StepChanged(node, section);
            if (RunPreview(node.Id))
                SetStatus("Updated " + node.Title + ".");
        }

        // ---------- selection ----------

        void Select(string id)
        {
            Select(id, true);
        }

        // reveal: bring the step's section into view in the sidebar (not needed when it was clicked there).
        void Select(string id, bool reveal)
        {
            _selected = id;
            _selectedLink = null;
            foreach (var view in _views.Values)
                view.Selected = view.Node.Id == id;
            UpdateWireShapes();
            ShowSelection(reveal);
        }

        void SelectLink(PipelineLink link)
        {
            _selected = null;
            _selectedLink = link;
            foreach (var view in _views.Values)
                view.Selected = false;
            UpdateWireShapes();
            ShowSelection(false);
        }

        void ShowSelection(bool reveal)
        {
            LinkPanel.Visibility = _selectedLink != null ? Visibility.Visible : Visibility.Collapsed;
            if (_selectedLink != null)
            {
                var to = PipelineWork.Find(_def, _selectedLink.To);
                var from = _views[_selectedLink.From];
                var output = from.Outputs.Count > 1
                    ? string.IsNullOrEmpty(_selectedLink.FromPort) ? " (all results)" :
                        " (result " + from.OutputNumber(_selectedLink.FromPort).ToString(CultureInfo.InvariantCulture) + ")"
                    : "";
                LinkText.Text = from.Node.Title + output + "  ->  " + to.Title +
                    (_selectedLink.Port == PipelineWork.PortIn ? "" : " (" + PipelineWork.PortLabel(to.Kind, _selectedLink.Port) + ")");
            }

            var node = _selected != null ? PipelineWork.Find(_def, _selected) : null;
            ShowResult(node);
            UpdateSections();
            if (node != null && reveal)
                RevealSection(node.Id);
        }

        void ShowResult(PipelineNode node)
        {
            DataTable table;
            var previous = OutResult.Tag as DataTable;
            var previousIndex = OutResult.SelectedIndex;
            OutResult.Tag = null;
            OutResult.ItemsSource = null;
            OutResult.Visibility = Visibility.Collapsed;
            if (node == null || !_cache.TryGetValue(node.Id, out table))
            {
                GridOut.ItemsSource = null;
                InfoOut.Text = node == null ? "Select a step to see its result." : "Click Run preview to see this step's result.";
                return;
            }
            var items = PipelineWork.ResultItems(table);
            if (items.Count > 1)
            {
                var labels = new List<string>();
                for (int i = 0; i < items.Count; i++)
                    labels.Add((i + 1).ToString(CultureInfo.InvariantCulture) + ". " +
                        (items[i].TargetPath == null ? items[i].Name : System.IO.Path.GetFileName(items[i].TargetPath)));
                OutResult.Tag = table;
                OutResult.ItemsSource = labels;
                OutResult.Visibility = Visibility.Visible;
                OutResult.SelectedIndex = previous == table && previousIndex >= 0 ? previousIndex : 0;
            }
            else
                ShowResultTable(node, items[0].Table, 1);
        }

        void OutResult_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var table = OutResult.Tag as DataTable;
            var node = _selected == null ? null : PipelineWork.Find(_def, _selected);
            if (table == null || node == null || OutResult.SelectedIndex < 0)
                return;
            var items = PipelineWork.ResultItems(table);
            ShowResultTable(node, items[OutResult.SelectedIndex].Table, items.Count);
        }

        void ShowResultTable(PipelineNode node, DataTable table, int results)
        {
            var shown = table;
            if (table.Rows.Count > PreviewRows)
            {
                shown = table.Clone();
                for (var i = 0; i < PreviewRows; i++)
                    shown.ImportRow(table.Rows[i]);
            }
            GridOut.ItemsSource = shown.DefaultView;
            var rows = table.Rows.Count.ToString("N0", CultureInfo.InvariantCulture);
            InfoOut.Text = node.Title + (results > 1 ? " (" + results.ToString(CultureInfo.InvariantCulture) + " results)" : "") +
                ": " + rows + " rows, " + table.Columns.Count.ToString(CultureInfo.InvariantCulture) + " columns" +
                (shown != table ? " (showing the first " + PreviewRows.ToString("N0", CultureInfo.InvariantCulture) + ")" : "");
        }

        static FrameworkElement Notice(string text)
        {
            return new TextBlock { Text = text, Foreground = Brushes.DimGray, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0) };
        }

        // ---------- the step's editor ----------

        FrameworkElement BuildEditor(StepSection section)
        {
            var node = section.Node;
            // The editor needs the real inputs, so run the steps before this one first.
            var ports = PipelineWork.InputPorts(_def, node);
            var inputs = new DataTable[ports.Length];
            var labels = new string[ports.Length];
            for (var i = 0; i < ports.Length; i++)
            {
                var link = PipelineWork.InputLink(_def, node.Id, ports[i]);
                // A list can be typed in without an input; Append always has a free port.
                if (link == null && PipelineWork.IsOptionalInput(node.Kind, ports[i]))
                    continue;
                if (link == null)
                    return Notice("Connect " + PipelineWork.PortLabel(node.Kind, ports[i]) +
                        " of this step first: drag from the right dot of an earlier step to this step's left dot.");
                var result = Execute(link.From, false, false);
                if (result.Error != null)
                {
                    var failed = result.FailedNodeId != null ? PipelineWork.Find(_def, result.FailedNodeId) : null;
                    return Notice("This step cannot be set up until the steps before it work.\n\n" +
                        (failed != null ? failed.Title + ": " : "") + result.Error);
                }
                try
                {
                    inputs[i] = PipelineWork.LinkedResult(_def, link, _cache[link.From], _outputLayout);
                }
                catch (InvalidOperationException ex)
                {
                    return Notice(ex.Message);
                }
                labels[i] = "From step: " + PipelineWork.Find(_def, link.From).Title;
            }

            return PipelineStepEditor.Create(_def, node, inputs, labels, r => UseSettings(section, r));
        }
    }
}
