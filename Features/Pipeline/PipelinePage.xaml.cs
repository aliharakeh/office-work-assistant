using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Microsoft.Win32;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.Templates;
using OfficeWorkAssistant.Views;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // The canvas. Steps are edited on the feature pages themselves (pipeline mode); this page
    // owns the graph, runs it, and shows each step's result. The Frame keeps this instance
    // alive while a feature page is open, so all set-up happens in the constructor.
    public partial class PipelinePage : Page
    {
        const int PreviewRows = 1000;

        static readonly Brush WireBrush = Frozen(Color.FromRgb(0x60, 0x7D, 0x8B));
        static readonly Brush WireSelectedBrush = Frozen(Color.FromRgb(0x1E, 0x88, 0xE5));

        PipelineDefinition _def = new PipelineDefinition();
        readonly Dictionary<string, DataTable> _cache = new Dictionary<string, DataTable>();
        readonly Dictionary<string, PipelineNodeView> _views = new Dictionary<string, PipelineNodeView>();
        readonly Dictionary<PipelineLink, Path> _wires = new Dictionary<PipelineLink, Path>();
        readonly Dictionary<string, string[]> _sheets = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        string _selected;
        PipelineLink _selectedLink;
        string _filePath;
        bool _dirty;
        bool _filling;
        string _previewOnReturn;

        PipelineNodeView _dragNode;
        Point _dragOffset;
        bool _dragMoved;
        string _linkFrom;
        Path _tempWire;

        public PipelinePage()
        {
            InitializeComponent();
            ExcelGrid.Hook(GridOut);
            Board.MouseLeftButtonDown += Board_MouseDown;
            Board.MouseMove += Board_MouseMove;
            Board.MouseLeftButtonUp += Board_MouseUp;
            PreviewKeyDown += Page_PreviewKeyDown;
            Loaded += Page_Loaded;
            ShowSelection();
            UpdateHeader();
        }

        // Fires again when a feature page goes back here: show what the edited step now gives.
        void Page_Loaded(object sender, RoutedEventArgs e)
        {
            if (_previewOnReturn == null)
                return;
            var id = _previewOnReturn;
            _previewOnReturn = null;
            Select(id);
            RunPreview(id);
        }

        // ---------- toolbar ----------

        void Home_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscard())
                return;
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
            else if (NavigationService != null)
                NavigationService.Navigate(new HomePage());
        }

        void New_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscard())
                return;
            Reset(new PipelineDefinition(), null);
        }

        void Open_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscard())
                return;
            var dlg = new OpenFileDialog
            {
                Filter = "Pipeline files (*.xml)|*.xml",
                Title = "Open pipeline"
            };
            if (dlg.ShowDialog() != true)
                return;
            try
            {
                Reset(PipelineWork.LoadFile(dlg.FileName), dlg.FileName);
                SetStatus("Opened. Point the Load steps at this month's files if needed, then Run and save files.");
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not open pipeline", MessageBoxImage.Error);
            }
        }

        void Save_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Pipeline files (*.xml)|*.xml",
                FileName = _filePath != null ? System.IO.Path.GetFileName(_filePath) : "pipeline.xml"
            };
            if (_filePath != null)
                dlg.InitialDirectory = System.IO.Path.GetDirectoryName(_filePath);
            if (dlg.ShowDialog() != true)
                return;
            try
            {
                _def.Name = System.IO.Path.GetFileNameWithoutExtension(dlg.FileName);
                PipelineWork.SaveFile(_def, dlg.FileName);
                _filePath = dlg.FileName;
                _dirty = false;
                UpdateHeader();
                SetStatus("Pipeline saved.");
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not save pipeline", MessageBoxImage.Error);
            }
        }

        void RunPreview_Click(object sender, RoutedEventArgs e)
        {
            if (_def.Nodes.Count == 0)
            {
                SetStatus("Add a step first.");
                return;
            }
            if (RunPreview(_selected))
                SetStatus(_selected == null ? "Every step ran. Click a step to see its result." : "Preview ready.");
        }

        void RunAll_Click(object sender, RoutedEventArgs e)
        {
            var saves = new List<SaveStepSettings>();
            foreach (var node in _def.Nodes)
            {
                var s = node.Settings as SaveStepSettings;
                if (s != null)
                    saves.Add(s);
            }
            if (saves.Count == 0)
            {
                Alert("Add a Save file step to write a result.", "Nothing to save", MessageBoxImage.Warning);
                return;
            }
            var replaced = new List<string>();
            foreach (var s in saves)
            {
                if (s.Mode == SaveStepMode.NewFile && !string.IsNullOrWhiteSpace(s.Path) && System.IO.File.Exists(s.Path))
                    replaced.Add(s.Path);
            }
            if (replaced.Count > 0 &&
                MessageBox.Show(Window.GetWindow(this), "These files will be replaced:\n\n" + string.Join("\n", replaced.ToArray()) +
                    "\n\nContinue?", "Run and save", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            // Start clean so every Load step reads its file again.
            _cache.Clear();
            var result = Execute(null, true);
            if (result.Error != null)
            {
                Alert(result.Error, "Pipeline stopped", MessageBoxImage.Error);
                return;
            }
            SetStatus("Saved " + result.Written.Count.ToString(CultureInfo.InvariantCulture) + " result(s).");
            Alert("Saved:\n\n" + string.Join("\n", result.Written.ToArray()), "Done", MessageBoxImage.Information);
        }

        void AddStep_Click(object sender, RoutedEventArgs e)
        {
            var kind = (PipelineStepKind)Enum.Parse(typeof(PipelineStepKind), (string)((Button)sender).Tag);
            PipelineNode from = _selected != null ? PipelineWork.Find(_def, _selected) : null;
            double x;
            double y;
            if (from != null)
            {
                x = from.X + PipelineNodeView.Width + 60;
                y = from.Y;
            }
            else
            {
                var n = _def.Nodes.Count % 8;
                x = BoardScroll.HorizontalOffset + 30 + n * 24;
                y = BoardScroll.VerticalOffset + 30 + n * 24;
            }
            var node = PipelineWork.AddNode(_def, kind, x, y);
            AddView(node);

            // Building a chain: the new step reads from the selected one.
            var ports = PipelineWork.Ports(kind);
            if (from != null && ports.Length > 0)
                Connect(from.Id, node.Id, ports[0]);

            MarkDirty();
            GrowBoard();
            Select(node.Id);
            if (kind == PipelineStepKind.Load)
                LoadBrowse_Click(null, null);
        }

        // ---------- canvas ----------

        void Reset(PipelineDefinition def, string path)
        {
            _def = def;
            _filePath = path;
            _dirty = false;
            _cache.Clear();
            _views.Clear();
            _wires.Clear();
            _selected = null;
            _selectedLink = null;
            Board.Children.Clear();
            foreach (var node in _def.Nodes)
                AddView(node);
            RedrawWires();
            GrowBoard();
            ShowSelection();
            UpdateHeader();
        }

        void AddView(PipelineNode node)
        {
            var view = new PipelineNodeView(node);
            view.Box.MouseLeftButtonDown += Node_MouseDown;
            view.Output.MouseLeftButtonDown += Output_MouseDown;
            view.Box.Tag = node.Id;
            _views[node.Id] = view;
            Board.Children.Add(view.Root);
        }

        void Node_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var id = (string)((FrameworkElement)sender).Tag;
            var view = _views[id];
            e.Handled = true;
            Board.Focus();
            Select(id);
            if (e.ClickCount == 2)
            {
                EditStep(view.Node);
                return;
            }
            _dragNode = view;
            _dragMoved = false;
            _dragOffset = e.GetPosition(view.Root);
            view.Root.CaptureMouse();
        }

        void Output_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var port = (PipelinePort)((FrameworkElement)sender).Tag;
            e.Handled = true;
            _linkFrom = port.NodeId;
            _tempWire = new Path
            {
                Stroke = WireSelectedBrush,
                StrokeThickness = 2,
                StrokeDashArray = new DoubleCollection { 4, 3 },
                IsHitTestVisible = false
            };
            Panel.SetZIndex(_tempWire, 2);
            Board.Children.Add(_tempWire);
            Board.CaptureMouse();
        }

        void Board_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // Only clicks on empty canvas get here; steps and links handle their own.
            Board.Focus();
            Select(null);
        }

        void Board_MouseMove(object sender, MouseEventArgs e)
        {
            var p = e.GetPosition(Board);
            if (_dragNode != null)
            {
                var node = _dragNode.Node;
                node.X = Math.Max(0, p.X - _dragOffset.X);
                node.Y = Math.Max(0, p.Y - _dragOffset.Y);
                _dragNode.Place();
                UpdateWireShapes();
                _dragMoved = true;
            }
            else if (_tempWire != null)
            {
                _tempWire.Data = Bezier(_views[_linkFrom].OutputPoint(), p);
            }
        }

        void Board_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_dragNode != null)
            {
                _dragNode.Root.ReleaseMouseCapture();
                _dragNode = null;
                if (_dragMoved)
                {
                    MarkDirty();
                    GrowBoard();
                }
                return;
            }
            if (_tempWire == null)
                return;

            var from = _linkFrom;
            Board.Children.Remove(_tempWire);
            _tempWire = null;
            _linkFrom = null;
            Board.ReleaseMouseCapture();

            var target = PortAt(e.GetPosition(Board));
            if (target == null)
            {
                SetStatus("Drop the link on a step's left dot.");
                return;
            }
            Connect(from, target.NodeId, target.Port);
        }

        PipelinePort PortAt(Point p)
        {
            var hit = VisualTreeHelper.HitTest(Board, p);
            var element = hit != null ? hit.VisualHit as FrameworkElement : null;
            var port = element != null ? element.Tag as PipelinePort : null;
            if (port != null && port.Port != PipelinePort.Out)
                return port;
            // Dropped on a step's box: use its first free input (or its only one).
            var nodeId = element != null ? FindNodeId(element) : null;
            if (nodeId == null)
                return null;
            var node = PipelineWork.Find(_def, nodeId);
            var ports = PipelineWork.Ports(node.Kind);
            if (ports.Length == 0)
                return null;
            foreach (var name in ports)
            {
                if (PipelineWork.InputLink(_def, nodeId, name) == null)
                    return new PipelinePort { NodeId = nodeId, Port = name };
            }
            return ports.Length == 1 ? new PipelinePort { NodeId = nodeId, Port = ports[0] } : null;
        }

        string FindNodeId(DependencyObject element)
        {
            while (element != null && element != Board)
            {
                foreach (var view in _views.Values)
                {
                    if (view.Root == element)
                        return view.Node.Id;
                }
                element = VisualTreeHelper.GetParent(element);
            }
            return null;
        }

        void Connect(string from, string to, string port)
        {
            if (from == to)
            {
                SetStatus("A step cannot feed itself.");
                return;
            }
            if (PipelineWork.WouldCycle(_def, from, to))
            {
                SetStatus("That link would make a loop, so it was not added.");
                return;
            }
            var old = PipelineWork.InputLink(_def, to, port);
            if (old != null)
            {
                if (old.From == from)
                    return;
                _def.Links.Remove(old);
            }
            _def.Links.Add(new PipelineLink { From = from, To = to, Port = port });
            Invalidate(to);
            MarkDirty();
            RedrawWires();
            var node = PipelineWork.Find(_def, to);
            SetStatus("Linked " + PipelineWork.Find(_def, from).Title + " to " + node.Title +
                (port == PipelineWork.PortIn ? "." : " (" + PipelineWork.PortLabel(node.Kind, port) + ")."));
        }

        void RedrawWires()
        {
            foreach (var path in _wires.Values)
                Board.Children.Remove(path);
            _wires.Clear();
            foreach (var link in _def.Links)
            {
                var path = new Path { StrokeThickness = 3, Cursor = Cursors.Hand };
                var captured = link;
                path.MouseLeftButtonDown += (s, e) =>
                {
                    e.Handled = true;
                    SelectLink(captured);
                };
                Panel.SetZIndex(path, 0);
                _wires[link] = path;
                Board.Children.Add(path);
            }
            UpdateWireShapes();
        }

        void UpdateWireShapes()
        {
            foreach (var pair in _wires)
            {
                PipelineNodeView from;
                PipelineNodeView to;
                if (!_views.TryGetValue(pair.Key.From, out from) || !_views.TryGetValue(pair.Key.To, out to))
                    continue;
                pair.Value.Data = Bezier(from.OutputPoint(), to.InputPoint(pair.Key.Port));
                pair.Value.Stroke = pair.Key == _selectedLink ? WireSelectedBrush : WireBrush;
            }
        }

        static Geometry Bezier(Point a, Point b)
        {
            var dx = Math.Max(40, Math.Abs(b.X - a.X) / 2);
            var figure = new PathFigure { StartPoint = a };
            figure.Segments.Add(new BezierSegment(new Point(a.X + dx, a.Y), new Point(b.X - dx, b.Y), b, true));
            var geometry = new PathGeometry();
            geometry.Figures.Add(figure);
            return geometry;
        }

        void GrowBoard()
        {
            double w = 2400;
            double h = 1600;
            foreach (var node in _def.Nodes)
            {
                w = Math.Max(w, node.X + PipelineNodeView.Width + 400);
                h = Math.Max(h, node.Y + PipelineNodeView.Height + 400);
            }
            Board.Width = w;
            Board.Height = h;
        }

        void Page_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Delete || Keyboard.FocusedElement is TextBox)
                return;
            if (_selected == null && _selectedLink == null)
                return;
            DeleteSelected();
            e.Handled = true;
        }

        void Delete_Click(object sender, RoutedEventArgs e)
        {
            DeleteSelected();
        }

        void DeleteSelected()
        {
            if (_selectedLink != null)
            {
                _def.Links.Remove(_selectedLink);
                Invalidate(_selectedLink.To);
                _selectedLink = null;
            }
            else if (_selected != null)
            {
                foreach (var id in PipelineWork.Downstream(_def, _selected))
                    Forget(id);
                Board.Children.Remove(_views[_selected].Root);
                _views.Remove(_selected);
                PipelineWork.RemoveNode(_def, _selected);
                _selected = null;
            }
            else
                return;
            MarkDirty();
            RedrawWires();
            ShowSelection();
        }

        // ---------- selection and side panel ----------

        void Select(string id)
        {
            _selected = id;
            _selectedLink = null;
            foreach (var view in _views.Values)
                view.Selected = view.Node.Id == id;
            UpdateWireShapes();
            ShowSelection();
        }

        void SelectLink(PipelineLink link)
        {
            _selected = null;
            _selectedLink = link;
            foreach (var view in _views.Values)
                view.Selected = false;
            UpdateWireShapes();
            ShowSelection();
        }

        void ShowSelection()
        {
            _filling = true;
            try
            {
                var node = _selected != null ? PipelineWork.Find(_def, _selected) : null;
                NoSelection.Visibility = node == null && _selectedLink == null ? Visibility.Visible : Visibility.Collapsed;
                StepPanel.Visibility = node != null ? Visibility.Visible : Visibility.Collapsed;
                LinkPanel.Visibility = _selectedLink != null ? Visibility.Visible : Visibility.Collapsed;

                if (_selectedLink != null)
                {
                    var to = PipelineWork.Find(_def, _selectedLink.To);
                    LinkText.Text = PipelineWork.Find(_def, _selectedLink.From).Title + "  ->  " + to.Title +
                        (_selectedLink.Port == PipelineWork.PortIn ? "" : " (" + PipelineWork.PortLabel(to.Kind, _selectedLink.Port) + ")");
                }

                if (node == null)
                {
                    ShowResult(null);
                    return;
                }

                StepKind.Text = PipelineWork.KindLabel(node.Kind).ToUpperInvariant();
                StepTitle.Text = node.Title;
                LoadPanel.Visibility = node.Kind == PipelineStepKind.Load ? Visibility.Visible : Visibility.Collapsed;
                SavePanel.Visibility = node.Kind == PipelineStepKind.Save ? Visibility.Visible : Visibility.Collapsed;
                EditBtn.Visibility = node.Kind == PipelineStepKind.Load || node.Kind == PipelineStepKind.Save
                    ? Visibility.Collapsed : Visibility.Visible;
                StepError.Text = _views[node.Id].Box.ToolTip as string ?? "";

                var load = node.Settings as LoadStepSettings;
                if (load != null)
                {
                    LoadPath.Text = load.Path ?? "";
                    LoadSheet.ItemsSource = SheetsOf(load);
                    LoadSheet.SelectedItem = load.Sheet;
                }
                var save = node.Settings as SaveStepSettings;
                if (save != null)
                {
                    SavePath.Text = save.Path ?? "";
                    SaveSheet.Text = save.SheetName ?? "";
                    SaveNewFile.IsChecked = save.Mode == SaveStepMode.NewFile;
                    SaveNewSheet.IsChecked = save.Mode == SaveStepMode.NewSheet;
                }
                ShowResult(node);
            }
            finally
            {
                _filling = false;
            }
        }

        void ShowResult(PipelineNode node)
        {
            DataTable table;
            if (node == null || !_cache.TryGetValue(node.Id, out table))
            {
                GridOut.ItemsSource = null;
                InfoOut.Text = node == null ? "Select a step to see its result." : "Click Run preview to see this step's result.";
                return;
            }
            var shown = table;
            if (table.Rows.Count > PreviewRows)
            {
                shown = table.Clone();
                for (var i = 0; i < PreviewRows; i++)
                    shown.ImportRow(table.Rows[i]);
            }
            GridOut.ItemsSource = shown.DefaultView;
            var rows = table.Rows.Count.ToString("N0", CultureInfo.InvariantCulture);
            InfoOut.Text = node.Title + ": " + rows + " rows, " + table.Columns.Count.ToString(CultureInfo.InvariantCulture) + " columns" +
                (shown != table ? " (showing the first " + PreviewRows.ToString("N0", CultureInfo.InvariantCulture) + ")" : "");
        }

        // Sheet names for the combo; read once per file because opening a workbook is slow.
        string[] SheetsOf(LoadStepSettings load)
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
                StepError.Text = ex.Message;
                return string.IsNullOrEmpty(load.Sheet) ? new string[0] : new[] { load.Sheet };
            }
        }

        PipelineNode SelectedNode()
        {
            return _selected != null ? PipelineWork.Find(_def, _selected) : null;
        }

        void StepTitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            var node = SelectedNode();
            if (_filling || node == null)
                return;
            node.Title = StepTitle.Text;
            _views[node.Id].Refresh();
            MarkDirty();
        }

        void LoadBrowse_Click(object sender, RoutedEventArgs e)
        {
            var node = SelectedNode();
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
            StepChanged(node);
            ShowSelection();
        }

        void LoadSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var node = SelectedNode();
            var load = node != null ? node.Settings as LoadStepSettings : null;
            var sheet = LoadSheet.SelectedItem as string;
            if (_filling || load == null || sheet == null || sheet == load.Sheet)
                return;
            load.Sheet = sheet;
            StepChanged(node);
        }

        void SaveBrowse_Click(object sender, RoutedEventArgs e)
        {
            var node = SelectedNode();
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
            StepChanged(node);
            ShowSelection();
        }

        void SaveMode_Changed(object sender, RoutedEventArgs e)
        {
            var node = SelectedNode();
            var save = node != null ? node.Settings as SaveStepSettings : null;
            if (_filling || save == null)
                return;
            save.Mode = SaveNewSheet.IsChecked == true ? SaveStepMode.NewSheet : SaveStepMode.NewFile;
            StepChanged(node);
        }

        void SaveSheet_TextChanged(object sender, TextChangedEventArgs e)
        {
            var node = SelectedNode();
            var save = node != null ? node.Settings as SaveStepSettings : null;
            if (_filling || save == null)
                return;
            save.SheetName = SaveSheet.Text;
            StepChanged(node);
        }

        void StepChanged(PipelineNode node)
        {
            Invalidate(node.Id);
            _views[node.Id].Refresh();
            MarkDirty();
        }

        // ---------- editing a step on its feature page ----------

        void EditStep_Click(object sender, RoutedEventArgs e)
        {
            var node = SelectedNode();
            if (node != null)
                EditStep(node);
        }

        void EditStep(PipelineNode node)
        {
            if (node.Kind == PipelineStepKind.Load)
            {
                LoadBrowse_Click(null, null);
                return;
            }
            if (node.Kind == PipelineStepKind.Save)
            {
                SaveBrowse_Click(null, null);
                return;
            }

            // The page needs the real inputs, so run the steps before this one first.
            var ports = PipelineWork.Ports(node.Kind);
            var inputs = new DataTable[ports.Length];
            var labels = new string[ports.Length];
            for (var i = 0; i < ports.Length; i++)
            {
                var link = PipelineWork.InputLink(_def, node.Id, ports[i]);
                if (link == null)
                {
                    Alert("Connect " + PipelineWork.PortLabel(node.Kind, ports[i]) + " of this step first: drag from the right dot of an earlier step.",
                        "Not connected", MessageBoxImage.Warning);
                    return;
                }
                if (!RunPreview(link.From))
                    return;
                inputs[i] = _cache[link.From];
                labels[i] = "From step: " + PipelineWork.Find(_def, link.From).Title;
            }

            Page page;
            switch (node.Kind)
            {
                case PipelineStepKind.FilterSort:
                {
                    var input = inputs[0];
                    var s = node.Settings as FilterSortSettings;
                    page = new FilterSortPage(input, labels[0], s == null ? null : PipelineWork.FilterSortToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.FilterSortToNames(r, input)));
                    break;
                }
                case PipelineStepKind.Templates:
                {
                    var input = inputs[0];
                    var t = node.Settings as TemplateDefinition;
                    page = new TemplatesPage(input, labels[0], t == null ? null : PipelineWork.TemplateToLetters(t, input, false),
                        r => UseSettings(node, PipelineWork.TemplateToNames(r, input)));
                    break;
                }
                case PipelineStepKind.Compare:
                    page = new ExcelProcessingPage(inputs[0], labels[0], inputs[1], labels[1],
                        PipelineWork.Clone(node.Settings as ExcelProcessingSettings), r => UseSettings(node, r));
                    break;
                default:
                    page = new FillColumnsPage(inputs[0], labels[0], inputs[1], labels[1],
                        PipelineWork.Clone(node.Settings as FillColumnsSettings), r => UseSettings(node, r));
                    break;
            }
            NavigationService.Navigate(page);
        }

        void UseSettings(PipelineNode node, object settings)
        {
            node.Settings = settings;
            StepChanged(node);
            _previewOnReturn = node.Id;
        }

        // ---------- running ----------

        // Runs targetId and what it needs (everything when null). False when a step failed.
        bool RunPreview(string targetId)
        {
            var result = Execute(targetId, false);
            if (result.Error == null)
                return true;
            // A failing step is already selected and named in the status line.
            if (result.FailedNodeId == null)
            {
                SetStatus(result.Error);
                Alert(result.Error, "Pipeline", MessageBoxImage.Warning);
            }
            return false;
        }

        PipelineRunResult Execute(string targetId, bool writeFiles)
        {
            PipelineRunResult result;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                result = PipelineWork.Run(_def, _cache, targetId, writeFiles);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }

            foreach (var view in _views.Values)
            {
                var id = view.Node.Id;
                if (id == result.FailedNodeId)
                    view.SetState(PipelineNodeState.Error, result.Error);
                else if (_cache.ContainsKey(id))
                    view.SetState(PipelineNodeState.Ok, null);
                else
                    view.SetState(PipelineNodeState.Idle, null);
            }
            if (result.FailedNodeId != null)
            {
                Select(result.FailedNodeId);
                SetStatus(PipelineWork.Find(_def, result.FailedNodeId).Title + ": " + result.Error);
            }
            else
                ShowSelection();
            return result;
        }

        // A step changed: its result and everything after it are out of date.
        void Invalidate(string id)
        {
            foreach (var d in PipelineWork.Downstream(_def, id))
                Forget(d);
        }

        void Forget(string id)
        {
            _cache.Remove(id);
            PipelineNodeView view;
            if (_views.TryGetValue(id, out view))
                view.SetState(PipelineNodeState.Idle, null);
        }

        // ---------- helpers ----------

        void MarkDirty()
        {
            if (_dirty)
                return;
            _dirty = true;
            UpdateHeader();
        }

        void UpdateHeader()
        {
            HeaderTitle.Text = "Pipeline" + (_filePath != null ? " - " + System.IO.Path.GetFileName(_filePath) : "") + (_dirty ? " *" : "");
        }

        bool ConfirmDiscard()
        {
            if (!_dirty || _def.Nodes.Count == 0)
                return true;
            return MessageBox.Show(Window.GetWindow(this), "The pipeline has unsaved changes. Discard them?", "Pipeline",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
        }

        void SetStatus(string text)
        {
            Status.Text = text;
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }

        static Brush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
