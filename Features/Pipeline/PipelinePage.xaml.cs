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
using OfficeWorkAssistant.Features.Append;
using OfficeWorkAssistant.Features.ArrangeColumns;
using OfficeWorkAssistant.Features.CleanText;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.FormatSheet;
using OfficeWorkAssistant.Features.FormulaGuide;
using OfficeWorkAssistant.Features.Highlight;
using OfficeWorkAssistant.Features.MergeDuplicates;
using OfficeWorkAssistant.Features.RemoveDuplicates;
using OfficeWorkAssistant.Features.SplitCombine;
using OfficeWorkAssistant.Features.Templates;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // The app's main page. Steps are edited in the right sidebar with the feature's editor control;
    // this page owns the graph, the saved pipelines, runs it, and shows each step's result.
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
        string _editorFor;

        PipelineNodeView _dragNode;
        Point _dragOffset;
        bool _dragMoved;
        string _linkFrom;
        Path _tempWire;

        // Palette item being pressed: a click adds the step, a drag places it on the canvas.
        PipelineStepKind? _paletteKind;
        Point _paletteStart;
        bool _listing;

        static readonly PipelineStepKind[][] PaletteGroups =
        {
            // Excel, in the order a pipeline usually runs: load, rows, text and columns, two-table steps, style, save.
            new[] { PipelineStepKind.Load,
                PipelineStepKind.FilterSort, PipelineStepKind.RemoveDuplicates,
                PipelineStepKind.CleanText, PipelineStepKind.Templates, PipelineStepKind.SplitColumn,
                PipelineStepKind.CombineColumns, PipelineStepKind.ArrangeColumns,
                PipelineStepKind.Compare, PipelineStepKind.Append, PipelineStepKind.FillColumns,
                PipelineStepKind.Highlight, PipelineStepKind.FormatSheet,
                PipelineStepKind.Save },
            new[] { PipelineStepKind.ValueList, PipelineStepKind.ListFolder, PipelineStepKind.FindFiles,
                PipelineStepKind.FileAction, PipelineStepKind.MergeFolders }
        };
        static readonly string[] PaletteHeaders =
        {
            "Excel", "Files and folders"
        };

        public PipelinePage()
        {
            InitializeComponent();
            ExcelGrid.Hook(GridOut);
            Board.MouseLeftButtonDown += Board_MouseDown;
            Board.MouseMove += Board_MouseMove;
            Board.MouseLeftButtonUp += Board_MouseUp;
            Board.DragOver += Board_DragOver;
            Board.Drop += Board_Drop;
            PreviewKeyDown += Page_PreviewKeyDown;
            Loaded += Page_Loaded;
            BuildPalette();
            RefreshPipelines();

            // Pick up where the user left off: the pipeline saved most recently.
            var saved = PipelineStore.List();
            if (saved.Count > 0)
            {
                try
                {
                    Reset(PipelineStore.Load(saved[0].Path), saved[0].Path);
                    SetStatus("Opened " + saved[0].Name + ". Point the Load steps at new files if needed, then Run and save files.");
                }
                catch (Exception ex)
                {
                    Reset(new PipelineDefinition(), null);
                    SetStatus("Could not open " + saved[0].Name + ": " + ex.Message);
                }
            }
            else
                Reset(new PipelineDefinition(), null);
        }

        void Page_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateHeader();
        }

        // ---------- saved pipelines ----------

        void RefreshPipelines()
        {
            _listing = true;
            try
            {
                List<SavedPipeline> saved;
                try
                {
                    saved = PipelineStore.List();
                }
                catch (Exception ex)
                {
                    saved = new List<SavedPipeline>();
                    SetStatus("Could not read the pipelines folder: " + ex.Message);
                }
                PipelineList.ItemsSource = saved;
                NoPipelines.Visibility = saved.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
                PipelineList.SelectedItem = null;
                foreach (var p in saved)
                {
                    if (_filePath != null && string.Equals(p.Path, _filePath, StringComparison.OrdinalIgnoreCase))
                        PipelineList.SelectedItem = p;
                }
            }
            finally
            {
                _listing = false;
            }
        }

        void PipelineList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var picked = PipelineList.SelectedItem as SavedPipeline;
            if (_listing || picked == null || (_filePath != null && string.Equals(picked.Path, _filePath, StringComparison.OrdinalIgnoreCase)))
                return;
            if (!ConfirmDiscard())
            {
                RefreshPipelines();
                return;
            }
            OpenSaved(picked.Path);
        }

        void OpenSaved(string path)
        {
            try
            {
                Reset(PipelineStore.Load(path), path);
                SetStatus("Opened " + _def.Name + ". Point the Load steps at new files if needed, then Run and save files.");
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not open pipeline", MessageBoxImage.Error);
            }
            RefreshPipelines();
        }

        void New_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscard())
                return;
            Reset(new PipelineDefinition(), null);
            RefreshPipelines();
            NameBox.Focus();
            SetStatus("New pipeline. Name it at the top, add steps from the left, then Save.");
        }

        void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrent();
        }

        // Saves under the name in the header. False when it was not saved.
        bool SaveCurrent()
        {
            var name = NameBox.Text.Trim();
            var error = PipelineStore.CheckName(name);
            if (error != null)
            {
                Alert(error + " Type it in the box at the top, then Save.", "Name the pipeline", MessageBoxImage.Warning);
                NameBox.Focus();
                return false;
            }
            var path = PipelineStore.PathFor(name);
            if (System.IO.File.Exists(path) && (_filePath == null || !string.Equals(System.IO.Path.GetFullPath(_filePath), path, StringComparison.OrdinalIgnoreCase)) &&
                MessageBox.Show(Window.GetWindow(this), "A pipeline named \"" + name + "\" already exists. Replace it?", "Save",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return false;
            try
            {
                _filePath = PipelineStore.Save(_def, name, _filePath);
                _dirty = false;
                UpdateHeader();
                RefreshPipelines();
                SetStatus("Saved to " + _filePath);
                return true;
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not save pipeline", MessageBoxImage.Error);
                return false;
            }
        }

        void Duplicate_Click(object sender, RoutedEventArgs e)
        {
            var picked = PipelineList.SelectedItem as SavedPipeline;
            if (picked == null)
            {
                SetStatus("Select a saved pipeline to duplicate.");
                return;
            }
            if (!ConfirmDiscard())
                return;
            try
            {
                OpenSaved(PipelineStore.Duplicate(picked.Path));
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not duplicate", MessageBoxImage.Error);
            }
        }

        void DeletePipeline_Click(object sender, RoutedEventArgs e)
        {
            var picked = PipelineList.SelectedItem as SavedPipeline;
            if (picked == null)
            {
                SetStatus("Select a saved pipeline to delete.");
                return;
            }
            if (MessageBox.Show(Window.GetWindow(this), "Move the pipeline \"" + picked.Name + "\" to the Recycle Bin?", "Delete pipeline",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;
            try
            {
                PipelineStore.Delete(picked.Path);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not delete", MessageBoxImage.Error);
                return;
            }
            // The open pipeline stays on the canvas, as an unsaved one.
            if (_filePath != null && string.Equals(_filePath, picked.Path, StringComparison.OrdinalIgnoreCase))
            {
                _filePath = null;
                MarkDirty();
            }
            RefreshPipelines();
            SetStatus("Moved " + picked.Name + " to the Recycle Bin.");
        }

        void Import_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Pipeline files (*.xml)|*.xml",
                Title = "Import a pipeline"
            };
            if (dlg.ShowDialog() != true || !ConfirmDiscard())
                return;
            try
            {
                OpenSaved(PipelineStore.Import(dlg.FileName));
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not import pipeline", MessageBoxImage.Error);
            }
        }

        void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.IO.Directory.CreateDirectory(PipelineStore.Folder);
                System.Diagnostics.Process.Start("explorer.exe", "\"" + PipelineStore.Folder + "\"");
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not open folder", MessageBoxImage.Error);
            }
        }

        void NameBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            NameHint.Visibility = NameBox.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
            if (_filling)
                return;
            _def.Name = NameBox.Text.Trim();
            MarkDirty();
        }

        void FormulaGuide_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new FormulaGuidePage());
        }

        // ---------- step palette ----------

        void BuildPalette()
        {
            for (var g = 0; g < PaletteGroups.Length; g++)
            {
                Palette.Children.Add(new TextBlock
                {
                    Text = PaletteHeaders[g],
                    Foreground = Brushes.Gray,
                    FontSize = 11,
                    Margin = new Thickness(0, g == 0 ? 0 : 10, 0, 4)
                });
                foreach (var kind in PaletteGroups[g])
                    Palette.Children.Add(PaletteItem(kind));
            }
        }

        FrameworkElement PaletteItem(PipelineStepKind kind)
        {
            var text = new StackPanel { Margin = new Thickness(8, 3, 6, 4) };
            text.Children.Add(new TextBlock { Text = PipelineWork.KindLabel(kind), FontWeight = FontWeights.SemiBold });
            text.Children.Add(new TextBlock { Text = PipelineWork.KindDescription(kind), FontSize = 11, Foreground = Brushes.DimGray, TextWrapping = TextWrapping.Wrap });
            var stripe = new Border { Width = 4, Background = PipelineNodeView.AccentOf(kind), CornerRadius = new CornerRadius(3, 0, 0, 3) };
            DockPanel.SetDock(stripe, Dock.Left);
            var icon = PipelineIcons.Create(kind, 30);
            icon.Margin = new Thickness(8, 0, 0, 0);
            icon.VerticalAlignment = VerticalAlignment.Center;
            DockPanel.SetDock(icon, Dock.Left);
            var inner = new DockPanel();
            inner.Children.Add(stripe);
            inner.Children.Add(icon);
            inner.Children.Add(text);
            var item = new Border
            {
                Child = inner,
                Margin = new Thickness(0, 0, 0, 4),
                BorderThickness = new Thickness(1),
                BorderBrush = PaletteBorder,
                Background = Brushes.White,
                CornerRadius = new CornerRadius(3),
                Cursor = Cursors.Hand,
                Tag = kind,
                ToolTip = "Click to add (linked to the selected step), or drag onto the canvas."
            };
            item.MouseEnter += (s, e) => item.Background = PaletteHover;
            item.MouseLeave += (s, e) => item.Background = Brushes.White;
            item.MouseLeftButtonDown += Palette_MouseDown;
            item.MouseMove += Palette_MouseMove;
            item.MouseLeftButtonUp += Palette_MouseUp;
            return item;
        }

        static readonly Brush PaletteBorder = Frozen(Color.FromRgb(0xE0, 0xE6, 0xEB));
        static readonly Brush PaletteHover = Frozen(Color.FromRgb(0xF1, 0xF6, 0xFB));

        void Palette_MouseDown(object sender, MouseButtonEventArgs e)
        {
            _paletteKind = (PipelineStepKind)((FrameworkElement)sender).Tag;
            _paletteStart = e.GetPosition(this);
            e.Handled = true;
        }

        void Palette_MouseMove(object sender, MouseEventArgs e)
        {
            if (_paletteKind == null || e.LeftButton != MouseButtonState.Pressed)
                return;
            var p = e.GetPosition(this);
            if (Math.Abs(p.X - _paletteStart.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(p.Y - _paletteStart.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;
            var kind = _paletteKind.Value;
            _paletteKind = null;
            DragDrop.DoDragDrop((DependencyObject)sender, new DataObject(typeof(PipelineStepKind), kind), DragDropEffects.Copy);
        }

        void Palette_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_paletteKind == null)
                return;
            var kind = _paletteKind.Value;
            _paletteKind = null;
            AddStep(kind, null);
        }

        void Board_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(typeof(PipelineStepKind)) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        void Board_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(PipelineStepKind)))
                return;
            var p = e.GetPosition(Board);
            AddStep((PipelineStepKind)e.Data.GetData(typeof(PipelineStepKind)),
                new Point(p.X - PipelineNodeView.Width / 2, p.Y - PipelineNodeView.Height / 2));
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
            var actions = new List<PipelineNode>();
            foreach (var node in _def.Nodes)
            {
                var s = node.Settings as SaveStepSettings;
                if (s != null)
                    saves.Add(s);
                if (PipelineWork.IsActionKind(node.Kind))
                    actions.Add(node);
            }
            if (saves.Count == 0 && actions.Count == 0)
            {
                Alert("Add a Save file, File action or Merge folders step: those are the steps that write or change files.", "Nothing to save", MessageBoxImage.Warning);
                return;
            }

            // Start clean so every Load step reads its file again.
            _cache.Clear();

            var warnings = new List<string>();
            var replaced = new List<string>();
            foreach (var s in saves)
            {
                if (s.Mode == SaveStepMode.NewFile && !string.IsNullOrWhiteSpace(s.Path) && System.IO.File.Exists(s.Path))
                    replaced.Add(s.Path);
            }
            if (replaced.Count > 0)
                warnings.Add("These files will be replaced:\n" + string.Join("\n", replaced.ToArray()));
            if (actions.Count > 0)
            {
                // A dry run first, so the question can say how many items each action touches.
                if (!RunPreview(null))
                    return;
                var lines = new List<string>();
                foreach (var node in actions)
                {
                    DataTable plan;
                    _cache.TryGetValue(node.Id, out plan);
                    lines.Add(node.Title + ": " + PipelineWork.DescribeAction(node, plan));
                }
                warnings.Add("These file actions will run:\n" + string.Join("\n", lines.ToArray()));
            }
            if (warnings.Count > 0 &&
                MessageBox.Show(Window.GetWindow(this), string.Join("\n\n", warnings.ToArray()) + "\n\nContinue?", "Run and save",
                    MessageBoxButton.YesNo, actions.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            var result = Execute(null, true, true);
            if (result.Error != null)
            {
                Alert(result.Error, "Pipeline stopped", MessageBoxImage.Error);
                return;
            }
            var done = new List<string>();
            if (result.Written.Count > 0)
                done.Add("Saved:\n" + string.Join("\n", result.Written.ToArray()));
            if (result.Actions.Count > 0)
                done.Add("File actions:\n" + string.Join("\n", result.Actions.ToArray()));
            SetStatus("Saved " + result.Written.Count.ToString(CultureInfo.InvariantCulture) + " result(s), ran " +
                result.Actions.Count.ToString(CultureInfo.InvariantCulture) + " file action(s). Select a File action step to see each item's status.");
            Alert(string.Join("\n\n", done.ToArray()), "Done", MessageBoxImage.Information);
        }

        // at: where the step was dropped; null places it next to the selected step.
        void AddStep(PipelineStepKind kind, Point? at)
        {
            PipelineNode from = _selected != null ? PipelineWork.Find(_def, _selected) : null;
            double x;
            double y;
            if (at != null)
            {
                x = Math.Max(0, at.Value.X);
                y = Math.Max(0, at.Value.Y);
            }
            else if (from != null)
            {
                x = from.X + PipelineNodeView.Width + 100;
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
            UpdateEmptyHint();
            Select(node.Id);
            BringIntoView(node);
            if (kind == PipelineStepKind.Load)
                LoadBrowse_Click(null, null);
            else if (PipelineWork.Ports(kind).Length == 0 || from != null)
                SetStatus("Added " + node.Title + ". Set it up in the panel on the right.");
            else
                SetStatus("Added " + node.Title + ". Link an earlier step's right dot to its left dot, then set it up in the panel on the right.");
        }

        void BringIntoView(PipelineNode node)
        {
            var left = node.X - BoardScroll.HorizontalOffset;
            var top = node.Y - BoardScroll.VerticalOffset;
            if (left < 0 || left + PipelineNodeView.Width > BoardScroll.ViewportWidth)
                BoardScroll.ScrollToHorizontalOffset(Math.Max(0, node.X - 40));
            if (top < 0 || top + PipelineNodeView.Height > BoardScroll.ViewportHeight)
                BoardScroll.ScrollToVerticalOffset(Math.Max(0, node.Y - 40));
        }

        void UpdateEmptyHint()
        {
            EmptyHint.Visibility = _def.Nodes.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
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
            UpdateEmptyHint();
            ShowSelection();
            _filling = true;
            NameBox.Text = _def.Name ?? "";
            _filling = false;
            UpdateHeader();
            BoardScroll.ScrollToHorizontalOffset(0);
            BoardScroll.ScrollToVerticalOffset(0);
        }

        void AddView(PipelineNode node)
        {
            var view = new PipelineNodeView(node, PipelineWork.InputPorts(_def, node));
            view.Box.MouseLeftButtonDown += Node_MouseDown;
            view.Output.MouseLeftButtonDown += Output_MouseDown;
            view.Box.Tag = node.Id;
            _views[node.Id] = view;
            Board.Children.Add(view.Root);
        }

        // An Append step's ports follow its links, so its box is built again when they change.
        void RebuildView(string id)
        {
            PipelineNodeView old;
            if (!_views.TryGetValue(id, out old) || old.Node.Kind != PipelineStepKind.Append)
                return;
            Board.Children.Remove(old.Root);
            AddView(old.Node);
            var view = _views[id];
            view.Selected = _selected == id;
            view.SetState(_cache.ContainsKey(id) ? PipelineNodeState.Ok : PipelineNodeState.Idle, null);
        }

        void Node_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var id = (string)((FrameworkElement)sender).Tag;
            var view = _views[id];
            e.Handled = true;
            Board.Focus();
            Select(id);
            if (e.ClickCount == 2 && view.Node.Kind == PipelineStepKind.Load)
            {
                LoadBrowse_Click(null, null);
                return;
            }
            if (e.ClickCount == 2 && view.Node.Kind == PipelineStepKind.Save)
            {
                SaveBrowse_Click(null, null);
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
            var ports = PipelineWork.InputPorts(_def, node);
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
            RebuildView(to);
            RedrawWires();
            if (_selected == to)
            {
                // Its editor was built for the old input.
                _editorFor = null;
                ShowSelection();
            }
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
            var ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
            if (ctrl && e.Key == Key.S)
            {
                SaveCurrent();
                e.Handled = true;
                return;
            }
            if (ctrl && e.Key == Key.N)
            {
                New_Click(null, null);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.F5)
            {
                RunPreview_Click(null, null);
                e.Handled = true;
                return;
            }
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
                var to = _selectedLink.To;
                _def.Links.Remove(_selectedLink);
                Invalidate(to);
                _selectedLink = null;
                PipelineWork.CompactPorts(_def, to);
                RebuildView(to);
                if (_editorFor == to)
                    _editorFor = null;
            }
            else if (_selected != null)
            {
                foreach (var id in PipelineWork.Downstream(_def, _selected))
                    Forget(id);
                Board.Children.Remove(_views[_selected].Root);
                _views.Remove(_selected);
                var fed = PipelineWork.RemoveNode(_def, _selected);
                _selected = null;
                foreach (var id in fed)
                    RebuildView(id);
            }
            else
                return;
            MarkDirty();
            RedrawWires();
            UpdateEmptyHint();
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
                    ShowEditor(null);
                    return;
                }

                StepKind.Text = PipelineWork.KindLabel(node.Kind).ToUpperInvariant();
                StepKind.Foreground = PipelineNodeView.AccentOf(node.Kind);
                StepAbout.Text = PipelineWork.KindDescription(node.Kind) +
                    (PipelineWork.IsActionKind(node.Kind) ? " Preview only shows the plan; files change on Run and save files." : "");
                StepTitle.Text = node.Title;
                LoadPanel.Visibility = node.Kind == PipelineStepKind.Load ? Visibility.Visible : Visibility.Collapsed;
                SavePanel.Visibility = node.Kind == PipelineStepKind.Save ? Visibility.Visible : Visibility.Collapsed;

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
                ShowEditor(node);
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

        // ---------- editing a step in the sidebar ----------

        // The editor stays while the same step is selected, so edits are not lost when a run
        // redraws the selection. It is built again when another step is selected.
        void ShowEditor(PipelineNode node)
        {
            if (node == null || node.Kind == PipelineStepKind.Load || node.Kind == PipelineStepKind.Save)
            {
                EditorHost.Child = null;
                _editorFor = null;
                return;
            }
            if (_editorFor == node.Id && EditorHost.Child != null)
                return;
            _editorFor = node.Id;
            EditorHost.Child = BuildEditor(node);
        }

        static FrameworkElement Notice(string text)
        {
            return new TextBlock { Text = text, Foreground = Brushes.DimGray, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0) };
        }

        FrameworkElement BuildEditor(PipelineNode node)
        {
            // The editor needs the real inputs, so run the steps before this one first.
            var ports = PipelineWork.InputPorts(_def, node);
            var inputs = new DataTable[ports.Length];
            var labels = new string[ports.Length];
            for (var i = 0; i < ports.Length; i++)
            {
                var link = PipelineWork.InputLink(_def, node.Id, ports[i]);
                // A list can be typed in without an input; Append always has a free port.
                if (link == null && (node.Kind == PipelineStepKind.ValueList || node.Kind == PipelineStepKind.Append))
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
                inputs[i] = _cache[link.From];
                labels[i] = "From step: " + PipelineWork.Find(_def, link.From).Title;
            }

            switch (node.Kind)
            {
                case PipelineStepKind.FilterSort:
                {
                    var input = inputs[0];
                    var s = node.Settings as FilterSortSettings;
                    return new FilterSortPage(input, labels[0], s == null ? null : PipelineWork.FilterSortToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.FilterSortToNames(r, input)));
                }
                case PipelineStepKind.Templates:
                {
                    var input = inputs[0];
                    var t = node.Settings as TemplateDefinition;
                    return new TemplatesPage(input, labels[0], t == null ? null : PipelineWork.TemplateToLetters(t, input, false),
                        r => UseSettings(node, PipelineWork.TemplateToNames(r, input)));
                }
                case PipelineStepKind.Compare:
                    return new ExcelProcessingPage(inputs[0], labels[0], inputs[1], labels[1],
                        PipelineWork.Clone(node.Settings as ExcelProcessingSettings), r => UseSettings(node, r));
                case PipelineStepKind.ValueList:
                    return new ValueListPage(inputs[0], labels[0], PipelineWork.Clone(node.Settings as ValueListSettings), r => UseSettings(node, r));
                case PipelineStepKind.ListFolder:
                    return new FindFilesPage(PipelineWork.Clone(node.Settings as FolderScanSettings), r => UseSettings(node, r));
                case PipelineStepKind.FindFiles:
                    return new FindFilesPage(inputs[0], labels[0], PipelineWork.Clone(node.Settings as FindFilesSettings), r => UseSettings(node, r));
                case PipelineStepKind.FileAction:
                    return new FileActionPage(inputs[0], labels[0], PipelineWork.Clone(node.Settings as FileActionSettings), r => UseSettings(node, r));
                case PipelineStepKind.MergeFolders:
                    return new MergeDuplicatesPage(PipelineWork.Clone(node.Settings as MergeFoldersSettings), r => UseSettings(node, r));
                case PipelineStepKind.CleanText:
                {
                    var input = inputs[0];
                    var s = node.Settings as CleanTextSettings;
                    return new CleanTextPage(input, labels[0], s == null ? null : PipelineWork.CleanTextToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.CleanTextToNames(r, input)));
                }
                case PipelineStepKind.SplitColumn:
                {
                    var input = inputs[0];
                    var s = node.Settings as SplitColumnSettings;
                    return new SplitColumnPage(input, labels[0], s == null ? null : PipelineWork.SplitToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.SplitToNames(r, input)));
                }
                case PipelineStepKind.CombineColumns:
                {
                    var input = inputs[0];
                    var s = node.Settings as CombineColumnsSettings;
                    return new CombineColumnsPage(input, labels[0], s == null ? null : PipelineWork.CombineToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.CombineToNames(r, input)));
                }
                case PipelineStepKind.RemoveDuplicates:
                {
                    var input = inputs[0];
                    var s = node.Settings as RemoveDuplicatesSettings;
                    return new RemoveDuplicatesPage(input, labels[0], s == null ? null : PipelineWork.RemoveDuplicatesToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.RemoveDuplicatesToNames(r, input)));
                }
                case PipelineStepKind.ArrangeColumns:
                {
                    var input = inputs[0];
                    var s = node.Settings as ArrangeColumnsSettings;
                    return new ArrangeColumnsPage(input, labels[0], s == null ? null : PipelineWork.ArrangeToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.ArrangeToNames(r, input)));
                }
                case PipelineStepKind.Append:
                {
                    var tables = new List<DataTable>();
                    var shown = new List<string>();
                    var names = new List<string>();
                    for (var i = 0; i < ports.Length; i++)
                    {
                        if (inputs[i] == null)
                            continue;
                        tables.Add(inputs[i]);
                        shown.Add(labels[i]);
                        names.Add(PipelineWork.Find(_def, PipelineWork.InputLink(_def, node.Id, ports[i]).From).Title);
                    }
                    if (tables.Count == 0)
                        return Notice("Link the steps whose rows you want to stack: drag from the right dot of each one to this step's left dot.");
                    return new AppendPage(tables, shown, names, PipelineWork.Clone(node.Settings as AppendSettings), r => UseSettings(node, r));
                }
                case PipelineStepKind.Highlight:
                {
                    var input = inputs[0];
                    var s = node.Settings as HighlightSettings;
                    return new HighlightPage(input, labels[0], s == null ? null : PipelineWork.HighlightToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.HighlightToNames(r, input)));
                }
                case PipelineStepKind.FormatSheet:
                {
                    var input = inputs[0];
                    var s = node.Settings as FormatSheetSettings;
                    return new FormatSheetPage(input, labels[0], s == null ? null : PipelineWork.FormatSheetToLetters(s, input, false),
                        r => UseSettings(node, PipelineWork.FormatSheetToNames(r, input)));
                }
                default:
                    return new FillColumnsPage(inputs[0], labels[0], inputs[1], labels[1],
                        PipelineWork.Clone(node.Settings as FillColumnsSettings), r => UseSettings(node, r));
            }
        }

        // The editor's Apply to step button: keep the settings and show what the step now gives.
        void UseSettings(PipelineNode node, object settings)
        {
            node.Settings = settings;
            StepChanged(node);
            if (RunPreview(node.Id))
                SetStatus("Applied to " + node.Title + ".");
        }
        // ---------- running ----------

        // Runs targetId and what it needs (everything when null). False when a step failed.
        bool RunPreview(string targetId)
        {
            var result = Execute(targetId, false, true);
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

        // selectFailed: select the step that failed and redraw the side panel. False while the side panel is
        // itself being built, which must not select another step or build itself again.
        PipelineRunResult Execute(string targetId, bool writeFiles, bool selectFailed)
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
                if (selectFailed)
                    Select(result.FailedNodeId);
                SetStatus(PipelineWork.Find(_def, result.FailedNodeId).Title + ": " + result.Error);
            }
            else if (selectFailed)
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
            DirtyMark.Visibility = _dirty ? Visibility.Visible : Visibility.Collapsed;
            NameHint.Visibility = NameBox.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
            var window = Window.GetWindow(this);
            if (window != null)
                window.Title = (string.IsNullOrWhiteSpace(_def.Name) ? "Untitled pipeline" : _def.Name) + (_dirty ? " *" : "") + " - Office Work Assistant";
        }

        // True when it is fine to replace what is on the canvas.
        bool ConfirmDiscard()
        {
            if (!_dirty || _def.Nodes.Count == 0)
                return true;
            var answer = MessageBox.Show(Window.GetWindow(this), "Save the changes to \"" +
                (string.IsNullOrWhiteSpace(NameBox.Text) ? "Untitled pipeline" : NameBox.Text.Trim()) + "\" first?", "Unsaved changes",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Yes)
                return SaveCurrent();
            return answer == MessageBoxResult.No;
        }

        // The window asks before closing with unsaved changes.
        public bool CanClose()
        {
            return ConfirmDiscard();
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
