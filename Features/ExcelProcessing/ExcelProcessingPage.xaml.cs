using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;
using WorkAssistant.Expressions;
using WorkAssistant.Views;

namespace WorkAssistant.Features.ExcelProcessing
{
    public partial class ExcelProcessingPage : Page
    {
        DataTable _tableA;
        DataTable _tableB;
        ExcelProcessingCompareResult _preview;
        List<ColumnPick> _picks;
        readonly List<CondRow> _conds = new List<CondRow>();
        bool _loadingSheet;
        bool _syncPicks;

        public ExcelProcessingPage()
        {
            InitializeComponent();
            AddCond(null);
        }

        void Home_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
            else if (NavigationService != null)
                NavigationService.Navigate(new HomePage());
        }

        void BrowseA_Click(object sender, RoutedEventArgs e)
        {
            LoadFile(true);
        }

        void BrowseB_Click(object sender, RoutedEventArgs e)
        {
            LoadFile(false);
        }

        void SheetA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loadingSheet)
                return;
            ReloadSheet(true);
        }

        void SheetB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loadingSheet)
                return;
            ReloadSheet(false);
        }

        void AddCond_Click(object sender, RoutedEventArgs e)
        {
            AddCond(null);
        }

        void Op_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshPreview(true);
        }

        void Save_Click(object sender, RoutedEventArgs e)
        {
            if (_preview == null)
            {
                Alert("Choose an output operation first.", "Nothing to save", MessageBoxImage.Warning);
                return;
            }

            var selected = SelectedPicks();
            if (selected.Count == 0)
            {
                Alert("Select at least one column.", "No columns", MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = "result.xlsx"
            };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                var table = Projected(selected);
                ExcelProcessingWork.Save(table, dlg.FileName);
                Alert("Saved " + table.Rows.Count + " rows.", "Done", MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not save", MessageBoxImage.Error);
            }
        }

        void LoadFile(bool isA)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = isA ? "Open File A" : "Open File B"
            };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                ApplyLoad(isA, ExcelProcessingWork.Load(dlg.FileName, null), dlg.FileName);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not open file", MessageBoxImage.Error);
            }
        }

        void ReloadSheet(bool isA)
        {
            var path = isA ? PathA.Text : PathB.Text;
            var sheet = isA
                ? SheetA.SelectedItem as string
                : SheetB.SelectedItem as string;
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(sheet))
                return;

            try
            {
                ApplyLoad(isA, ExcelProcessingWork.Load(path, sheet), path);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not load sheet", MessageBoxImage.Error);
            }
        }

        void ApplyLoad(bool isA, ExcelProcessingLoadResult loaded, string path)
        {
            _loadingSheet = true;
            try
            {
                if (isA)
                {
                    _tableA = loaded.Table;
                    PathA.Text = path;
                    SheetA.ItemsSource = loaded.Sheets;
                    SheetA.SelectedItem = loaded.Sheet;
                    GridA.ItemsSource = loaded.Table.DefaultView;
                    RefreshCondCombos();
                }
                else
                {
                    _tableB = loaded.Table;
                    PathB.Text = path;
                    SheetB.ItemsSource = loaded.Sheets;
                    SheetB.SelectedItem = loaded.Sheet;
                    GridB.ItemsSource = loaded.Table.DefaultView;
                    RefreshCondCombos();
                }
            }
            finally
            {
                _loadingSheet = false;
            }

            RefreshPreview(false);
        }

        static string[] ColumnNames(DataTable table)
        {
            var names = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                names[i] = table.Columns[i].ColumnName;
            return names;
        }

        void RefreshPreview(bool alert)
        {
            if (_loadingSheet)
                return;

            var tag = SelectedOp();
            if (tag == null)
            {
                ClearPreview();
                return;
            }

            DataTable a;
            DataTable b;
            if (!TryReady(out a, out b, alert))
            {
                ClearPreview();
                return;
            }

            try
            {
                var extra = CollectedConds();
                if (tag == "AOnly")
                    _preview = ExcelProcessingWork.OnlyInA(a, b, extra);
                else if (tag == "BOnly")
                    _preview = ExcelProcessingWork.OnlyInB(a, b, extra);
                else
                    _preview = ExcelProcessingWork.Common(a, b, extra);

                _picks = new List<ColumnPick>();
                for (var i = 0; i < _preview.Table.Columns.Count; i++)
                {
                    _picks.Add(new ColumnPick
                    {
                        InternalName = _preview.Table.Columns[i].ColumnName,
                        OriginalName = _preview.Originals[i],
                        Source = _preview.Sources[i],
                        Include = true
                    });
                }
                ApplyPickLabels();

                ColList.ItemsSource = _picks;
                ShowProjected();
                SyncUtility();
            }
            catch (Exception ex)
            {
                ClearPreview();
                Alert(ex.Message, "Could not build preview", MessageBoxImage.Error);
            }
        }

        void ClearPreview()
        {
            _preview = null;
            _picks = null;
            ColList.ItemsSource = null;
            GridOut.ItemsSource = null;
            SyncUtility();
        }

        string SelectedOp()
        {
            var item = OpList.SelectedItem as ListBoxItem;
            return item != null ? item.Tag as string : null;
        }

        List<ColumnPick> SelectedPicks()
        {
            var list = new List<ColumnPick>();
            if (_picks == null)
                return list;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Include)
                    list.Add(_picks[i]);
            }
            return list;
        }

        DataTable Projected(List<ColumnPick> selected)
        {
            var internals = new string[selected.Count];
            var originals = new string[selected.Count];
            var sources = new char[selected.Count];
            for (var i = 0; i < selected.Count; i++)
            {
                internals[i] = selected[i].InternalName;
                originals[i] = selected[i].OriginalName;
                sources[i] = selected[i].Source;
            }

            var outputs = new string[selected.Count];
            for (var i = 0; i < selected.Count; i++)
                outputs[i] = ExcelProcessingWork.OutputName(originals[i], sources[i], originals, sources);

            return ExcelProcessingWork.Project(_preview.Table, internals, outputs);
        }

        void ApplyPickLabels()
        {
            if (_picks == null)
                return;
            var originals = new string[_picks.Count];
            var sources = new char[_picks.Count];
            for (var i = 0; i < _picks.Count; i++)
            {
                originals[i] = _picks[i].OriginalName;
                sources[i] = _picks[i].Source;
            }
            for (var i = 0; i < _picks.Count; i++)
                _picks[i].Label = ExcelProcessingWork.OutputName(originals[i], sources[i], originals, sources);
        }

        void ShowProjected()
        {
            if (_preview == null)
            {
                GridOut.ItemsSource = null;
                return;
            }
            GridOut.ItemsSource = Projected(SelectedPicks()).DefaultView;
        }

        void Col_Changed(object sender, RoutedEventArgs e)
        {
            if (_syncPicks)
                return;
            ShowProjected();
            SyncUtility();
        }

        void PickA_Click(object sender, RoutedEventArgs e)
        {
            ToggleSource('A');
        }

        void PickB_Click(object sender, RoutedEventArgs e)
        {
            ToggleSource('B');
        }

        void PickBoth_Click(object sender, RoutedEventArgs e)
        {
            if (_syncPicks)
                return;
            var on = BothState() != true;
            SetSource('A', on);
            SetSource('B', on);
            AfterUtility();
        }

        void ToggleSource(char source)
        {
            if (_syncPicks)
                return;
            SetSource(source, SourceState(source) != true);
            AfterUtility();
        }

        void SetSource(char source, bool include)
        {
            if (_picks == null)
                return;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Source == source)
                    _picks[i].Include = include;
            }
        }

        void AfterUtility()
        {
            ColList.ItemsSource = null;
            ColList.ItemsSource = _picks;
            ShowProjected();
            SyncUtility();
        }

        void SyncUtility()
        {
            _syncPicks = true;
            try
            {
                PickA.IsChecked = SourceState('A');
                PickB.IsChecked = SourceState('B');
                PickBoth.IsChecked = BothState();
                PickA.IsEnabled = HasSource('A');
                PickB.IsEnabled = HasSource('B');
                PickBoth.IsEnabled = _picks != null && _picks.Count > 0;
            }
            finally
            {
                _syncPicks = false;
            }
        }

        bool HasSource(char source)
        {
            if (_picks == null)
                return false;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Source == source)
                    return true;
            }
            return false;
        }

        bool? SourceState(char source)
        {
            if (_picks == null)
                return false;
            var any = false;
            var all = true;
            var none = true;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Source != source)
                    continue;
                any = true;
                if (_picks[i].Include)
                    none = false;
                else
                    all = false;
            }
            if (!any)
                return false;
            if (all)
                return true;
            if (none)
                return false;
            return null;
        }

        bool? BothState()
        {
            if (_picks == null || _picks.Count == 0)
                return false;
            var all = true;
            var none = true;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Include)
                    none = false;
                else
                    all = false;
            }
            if (all)
                return true;
            if (none)
                return false;
            return null;
        }

        void AddCond(ExcelProcessingMatch seed)
        {
            var root = new StackPanel { Margin = new Thickness(0, 0, 8, 8) };
            var line = new Grid();
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(56) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var panelA = new StackPanel { Margin = new Thickness(0, 0, 6, 0) };
            var opBox = new ComboBox { Margin = new Thickness(0, 0, 6, 0) };
            opBox.Items.Add("==");
            opBox.Items.Add("!=");
            opBox.Items.Add(">");
            opBox.Items.Add(">=");
            opBox.Items.Add("<");
            opBox.Items.Add("<=");
            var opKeep = seed != null && !string.IsNullOrWhiteSpace(seed.Operator) ? seed.Operator : "==";
            opBox.SelectedItem = opKeep;
            if (opBox.SelectedIndex < 0)
                opBox.SelectedIndex = 0;

            var panelB = new StackPanel { Margin = new Thickness(0, 0, 6, 0) };
            var del = new Button { Content = "Remove", Width = 72, VerticalAlignment = VerticalAlignment.Top };

            Grid.SetColumn(opBox, 1);
            Grid.SetColumn(panelB, 2);
            Grid.SetColumn(del, 3);
            line.Children.Add(panelA);
            line.Children.Add(opBox);
            line.Children.Add(panelB);
            line.Children.Add(del);
            root.Children.Add(line);

            var item = new CondRow
            {
                Root = root,
                Op = opBox,
                PanelA = panelA,
                PanelB = panelB,
                ColA = new ComboBox(),
                ColB = new ComboBox(),
                SepA = NewSepBox(),
                SepB = NewSepBox(),
                CustomA = NewCustomSep(),
                CustomB = NewCustomSep(),
                ExprA = new TextBox { Text = seed != null && seed.ExpressionA != null ? seed.ExpressionA : "{A}" },
                ExprB = new TextBox { Text = seed != null && seed.ExpressionB != null ? seed.ExpressionB : "{A}" }
            };
            item.SplitRowA = MakeSplitRow(item.SepA, item.CustomA);
            item.SplitRowB = MakeSplitRow(item.SepB, item.CustomB);
            if (seed != null)
            {
                ApplySep(item.SepA, item.CustomA, seed.SplitA);
                ApplySep(item.SepB, item.CustomB, seed.SplitB);
            }

            FillCondCombo(item.ColA, _tableA, seed != null ? seed.ColumnA : null);
            FillCondCombo(item.ColB, _tableB, seed != null ? seed.ColumnB : null);
            WireCond(item.ColA, item.ColB, item.Op, item.SepA, item.SepB, item.CustomA, item.CustomB,
                item.ExprA, item.ExprB);
            item.SepA.SelectionChanged += (s, e) => item.CustomA.IsEnabled = item.SepA.SelectedIndex == 4;
            item.SepB.SelectionChanged += (s, e) => item.CustomB.IsEnabled = item.SepB.SelectedIndex == 4;

            del.Click += (s, e) =>
            {
                _conds.Remove(item);
                CondPanel.Children.Remove(root);
                if (_conds.Count == 0)
                    AddCond(null);
                else
                    RefreshPreview(false);
            };

            FillSidePanels(item);
            _conds.Add(item);
            CondPanel.Children.Add(root);
            RefreshPreview(false);
        }

        void WireCond(params Control[] controls)
        {
            for (var i = 0; i < controls.Length; i++)
            {
                var box = controls[i] as ComboBox;
                if (box != null)
                    box.SelectionChanged += Cond_Changed;
                var text = controls[i] as TextBox;
                if (text != null)
                    text.TextChanged += Cond_TextChanged;
            }
        }

        static ComboBox NewSepBox()
        {
            var sep = new ComboBox { Width = 110, Margin = new Thickness(0, 0, 8, 2) };
            sep.Items.Add("None");
            sep.Items.Add("Space");
            sep.Items.Add("Underscore");
            sep.Items.Add("Dash");
            sep.Items.Add("Custom");
            sep.SelectedIndex = 0;
            return sep;
        }

        static TextBox NewCustomSep()
        {
            return new TextBox { Width = 40, MaxLength = 1, IsEnabled = false, Margin = new Thickness(0, 0, 0, 2) };
        }

        static WrapPanel MakeSplitRow(ComboBox sep, TextBox custom)
        {
            var row = new WrapPanel();
            row.Children.Add(new TextBlock
            {
                Text = "Split",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 2)
            });
            row.Children.Add(sep);
            row.Children.Add(custom);
            return row;
        }

        static void ApplySep(ComboBox sep, TextBox custom, string stored)
        {
            if (stored == " ")
                sep.SelectedIndex = 1;
            else if (stored == "_")
                sep.SelectedIndex = 2;
            else if (stored == "-")
                sep.SelectedIndex = 3;
            else if (!string.IsNullOrEmpty(stored))
            {
                sep.SelectedIndex = 4;
                custom.Text = stored.Substring(0, 1);
                custom.IsEnabled = true;
            }
            else
                sep.SelectedIndex = 0;
        }

        static string ReadSep(ComboBox sep, TextBox custom)
        {
            switch (sep.SelectedIndex)
            {
                case 1: return " ";
                case 2: return "_";
                case 3: return "-";
                case 4: return custom.Text.Length > 0 ? custom.Text.Substring(0, 1) : "";
                default: return "";
            }
        }

        static void FillSidePanels(CondRow item)
        {
            item.PanelA.Children.Clear();
            item.PanelB.Children.Clear();
            item.PanelA.Children.Add(Hint("Column to split"));
            item.PanelA.Children.Add(item.ColA);
            item.PanelA.Children.Add(item.SplitRowA);
            item.PanelA.Children.Add(Hint("Formula"));
            item.PanelA.Children.Add(item.ExprA);
            item.PanelB.Children.Add(Hint("Column to split"));
            item.PanelB.Children.Add(item.ColB);
            item.PanelB.Children.Add(item.SplitRowB);
            item.PanelB.Children.Add(Hint("Formula"));
            item.PanelB.Children.Add(item.ExprB);
        }

        static TextBlock Hint(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = System.Windows.Media.Brushes.Gray,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 2)
            };
        }

        void FillCondCombo(ComboBox box, DataTable table, string keep)
        {
            var names = table == null ? new string[0] : ColumnNames(table);
            box.ItemsSource = names;
            if (keep != null && table != null && table.Columns.Contains(keep))
                box.SelectedItem = keep;
            else if (names.Length > 0)
                box.SelectedIndex = 0;
            else
                box.SelectedIndex = -1;
        }

        void RefreshCondCombos()
        {
            for (var i = 0; i < _conds.Count; i++)
            {
                var c = _conds[i];
                FillCondCombo(c.ColA, _tableA, c.ColA.SelectedItem as string);
                FillCondCombo(c.ColB, _tableB, c.ColB.SelectedItem as string);
            }
        }

        void Cond_Changed(object sender, SelectionChangedEventArgs e)
        {
            RefreshPreview(false);
        }

        void Cond_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPreview(false);
        }

        List<ExcelProcessingMatch> CollectedConds()
        {
            var list = new List<ExcelProcessingMatch>();
            for (var i = 0; i < _conds.Count; i++)
            {
                var c = _conds[i];
                list.Add(new ExcelProcessingMatch
                {
                    ColumnA = c.ColA.SelectedItem as string,
                    ColumnB = c.ColB.SelectedItem as string,
                    ExpressionA = c.ExprA.Text,
                    ExpressionB = c.ExprB.Text,
                    Operator = c.Op.SelectedItem as string,
                    SplitA = ReadSep(c.SepA, c.CustomA),
                    SplitB = ReadSep(c.SepB, c.CustomB)
                });
            }
            return list;
        }

        static bool HasJoin(IList<ExcelProcessingMatch> extra)
        {
            if (extra == null)
                return false;
            for (var i = 0; i < extra.Count; i++)
            {
                var m = extra[i];
                if (!string.IsNullOrWhiteSpace(m.ExpressionA) && !string.IsNullOrWhiteSpace(m.ExpressionB))
                    return true;
            }
            return false;
        }

        bool TryReady(out DataTable a, out DataTable b, bool alert)
        {
            a = _tableA;
            b = _tableB;

            if (a == null || b == null)
            {
                if (alert)
                    Alert("Open both Excel files first.", "Missing file", MessageBoxImage.Warning);
                return false;
            }
            if (!HasJoin(CollectedConds()))
            {
                if (alert)
                    Alert("Set a match rule that uses both A and B.", "Missing match", MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            var win = new Window
            {
                Title = "Built-in variables",
                Width = 680,
                Height = 540,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = Window.GetWindow(this)
            };
            var stack = new StackPanel { Margin = new Thickness(10) };
            win.Content = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };

            stack.Children.Add(new TextBlock
            {
                Text = "Type {Today} in a formula. {A} is the full value of the 1st column on that file, {B} the 2nd. After Split, $1 $2 are parts of the chosen column. IF() works inside the formula.",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            var varGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                MaxHeight = 250,
                Margin = new Thickness(0, 0, 0, 10)
            };
            varGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new Binding("Name"), Width = 140 });
            varGrid.Columns.Add(new DataGridTextColumn { Header = "Means", Binding = new Binding("Description"), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            varGrid.Columns.Add(new DataGridTextColumn { Header = "Value now", Binding = new Binding("Example"), Width = 110 });
            varGrid.ItemsSource = ExpressionEngine.GetVariableHelp();
            stack.Children.Add(varGrid);

            var fnGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                MaxHeight = 200,
                Margin = new Thickness(0, 0, 0, 10)
            };
            fnGrid.Columns.Add(new DataGridTextColumn { Header = "Function", Binding = new Binding("Signature"), Width = 210 });
            fnGrid.Columns.Add(new DataGridTextColumn { Header = "Means", Binding = new Binding("Description"), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            fnGrid.ItemsSource = ExpressionEngine.GetFunctionHelp();
            stack.Children.Add(fnGrid);

            var close = new Button { Content = "Close", Width = 90, HorizontalAlignment = HorizontalAlignment.Right };
            close.Click += delegate { win.Close(); };
            stack.Children.Add(close);
            win.ShowDialog();
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }

        public sealed class ColumnPick
        {
            public string InternalName { get; set; }
            public string OriginalName { get; set; }
            public char Source { get; set; }
            public bool Include { get; set; }
            public string Label { get; set; }
        }

        sealed class CondRow
        {
            public StackPanel Root;
            public ComboBox Op;
            public StackPanel PanelA;
            public StackPanel PanelB;
            public ComboBox ColA;
            public ComboBox ColB;
            public ComboBox SepA;
            public ComboBox SepB;
            public TextBox CustomA;
            public TextBox CustomB;
            public WrapPanel SplitRowA;
            public WrapPanel SplitRowB;
            public TextBox ExprA;
            public TextBox ExprB;
        }
    }
}
