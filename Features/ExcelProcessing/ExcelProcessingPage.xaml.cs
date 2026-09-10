using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
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

        void Key_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshPreview(false);
        }

        void AddCond_Click(object sender, RoutedEventArgs e)
        {
            AddCond(null, null, null);
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
                    KeyA.ItemsSource = ColumnNames(loaded.Table);
                    SelectKey(KeyA, loaded.Table);
                    GridA.ItemsSource = loaded.Table.DefaultView;
                    RefreshCondCombos();
                }
                else
                {
                    _tableB = loaded.Table;
                    PathB.Text = path;
                    SheetB.ItemsSource = loaded.Sheets;
                    SheetB.SelectedItem = loaded.Sheet;
                    KeyB.ItemsSource = ColumnNames(loaded.Table);
                    SelectKey(KeyB, loaded.Table);
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

        static void SelectKey(ComboBox box, DataTable table)
        {
            var keep = box.SelectedItem as string;
            if (keep != null && table.Columns.Contains(keep))
                box.SelectedItem = keep;
            else if (box.Items.Count > 0)
                box.SelectedIndex = 0;
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
            string keyA;
            string keyB;
            if (!TryReady(out a, out b, out keyA, out keyB, alert))
            {
                ClearPreview();
                return;
            }

            try
            {
                var extra = CollectedConds();
                if (tag == "AOnly")
                    _preview = ExcelProcessingWork.OnlyInA(a, b, keyA, keyB, extra);
                else if (tag == "BOnly")
                    _preview = ExcelProcessingWork.OnlyInB(a, b, keyA, keyB, extra);
                else
                    _preview = ExcelProcessingWork.Common(a, b, keyA, keyB, extra);

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

        void AddCond(string colA, string colB, string value)
        {
            var row = new Grid { Margin = new Thickness(0, 0, 8, 4) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var a = new ComboBox { Margin = new Thickness(0, 0, 6, 0) };
            var eq = new TextBlock
            {
                Text = "=",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            };
            var b = new ComboBox { Margin = new Thickness(0, 0, 6, 0) };
            var val = new TextBox { Margin = new Thickness(0, 0, 6, 0), Text = value ?? "" };
            var del = new Button { Content = "Remove", Width = 72 };

            Grid.SetColumn(eq, 1);
            Grid.SetColumn(b, 2);
            Grid.SetColumn(val, 3);
            Grid.SetColumn(del, 4);
            row.Children.Add(a);
            row.Children.Add(eq);
            row.Children.Add(b);
            row.Children.Add(val);
            row.Children.Add(del);

            var item = new CondRow { Root = row, ColA = a, ColB = b, Value = val };
            FillCondCombo(a, _tableA, colA);
            FillCondCombo(b, _tableB, colB);
            a.SelectionChanged += Cond_Changed;
            b.SelectionChanged += Cond_Changed;
            val.TextChanged += Cond_TextChanged;
            del.Click += (s, e) =>
            {
                _conds.Remove(item);
                CondPanel.Children.Remove(row);
                RefreshPreview(false);
            };

            _conds.Add(item);
            CondPanel.Children.Add(row);
            RefreshPreview(false);
        }

        void FillCondCombo(ComboBox box, DataTable table, string keep)
        {
            var names = table == null ? new string[0] : ColumnNames(table);
            box.ItemsSource = names;
            if (keep != null && table != null && table.Columns.Contains(keep))
                box.SelectedItem = keep;
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
                    Value = c.Value.Text
                });
            }
            return list;
        }

        bool TryReady(out DataTable a, out DataTable b, out string keyA, out string keyB, bool alert)
        {
            a = _tableA;
            b = _tableB;
            keyA = KeyA.SelectedItem as string;
            keyB = KeyB.SelectedItem as string;

            if (a == null || b == null)
            {
                if (alert)
                    Alert("Open both Excel files first.", "Missing file", MessageBoxImage.Warning);
                return false;
            }
            if (string.IsNullOrEmpty(keyA) || string.IsNullOrEmpty(keyB))
            {
                if (alert)
                    Alert("Choose a key column on each file.", "Missing key", MessageBoxImage.Warning);
                return false;
            }
            return true;
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
            public Grid Root;
            public ComboBox ColA;
            public ComboBox ColB;
            public TextBox Value;
        }
    }
}
