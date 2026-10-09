using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FilterSort
{
    public partial class FilterSortPage : UserControl
    {
        DataTable _source;
        int _totalRows;
        int _previewRows = 1000;
        List<FilterCondition> _conditions = new List<FilterCondition>();
        List<SortKey> _sorts = new List<SortKey>();
        List<ColumnPick> _picks;
        int _editingCond = -1;
        int _editingSort = -1;
        bool _syncPicks;
        // The input comes from an earlier pipeline step, and "Use in pipeline"
        // hands the settings back.
        DataTable _pipeInput;
        string _pipeLabel;
        Action<FilterSortSettings> _use;

        public FilterSortPage(DataTable input, string label, FilterSortSettings settings, Action<FilterSortSettings> use)
        {
            // Combo boxes fire their handlers while InitializeComponent runs,
            // before the other controls exist, so block them until loaded.
            _syncPicks = true;
            try
            {
                InitializeComponent();
            }
            finally
            {
                _syncPicks = false;
            }
            _pipeInput = input;
            _pipeLabel = label;
            _use = use;
            PreviewBox.Items.Add(200);
            PreviewBox.Items.Add(1000);
            PreviewBox.Items.Add(5000);
            PreviewBox.SelectedItem = 1000;
            MatchMode.SelectedIndex = 0;
            FormulaField.Watch(CondFormula, FormulaRequired);
            CondKind.SelectedIndex = 0;
            CondOp.SelectedIndex = 0;
            SortDirection.SelectedIndex = 0;
            RefreshConditionList();
            RefreshSortList();
            ExcelGrid.Hook(GridSource, GridOut);
            InjectInput();
            if (settings != null)
                ApplySettings(settings);
        }

        // The grid shows a capped preview of the pipeline input.
        void InjectInput()
        {
            bool cut = _pipeInput.Rows.Count > _previewRows;
            _source = Capped(_pipeInput, _previewRows);
            _totalRows = _pipeInput.Rows.Count;
            SourcePath.Text = _pipeLabel;
            GridSource.ItemsSource = _source.DefaultView;
            InfoSource.Text = RowInfo(_source.Rows.Count, _totalRows, cut);
            RebuildPicks();
            RefreshColumnCombos();
        }

        void PreviewBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewBox.SelectedItem == null)
                return;
            var rows = (int)PreviewBox.SelectedItem;
            if (rows == _previewRows)
                return;
            _previewRows = rows;
            InjectInput();
        }

        static string RowInfo(int shown, int total, bool cut)
        {
            var rows = shown.ToString("N0", CultureInfo.InvariantCulture);
            if (!cut)
                return rows + " rows";
            return "Showing " + rows + " of " + total.ToString("N0", CultureInfo.InvariantCulture) +
                   " rows (all rows are used when you preview or save)";
        }

        // Building the result needs every row, not just the capped preview.
        DataTable FullSource()
        {
            return _pipeInput;
        }

        // Keep the output grid light; Save writes the full table.
        static DataTable Capped(DataTable table, int maxRows)
        {
            if (maxRows <= 0 || table.Rows.Count <= maxRows)
                return table;
            var copy = table.Clone();
            for (var i = 0; i < maxRows; i++)
                copy.ImportRow(table.Rows[i]);
            return copy;
        }

        void RefreshColumnCombos()
        {
            string[] labels = FilterSortWork.SourceLabels(_source);
            CondColumn.ItemsSource = labels;
            SortColumn.ItemsSource = labels;
            if (CondColumn.Items.Count > 0 && CondColumn.SelectedIndex < 0)
                CondColumn.SelectedIndex = 0;
            if (SortColumn.Items.Count > 0 && SortColumn.SelectedIndex < 0)
                SortColumn.SelectedIndex = 0;
        }

        // ---------- conditions ----------

        void UniqueValues_Click(object sender, RoutedEventArgs e)
        {
            if (_source == null)
            {
                Alert("Choose a data file first.", "Unique values", MessageBoxImage.Warning);
                return;
            }
            int col = CondColumn.SelectedIndex;
            if (col < 0)
            {
                Alert("Choose a column first.", "Unique values", MessageBoxImage.Warning);
                return;
            }

            string[] values;
            try
            {
                values = ExcelFile.UniqueValues(FullSource(), col);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Unique values", MessageBoxImage.Error);
                return;
            }
            if (values.Length == 0)
            {
                Alert("That column has no values to pick.", "Unique values", MessageBoxImage.Information);
                return;
            }

            string picked = PickUnique(values);
            if (picked != null)
                CondValue.Text = picked;
        }

        string PickUnique(string[] values)
        {
            var list = new ListBox { ItemsSource = values };
            if (values.Length > 0)
                list.SelectedIndex = 0;
            var win = new Window
            {
                Title = "Unique values",
                Width = 360,
                Height = 420,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = Window.GetWindow(this)
            };
            string picked = null;
            var use = new Button { Content = "Use", Width = 80, IsDefault = true };
            use.Click += delegate
            {
                picked = list.SelectedItem as string;
                if (picked == null)
                    return;
                win.DialogResult = true;
            };
            list.MouseDoubleClick += delegate
            {
                picked = list.SelectedItem as string;
                if (picked != null)
                    win.DialogResult = true;
            };
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 8, 0, 0) };
            buttons.Children.Add(use);
            buttons.Children.Add(cancel);
            var root = new DockPanel { Margin = new Thickness(10) };
            DockPanel.SetDock(buttons, Dock.Bottom);
            root.Children.Add(buttons);
            root.Children.Add(list);
            win.Content = root;
            win.ShowDialog();
            return picked;
        }

        void CondKind_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (PanelCondColumn == null || PanelCondFormula == null)
                return;
            bool formula = CondKind.SelectedIndex == 1;
            PanelCondColumn.Visibility = formula ? Visibility.Collapsed : Visibility.Visible;
            PanelCondFormula.Visibility = formula ? Visibility.Visible : Visibility.Collapsed;
            FormulaField.Refresh(CondFormula);
        }

        bool FormulaRequired()
        {
            return CondKind.SelectedIndex == 1;
        }

        void Conds_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = CondList.SelectedIndex;
            if (index < 0 || index >= _conditions.Count)
                return;
            _editingCond = index;
            LoadCondEditor(_conditions[index]);
            AddCondBtn.Content = "Update condition";
        }

        void AddCond_Click(object sender, RoutedEventArgs e)
        {
            if (CondKind.SelectedIndex == 1 && !FormulaField.Check(CondFormula))
                return;

            FilterCondition cond;
            try
            {
                cond = BuildCondFromEditor();
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Condition", MessageBoxImage.Warning);
                return;
            }

            if (_editingCond >= 0 && _editingCond < _conditions.Count)
                _conditions[_editingCond] = cond;
            else
                _conditions.Add(cond);

            ClearCondEditor();
            RefreshConditionList();
        }

        void RemoveCond_Click(object sender, RoutedEventArgs e)
        {
            int index = CondList.SelectedIndex;
            if (index < 0 || index >= _conditions.Count)
            {
                Alert("Select a condition first.", "Remove", MessageBoxImage.Warning);
                return;
            }
            _conditions.RemoveAt(index);
            ClearCondEditor();
            RefreshConditionList();
        }

        void ClearCond_Click(object sender, RoutedEventArgs e)
        {
            ClearCondEditor();
            RefreshConditionList();
        }

        FilterCondition BuildCondFromEditor()
        {
            var cond = new FilterCondition();
            if (CondKind.SelectedIndex == 1)
            {
                cond.Kind = FilterConditionKind.Formula;
                cond.Expression = CondFormula.Text;
                string err = ExpressionEngine.Validate(cond.Expression);
                if (err != null)
                    throw new InvalidOperationException(err);
                return cond;
            }

            cond.Kind = FilterConditionKind.Column;
            string picked = CondColumn.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(picked))
                throw new InvalidOperationException("Choose a column first.");
            cond.Column = FilterSortWork.ExtractLetter(picked);
            int op = CondOp.SelectedIndex;
            cond.Operator = op >= 0 && op < FilterSortWork.Operators.Length ? FilterSortWork.Operators[op] : "==";
            cond.Value = CondValue.Text != null ? CondValue.Text : "";
            return cond;
        }

        void LoadCondEditor(FilterCondition cond)
        {
            if (cond == null)
                return;
            if (cond.Kind == FilterConditionKind.Formula)
            {
                CondKind.SelectedIndex = 1;
                CondFormula.Text = cond.Expression;
            }
            else
            {
                CondKind.SelectedIndex = 0;
                SelectColumn(CondColumn, cond.Column);
                int op = Array.IndexOf(FilterSortWork.Operators, cond.Operator);
                CondOp.SelectedIndex = op >= 0 ? op : 0;
                CondValue.Text = cond.Value;
            }
        }

        void ClearCondEditor()
        {
            _editingCond = -1;
            CondList.SelectedIndex = -1;
            AddCondBtn.Content = "Add condition";
            CondFormula.Text = "";
            CondValue.Text = "";
            CondOp.SelectedIndex = 0;
            CondKind.SelectedIndex = 0;
            if (CondColumn.Items.Count > 0)
                CondColumn.SelectedIndex = 0;
        }

        void RefreshConditionList()
        {
            CondList.ItemsSource = null;
            CondList.ItemsSource = _conditions;
        }

        // ---------- sorting ----------

        void Sorts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = SortList.SelectedIndex;
            if (index < 0 || index >= _sorts.Count)
                return;
            _editingSort = index;
            LoadSortEditor(_sorts[index]);
            AddSortBtn.Content = "Update sort";
        }

        void AddSort_Click(object sender, RoutedEventArgs e)
        {
            SortKey key;
            try
            {
                key = BuildSortFromEditor();
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Sort", MessageBoxImage.Warning);
                return;
            }

            if (_editingSort >= 0 && _editingSort < _sorts.Count)
                _sorts[_editingSort] = key;
            else
                _sorts.Add(key);

            ClearSortEditor();
            RefreshSortList();
        }

        void RemoveSort_Click(object sender, RoutedEventArgs e)
        {
            int index = SortList.SelectedIndex;
            if (index < 0 || index >= _sorts.Count)
            {
                Alert("Select a sort key first.", "Remove", MessageBoxImage.Warning);
                return;
            }
            _sorts.RemoveAt(index);
            ClearSortEditor();
            RefreshSortList();
        }

        void SortUp_Click(object sender, RoutedEventArgs e)
        {
            MoveSort(-1);
        }

        void SortDown_Click(object sender, RoutedEventArgs e)
        {
            MoveSort(1);
        }

        // The list order is the sort priority: first key wins.
        void MoveSort(int delta)
        {
            int index = SortList.SelectedIndex;
            if (index < 0 || index >= _sorts.Count)
            {
                Alert("Select a sort key first.", "Sort", MessageBoxImage.Warning);
                return;
            }
            int target = index + delta;
            if (target < 0 || target >= _sorts.Count)
                return;
            SortKey item = _sorts[index];
            _sorts[index] = _sorts[target];
            _sorts[target] = item;
            RefreshSortList();
            SortList.SelectedIndex = target;
        }

        SortKey BuildSortFromEditor()
        {
            string picked = SortColumn.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(picked))
                throw new InvalidOperationException("Choose a column to sort by.");
            var key = new SortKey();
            key.Column = FilterSortWork.ExtractLetter(picked);
            key.Descending = SortDirection.SelectedIndex == 1;
            return key;
        }

        void LoadSortEditor(SortKey key)
        {
            if (key == null)
                return;
            SelectColumn(SortColumn, key.Column);
            SortDirection.SelectedIndex = key.Descending ? 1 : 0;
        }

        void ClearSortEditor()
        {
            _editingSort = -1;
            SortList.SelectedIndex = -1;
            AddSortBtn.Content = "Add sort";
            if (SortColumn.Items.Count > 0)
                SortColumn.SelectedIndex = 0;
            SortDirection.SelectedIndex = 0;
        }

        void RefreshSortList()
        {
            SortList.ItemsSource = null;
            SortList.ItemsSource = _sorts;
        }

        static void SelectColumn(ComboBox combo, string column)
        {
            string want = FilterSortWork.ExtractLetter(column);
            for (var i = 0; i < combo.Items.Count; i++)
            {
                string item = combo.Items[i] as string;
                if (string.Equals(FilterSortWork.ExtractLetter(item), want, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedIndex = i;
                    return;
                }
            }
            combo.SelectedIndex = -1;
        }

        // ---------- columns to keep ----------

        void RebuildPicks()
        {
            _syncPicks = true;
            try
            {
                _picks = new List<ColumnPick>();
                if (_source != null)
                {
                    for (var c = 0; c < _source.Columns.Count; c++)
                    {
                        _picks.Add(new ColumnPick
                        {
                            Name = _source.Columns[c].ColumnName,
                            Label = FilterSortWork.SourceLabel(_source, c),
                            Include = true
                        });
                    }
                }
                KeepColumnList.ItemsSource = null;
                KeepColumnList.ItemsSource = _picks;
            }
            finally
            {
                _syncPicks = false;
            }
            SyncSelectAll();
        }

        void KeepColumn_Changed(object sender, RoutedEventArgs e)
        {
            if (_syncPicks)
                return;
            SyncSelectAll();
        }

        void SelectAllKeep_Click(object sender, RoutedEventArgs e)
        {
            if (_syncPicks || _picks == null)
                return;
            bool on = SelectAllKeep.IsChecked != false;
            _syncPicks = true;
            try
            {
                for (var i = 0; i < _picks.Count; i++)
                    _picks[i].Include = on;
                KeepColumnList.ItemsSource = null;
                KeepColumnList.ItemsSource = _picks;
            }
            finally
            {
                _syncPicks = false;
            }
            SyncSelectAll();
        }

        void SyncSelectAll()
        {
            if (SelectAllKeep == null || KeepColumnList == null)
                return;
            _syncPicks = true;
            try
            {
                bool enabled = _picks != null && _picks.Count > 0;
                SelectAllKeep.IsEnabled = enabled;
                if (!enabled)
                {
                    SelectAllKeep.IsChecked = false;
                    return;
                }
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
                    SelectAllKeep.IsChecked = true;
                else if (none)
                    SelectAllKeep.IsChecked = false;
                else
                    SelectAllKeep.IsChecked = null;
            }
            finally
            {
                _syncPicks = false;
            }
        }

        List<string> SelectedKeepColumns()
        {
            var list = new List<string>();
            if (_picks == null || _source == null)
                return list;
            // Store Excel letters (A, B, ...) so the result follows position, not header names.
            for (var i = 0; i < _picks.Count; i++)
            {
                if (!_picks[i].Include)
                    continue;
                int at = _source.Columns.IndexOf(_picks[i].Name);
                if (at < 0)
                    at = i;
                list.Add(FilterSortWork.ColumnLetter(at));
            }
            return list;
        }

        // ---------- result ----------

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "Type the Name exactly, e.g. $Today. Works in column values and formulas. Week starts Monday. Dates have no time except $Now. Use $A $B for columns (A = 1st, B = 2nd).");
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            DataTable output;
            if (!TryBuildOutput(out output))
                return;
            GridOut.ItemsSource = Capped(output, _previewRows).DefaultView;
            InfoOut.Text = output.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " of " +
                _totalRows.ToString("N0", CultureInfo.InvariantCulture) + " rows";
        }

        void UseInPipeline_Click(object sender, RoutedEventArgs e)
        {
            DataTable output;
            if (!TryBuildOutput(out output))
                return;
            _use(CurrentSettings());
        }

        FilterSortSettings CurrentSettings()
        {
            var s = new FilterSortSettings();
            s.KeepAll = SelectAllKeep.IsChecked == true;
            s.Keep = s.KeepAll ? new List<string>() : SelectedKeepColumns();
            s.Conditions = new List<FilterCondition>(_conditions);
            s.MatchAll = MatchMode.SelectedIndex != 1;
            s.Sorts = new List<SortKey>(_sorts);
            return s;
        }

        void ApplySettings(FilterSortSettings s)
        {
            _conditions = s.Conditions != null ? new List<FilterCondition>(s.Conditions) : new List<FilterCondition>();
            _sorts = s.Sorts != null ? new List<SortKey>(s.Sorts) : new List<SortKey>();
            MatchMode.SelectedIndex = s.MatchAll ? 0 : 1;
            if (!s.KeepAll && _picks != null)
            {
                var keep = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string column in s.Keep ?? new List<string>())
                    keep.Add(FilterSortWork.ExtractLetter(column));
                _syncPicks = true;
                try
                {
                    for (var i = 0; i < _picks.Count; i++)
                        _picks[i].Include = keep.Contains(FilterSortWork.ColumnLetter(i));
                    KeepColumnList.ItemsSource = null;
                    KeepColumnList.ItemsSource = _picks;
                }
                finally
                {
                    _syncPicks = false;
                }
                SyncSelectAll();
            }
            RefreshConditionList();
            RefreshSortList();
        }

        bool TryBuildOutput(out DataTable output)
        {
            output = null;
            if (_source == null)
            {
                Alert("Choose a data file first.", "Missing file", MessageBoxImage.Warning);
                return false;
            }
            if (_source.Columns.Count == 0)
            {
                Alert("The sheet has no data.", "Missing data", MessageBoxImage.Warning);
                return false;
            }
            string error = FilterSortWork.Validate(_conditions);
            if (error != null)
            {
                Alert(error, "Filter", MessageBoxImage.Warning);
                return false;
            }
            List<string> keep = SelectedKeepColumns();
            if (keep.Count == 0)
            {
                Alert("Keep at least one column.", "Columns", MessageBoxImage.Warning);
                return false;
            }
            try
            {
                output = FilterSortWork.Apply(FullSource(), keep, _conditions, MatchMode.SelectedIndex != 1, _sorts);
                return true;
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not build result", MessageBoxImage.Error);
                return false;
            }
        }

        public sealed class ColumnPick
        {
            public string Name { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }
    }
}
