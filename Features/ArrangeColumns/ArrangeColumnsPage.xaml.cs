using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.ArrangeColumns
{
    public partial class ArrangeColumnsPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<ArrangeColumnsSettings> _use;
        List<Row> _rows = new List<Row>();

        public ArrangeColumnsPage(DataTable input, string label, ArrangeColumnsSettings settings, Action<ArrangeColumnsSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - " + input.Columns.Count + " columns, " + input.Rows.Count + " rows";

            var placed = new HashSet<int>();
            if (settings != null)
            {
                // The saved order first, then input columns the settings do not know yet.
                foreach (var c in settings.Columns)
                {
                    int index = ExcelColumns.IndexOf(input, c.Source);
                    if (index < 0 || !placed.Add(index))
                        continue;
                    _rows.Add(new Row { Index = index, Label = ExcelColumns.Label(input, index), Keep = c.Keep, Header = c.Header ?? "" });
                }
            }
            for (var i = 0; i < input.Columns.Count; i++)
            {
                if (placed.Contains(i))
                    continue;
                _rows.Add(new Row { Index = i, Label = ExcelColumns.Label(input, i), Keep = settings == null || settings.KeepNew, Header = "" });
            }
            KeepNewBox.IsChecked = settings == null || settings.KeepNew;
            // Typing a new name in a row selects that row too.
            ColumnList.AddHandler(UIElement.GotKeyboardFocusEvent, new RoutedEventHandler(RowFocused), true);
            ShowRows(-1);
        }

        void RowFocused(object sender, RoutedEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement(ColumnList, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (item != null && !item.IsSelected)
                item.IsSelected = true;
        }

        void ShowRows(int select)
        {
            ColumnList.ItemsSource = null;
            ColumnList.ItemsSource = _rows;
            if (select >= 0)
                ColumnList.SelectedIndex = select;
        }

        void Up_Click(object sender, RoutedEventArgs e)
        {
            Move(-1);
        }

        void Down_Click(object sender, RoutedEventArgs e)
        {
            Move(1);
        }

        void Move(int delta)
        {
            int index = ColumnList.SelectedIndex;
            if (index < 0)
            {
                Status.Text = "Select a column first: click its row, beside the box.";
                return;
            }
            int target = index + delta;
            if (target < 0 || target >= _rows.Count)
                return;
            var row = _rows[index];
            _rows[index] = _rows[target];
            _rows[target] = row;
            ShowRows(target);
        }

        void KeepAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows)
                r.Keep = true;
            ShowRows(ColumnList.SelectedIndex);
        }

        void KeepNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows)
                r.Keep = false;
            ShowRows(ColumnList.SelectedIndex);
        }

        void ClearNames_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows)
                r.Header = "";
            ShowRows(ColumnList.SelectedIndex);
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var s = new ArrangeColumnsSettings { KeepNew = KeepNewBox.IsChecked == true };
            foreach (var r in _rows)
                s.Columns.Add(new ArrangeColumn { Source = ExcelColumns.Letter(r.Index), Header = (r.Header ?? "").Trim(), Keep = r.Keep });
            DataTable result;
            try
            {
                result = ArrangeColumnsWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Arrange columns", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = "The result has " + result.Columns.Count + " columns.";
            _use(s);
        }

        public sealed class Row
        {
            public int Index { get; set; }
            public string Label { get; set; }
            public bool Keep { get; set; }
            public string Header { get; set; }
        }
    }
}
