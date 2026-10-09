using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace OfficeWorkAssistant.Excel
{
    public static class ExcelGrid
    {
        public static void Hook(params DataGrid[] grids)
        {
            if (grids == null)
                return;
            for (var i = 0; i < grids.Length; i++)
            {
                if (grids[i] != null)
                    grids[i].AutoGeneratingColumn += OnAutoGeneratingColumn;
            }
        }

        static void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            var grid = sender as DataGrid;
            var view = grid != null ? grid.ItemsSource as DataView : null;
            if (view == null || view.Table == null || !view.Table.Columns.Contains(e.PropertyName))
                return;
            var col = view.Table.Columns[e.PropertyName];
            e.Column.Header = ExcelFile.Header(col);

            // A table from Highlight / Format sheet shows its colours, as Save will write them.
            var style = ExcelFile.GetStyle(view.Table);
            if (style != null)
                e.Column.CellStyle = CellStyleFor(StyledTable.For(view.Table, style), col.Ordinal);
        }

        static Style CellStyleFor(StyledTable table, int col)
        {
            var st = new Style(typeof(DataGridCell));
            st.Setters.Add(new Setter(Control.BackgroundProperty, new Binding { Converter = new CellPart(table, col, 0) }));
            st.Setters.Add(new Setter(Control.ForegroundProperty, new Binding { Converter = new CellPart(table, col, 1) }));
            st.Setters.Add(new Setter(Control.FontWeightProperty, new Binding { Converter = new CellPart(table, col, 2) }));
            st.Setters.Add(new Setter(Control.FontStyleProperty, new Binding { Converter = new CellPart(table, col, 3) }));
            return st;
        }

        // Row positions of one shown table, so a cell can find its style even when the grid is sorted.
        sealed class StyledTable
        {
            static readonly ConditionalWeakTable<DataTable, StyledTable> Known = new ConditionalWeakTable<DataTable, StyledTable>();

            public readonly SheetStyle Style;
            readonly DataTable _table;
            Dictionary<DataRow, int> _rows;

            StyledTable(DataTable table, SheetStyle style)
            {
                _table = table;
                Style = style;
            }

            public static StyledTable For(DataTable table, SheetStyle style)
            {
                StyledTable known;
                if (Known.TryGetValue(table, out known) && known.Style == style)
                    return known;
                Known.Remove(table);
                known = new StyledTable(table, style);
                Known.Add(table, known);
                return known;
            }

            public int RowOf(DataRow row)
            {
                if (_rows == null || _rows.Count != _table.Rows.Count)
                {
                    _rows = new Dictionary<DataRow, int>(_table.Rows.Count);
                    for (var i = 0; i < _table.Rows.Count; i++)
                        _rows[_table.Rows[i]] = i;
                }
                int index;
                return _rows.TryGetValue(row, out index) ? index : -1;
            }
        }

        // One part of a cell's look (0 fill, 1 text colour, 2 weight, 3 italic) from its row.
        sealed class CellPart : IValueConverter
        {
            static readonly Dictionary<string, Brush> Brushes = new Dictionary<string, Brush>(StringComparer.OrdinalIgnoreCase);

            readonly StyledTable _table;
            readonly int _col;
            readonly int _part;

            public CellPart(StyledTable table, int col, int part)
            {
                _table = table;
                _col = col;
                _part = part;
            }

            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                var rowView = value as DataRowView;
                if (rowView == null)
                    return DependencyProperty.UnsetValue;
                int row = _table.RowOf(rowView.Row);
                var s = row < 0 ? null : _table.Style.Effective(row, _col);
                if (s == null)
                    return DependencyProperty.UnsetValue;
                if (_part == 0)
                    return BrushOf(s.Fill);
                if (_part == 1)
                    return BrushOf(s.FontColor);
                if (_part == 2)
                    return s.Bold ? (object)FontWeights.Bold : DependencyProperty.UnsetValue;
                return s.Italic ? (object)FontStyles.Italic : DependencyProperty.UnsetValue;
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                throw new NotSupportedException();
            }

            static object BrushOf(string color)
            {
                byte r;
                byte g;
                byte b;
                if (!NamedColors.TryRgb(color, out r, out g, out b))
                    return DependencyProperty.UnsetValue;
                Brush brush;
                if (!Brushes.TryGetValue(color, out brush))
                {
                    brush = new SolidColorBrush(Color.FromRgb(r, g, b));
                    brush.Freeze();
                    Brushes[color] = brush;
                }
                return brush;
            }
        }
    }
}
