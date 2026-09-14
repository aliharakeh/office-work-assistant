using System.Data;
using System.Windows.Controls;

namespace WorkAssistant.Excel
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
            e.Column.Header = ExcelFile.Header(view.Table.Columns[e.PropertyName]);
        }
    }
}
