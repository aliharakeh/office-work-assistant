using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace OfficeWorkAssistant.Expressions
{
    public static class ExpressionHelp
    {
        public static void Show(Window owner, string intro)
        {
            Show(owner, intro, null);
        }

        public static void Show(Window owner, string intro, IList<VariableHelp> extras)
        {
            var rows = new List<VariableHelp>();
            if (extras != null)
            {
                foreach (var row in extras)
                    rows.Add(row);
            }
            rows.AddRange(ExpressionEngine.GetVariableHelp());

            var win = new Window
            {
                Title = "Built-in variables",
                Width = 680,
                Height = 540,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = owner
            };
            var stack = new StackPanel { Margin = new Thickness(10) };
            win.Content = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = stack
            };

            if (!string.IsNullOrWhiteSpace(intro))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = intro,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 8)
                });
            }

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
            varGrid.ItemsSource = rows;
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
    }
}
