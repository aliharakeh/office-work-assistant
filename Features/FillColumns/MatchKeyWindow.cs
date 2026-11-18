using System;
using System.Data;
using System.Collections.Generic;
using OfficeWorkAssistant.Excel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FillColumns
{
    // Creates or edits one match key of the Fill columns step. The step lists its keys; this window holds the form.
    sealed class MatchKeyWindow : Window
    {
        readonly TextBox _exprA = new TextBox();
        readonly TextBox _exprB = new TextBox();
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        FillColumnsKey _result;

        MatchKeyWindow(Window owner, FillColumnsKey seed, int source, IList<DataTable> sources, DataTable target)
        {
            var s = seed ?? new FillColumnsKey();

            Title = seed == null ? "Add match key" : "Edit match key";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Height = 640;
            MinHeight = 300;
            MinWidth = 280;
            Width = 560;

            _exprA.Text = s.FormulaA ?? "$" + (source + 2) + "_A";
            _exprA.ToolTip = "Formula on the lookup row, e.g. $ID or TRIM($Code).";
            _exprB.Text = s.FormulaB ?? "$1_A";
            _exprB.ToolTip = "Formula on the row being filled, e.g. $ID or TRIM($Code).";
            FormulaField.Watch(_exprA);
            FormulaField.Watch(_exprB);

            var grid = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };

            var boxA = new GroupBox { Header = "Value from input " + (source + 2), Padding = new Thickness(8), Content = _exprA };
            var eq = new TextBlock { Text = "must equal", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 8, 0) };
            var boxB = new GroupBox { Header = "Value from input 1 or an earlier match", Padding = new Thickness(8), Content = _exprB };

            grid.Children.Add(boxA);
            grid.Children.Add(ColumnGuide(source + 2, sources[source]));
            grid.Children.Add(eq);
            grid.Children.Add(boxB);
            grid.Children.Add(ColumnGuide(1, target));
            for (int i = 0; i < source; i++) grid.Children.Add(ColumnGuide(i + 2, sources[i]));

            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            var hint = new TextBlock
            {
                Text = "Use the input number before the column, for example $1_BuyerID for BuyerID on input 1. Empty keys are skipped. Text matches ignore case and surrounding spaces.",
                FontSize = 11,
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            };

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(grid);
            root.Children.Add(hint);
            root.Children.Add(_error);
            root.Children.Add(buttons);
            Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/OfficeWorkAssistant;component/Excel/EditorStyles.xaml", UriKind.Relative) });
            Content = new ScrollViewer { Content = root, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        }

        static TextBlock ColumnGuide(int input, DataTable table)
        {
            return new TextBlock { Text = "Input " + input + " ($" + input + "_A for column A): " + string.Join(", ", ExcelColumns.Labels(table)),
                TextWrapping = TextWrapping.Wrap };
        }

        // Returns the key the user saved, or null when the window was cancelled.
        public static FillColumnsKey Edit(Window owner, FillColumnsKey seed, int source, IList<DataTable> sources, DataTable target)
        {
            var win = new MatchKeyWindow(owner, seed, source, sources, target);
            win.Title += " — input " + (source + 2);
            return win.ShowDialog() == true ? win._result : null;
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!FormulaField.Ok(_exprA) || !FormulaField.Ok(_exprB))
            {
                _error.Text = "A formula is not valid. Fix the one shown in red.";
                _error.Visibility = Visibility.Visible;
                return;
            }
            _result = new FillColumnsKey { FormulaA = _exprA.Text, FormulaB = _exprB.Text };
            DialogResult = true;
        }
    }
}
