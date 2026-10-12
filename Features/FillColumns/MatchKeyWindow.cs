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

        MatchKeyWindow(Window owner, FillColumnsKey seed)
        {
            var s = seed ?? new FillColumnsKey();

            Title = seed == null ? "Add match key" : "Edit match key";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Height;
            Width = 560;

            _exprA.Text = s.FormulaA ?? "$A";
            _exprA.ToolTip = "Formula on the A row, e.g. $ID or TRIM($Code).";
            _exprB.Text = s.FormulaB ?? "$A";
            _exprB.ToolTip = "Formula on the B row, e.g. $ID or TRIM($Code).";
            FormulaField.Watch(_exprA);
            FormulaField.Watch(_exprB);

            var grid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var boxA = new GroupBox { Header = "File A", Padding = new Thickness(8), Content = _exprA };
            var eq = new TextBlock { Text = "==", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 8, 0) };
            var boxB = new GroupBox { Header = "File B", Padding = new Thickness(8), Content = _exprB };
            Grid.SetColumn(eq, 1);
            Grid.SetColumn(boxB, 2);
            grid.Children.Add(boxA);
            grid.Children.Add(eq);
            grid.Children.Add(boxB);

            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            var hint = new TextBlock
            {
                Text = "Each side sees only its own file. Text matches ignore case. The first A row with a key is used.",
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
            Content = root;
        }

        // Returns the key the user saved, or null when the window was cancelled.
        public static FillColumnsKey Edit(Window owner, FillColumnsKey seed)
        {
            var win = new MatchKeyWindow(owner, seed);
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
