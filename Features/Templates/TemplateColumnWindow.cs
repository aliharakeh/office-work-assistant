using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.Templates
{
    // Creates or edits one added column of the Add columns step. The step lists its columns; this window holds the form.
    sealed class TemplateColumnWindow : Window
    {
        readonly DataTable _input;
        readonly IList<TemplateColumn> _existing;
        readonly int _editingIndex;
        readonly TextBox _name = new TextBox { Margin = new Thickness(0, 0, 12, 0) };
        readonly ComboBox _kind = new ComboBox();
        readonly TextBox _fixedValue = new TextBox();
        readonly ComboBox _copySource = new ComboBox();
        readonly TextBox _pattern = new TextBox();
        readonly TextBox _math = new TextBox();
        readonly TextBox _condition = new TextBox();
        readonly TextBox _trueValue = new TextBox();
        readonly TextBox _falseValue = new TextBox();
        readonly StackPanel _panelFixed = new StackPanel { Margin = new Thickness(0, 0, 0, 6) };
        readonly StackPanel _panelCopy = new StackPanel { Margin = new Thickness(0, 0, 0, 6), Visibility = Visibility.Collapsed };
        readonly StackPanel _panelConcat = new StackPanel { Margin = new Thickness(0, 0, 0, 6), Visibility = Visibility.Collapsed };
        readonly StackPanel _panelMath = new StackPanel { Margin = new Thickness(0, 0, 0, 6), Visibility = Visibility.Collapsed };
        readonly StackPanel _panelConditional = new StackPanel { Margin = new Thickness(0, 0, 0, 6), Visibility = Visibility.Collapsed };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        TemplateColumn _result;

        TemplateColumnWindow(Window owner, DataTable input, TemplateColumn seed, IList<TemplateColumn> existing, int editingIndex)
        {
            _input = input;
            _existing = existing;
            _editingIndex = editingIndex;
            var s = seed ?? new TemplateColumn();

            Title = seed == null ? "Add column" : "Edit column";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Height;
            Width = 640;

            _name.Text = s.Name != null ? s.Name : "";
            _kind.Items.Add("Fixed value");
            _kind.Items.Add("Copy column");
            _kind.Items.Add("Combine text");
            _kind.Items.Add("Math");
            _kind.Items.Add("Conditional");
            _kind.SelectionChanged += (sender, e) => ShowKind();

            _copySource.ItemsSource = TemplatesWork.SourceLabels(input);
            FormulaField.Watch(_math, MathRequired);
            FormulaField.Watch(_condition, CondRequired);
            FormulaField.WatchOptional(_trueValue);
            FormulaField.WatchOptional(_falseValue);

            _panelFixed.Children.Add(Hint("Plain text, or a built-in like $Today. See Variables below."));
            _panelFixed.Children.Add(Label("Value"));
            _panelFixed.Children.Add(_fixedValue);

            _panelCopy.Children.Add(Label("From column"));
            _panelCopy.Children.Add(_copySource);

            _panelConcat.Children.Add(Label("Pattern"));
            _panelConcat.Children.Add(_pattern);
            _panelConcat.Children.Add(Hint("Use $A $B for columns (A = 1st, B = 2nd, ...). E.g. $A $B. Built-ins like $Today work too."));

            _panelMath.Children.Add(Label("Formula"));
            _panelMath.Children.Add(_math);
            _panelMath.Children.Add(Hint("Mix $A $B columns, $Today-style variables, numbers, text in quotes and + - * / ^ with ( ). E.g. $A * $B, $Today + 1, ADDDAYS($Today, 7), FORMAT($Today, \"dd/MM/yyyy\"). Comparisons and IF also work, e.g. IF($A > 10, \"Big\", \"Small\"). FIRSTWORD() and LASTWORD() pick words from a text."));

            _panelConditional.Children.Add(Label("Condition"));
            _panelConditional.Children.Add(_condition);
            _panelConditional.Children.Add(Label("Then value"));
            _panelConditional.Children.Add(_trueValue);
            _panelConditional.Children.Add(Label("Else value"));
            _panelConditional.Children.Add(_falseValue);
            _panelConditional.Children.Add(Hint("Condition uses > >= == < <= != with && || !, e.g. $A > 10 && $B == \"OK\" or $C > $Today. Then/Else are formulas like $B * 2 or text in quotes."));

            var nameRow = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            nameRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            nameRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            nameRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });
            nameRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var nameLabel = new TextBlock { Text = "Col name", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(_name, 1);
            var kindLabel = new TextBlock { Text = "Kind", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(kindLabel, 2);
            Grid.SetColumn(_kind, 3);
            nameRow.Children.Add(nameLabel);
            nameRow.Children.Add(_name);
            nameRow.Children.Add(kindLabel);
            nameRow.Children.Add(_kind);

            var variables = new Button { Content = "Built-in variables...", Width = 140, HorizontalAlignment = HorizontalAlignment.Left };
            variables.Click += (sender, e) => ExpressionHelp.Show(this,
                "Type the Name exactly, e.g. $Today. Works in Fixed, Combine, Math and Conditional. Week starts Monday. Dates have no time except $Now. Use $A $B for columns (A = 1st, B = 2nd). FIRSTWORD() and LASTWORD() pick the first/last word of a text.");
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(nameRow);
            root.Children.Add(_panelFixed);
            root.Children.Add(_panelCopy);
            root.Children.Add(_panelConcat);
            root.Children.Add(_panelMath);
            root.Children.Add(_panelConditional);
            root.Children.Add(new DockPanel { Margin = new Thickness(0, 0, 0, 6), LastChildFill = false });
            root.Children.Add(variables);
            root.Children.Add(_error);
            root.Children.Add(buttons);
            Content = root;

            LoadEditor(s);
        }

        // Returns the column the user saved, or null when the window was cancelled.
        public static TemplateColumn Edit(Window owner, DataTable input, TemplateColumn seed, IList<TemplateColumn> existing, int editingIndex)
        {
            var win = new TemplateColumnWindow(owner, input, seed, existing, editingIndex);
            return win.ShowDialog() == true ? win._result : null;
        }

        static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.DimGray, Margin = new Thickness(0, 4, 0, 2) };
        }

        static TextBlock Hint(string text)
        {
            return new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.Gray, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0) };
        }

        bool MathRequired()
        {
            return _kind.SelectedIndex == 3;
        }

        bool CondRequired()
        {
            return _kind.SelectedIndex == 4;
        }

        void ShowKind()
        {
            int kind = Math.Max(0, _kind.SelectedIndex);
            _panelFixed.Visibility = kind == 0 ? Visibility.Visible : Visibility.Collapsed;
            _panelCopy.Visibility = kind == 1 ? Visibility.Visible : Visibility.Collapsed;
            _panelConcat.Visibility = kind == 2 ? Visibility.Visible : Visibility.Collapsed;
            _panelMath.Visibility = kind == 3 ? Visibility.Visible : Visibility.Collapsed;
            _panelConditional.Visibility = kind == 4 ? Visibility.Visible : Visibility.Collapsed;
            FormulaField.Refresh(_math);
            FormulaField.Refresh(_condition);
            FormulaField.Refresh(_trueValue);
            FormulaField.Refresh(_falseValue);
        }

        void LoadEditor(TemplateColumn col)
        {
            _name.Text = col.Name;
            if (col.Kind == TemplateColumnKind.Copy)
            {
                _kind.SelectedIndex = 1;
                _copySource.Text = "";
                _copySource.SelectedIndex = -1;
                string wantLetter = TemplatesWork.ExtractLetter(col.SourceColumn);
                for (var i = 0; i < _copySource.Items.Count; i++)
                {
                    string item = _copySource.Items[i] as string;
                    // Match by letter first (new templates store "A"), then by old header name.
                    if (string.Equals(TemplatesWork.ExtractLetter(item), wantLetter, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(item, col.SourceColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        _copySource.SelectedIndex = i;
                        break;
                    }
                }
                if (_copySource.SelectedIndex < 0)
                    _copySource.Text = col.SourceColumn;
            }
            else if (col.Kind == TemplateColumnKind.Concat)
            {
                _kind.SelectedIndex = 2;
                _pattern.Text = col.Pattern;
            }
            else if (col.Kind == TemplateColumnKind.Math)
            {
                _kind.SelectedIndex = 3;
                string expr = col.Expression;
                if (string.IsNullOrWhiteSpace(expr) && !string.IsNullOrWhiteSpace(col.Left))
                    expr = col.Left + " " + col.Operator + " " + col.Right;
                _math.Text = expr;
            }
            else if (col.Kind == TemplateColumnKind.Conditional)
            {
                _kind.SelectedIndex = 4;
                _condition.Text = col.Condition;
                _trueValue.Text = col.TrueExpression;
                _falseValue.Text = col.FalseExpression;
            }
            else
            {
                _kind.SelectedIndex = 0;
                _fixedValue.Text = col.FixedValue;
            }
            ShowKind();
        }

        TemplateColumn BuildColumn()
        {
            string name = _name.Text != null ? _name.Text.Trim() : "";
            if (name.Length == 0)
                throw new InvalidOperationException("Enter a column name.");

            var col = new TemplateColumn();
            col.Name = name;
            int kind = Math.Max(0, _kind.SelectedIndex);
            if (kind == 1)
            {
                col.Kind = TemplateColumnKind.Copy;
                string picked = _copySource.SelectedItem as string;
                if (string.IsNullOrWhiteSpace(picked))
                    picked = _copySource.Text != null ? _copySource.Text.Trim() : "";
                // Store the Excel letter (A, B, ...) so templates use position, not header names.
                col.SourceColumn = TemplatesWork.ExtractLetter(picked);
            }
            else if (kind == 2)
            {
                col.Kind = TemplateColumnKind.Concat;
                col.Pattern = _pattern.Text;
            }
            else if (kind == 3)
            {
                col.Kind = TemplateColumnKind.Math;
                col.Expression = _math.Text;
                // Keep legacy fields filled so old readers still see something.
                col.Left = _math.Text;
                col.Operator = "+";
                col.Right = "";
            }
            else if (kind == 4)
            {
                col.Kind = TemplateColumnKind.Conditional;
                col.Condition = _condition.Text;
                col.TrueExpression = _trueValue.Text;
                col.FalseExpression = _falseValue.Text;
            }
            else
            {
                col.Kind = TemplateColumnKind.Fixed;
                col.FixedValue = _fixedValue.Text;
            }
            return col;
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (_kind.SelectedIndex == 3 && !FormulaField.Check(_math))
                return;
            if (_kind.SelectedIndex == 4 &&
                (!FormulaField.Check(_condition) || !FormulaField.Check(_trueValue) || !FormulaField.Check(_falseValue)))
                return;

            TemplateColumn col;
            try
            {
                col = BuildColumn();
            }
            catch (Exception ex)
            {
                Fail(ex.Message);
                return;
            }

            string error = TemplatesWork.ValidateColumn(col);
            if (error != null)
            {
                Fail(error);
                return;
            }

            if (_existing != null)
            {
                for (var i = 0; i < _existing.Count; i++)
                {
                    if (i != _editingIndex &&
                        string.Equals((_existing[i].Name ?? "").Trim(), col.Name.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        Fail("Duplicate column name: " + col.Name.Trim() + ".");
                        return;
                    }
                }
            }

            _result = col;
            DialogResult = true;
        }

        void Fail(string message)
        {
            _error.Text = message;
            _error.Visibility = Visibility.Visible;
        }
    }
}
