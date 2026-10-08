using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Microsoft.Win32;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;
using OfficeWorkAssistant.Views;

namespace OfficeWorkAssistant.Features.MergeColumns
{
    public partial class MergeColumnsPage : Page
    {
        DataTable _tableA;
        DataTable _tableB;
        DataTable _fullA;
        DataTable _fullB;
        bool _cutA;
        bool _cutB;
        int _previewRows = 1000;
        bool _loadingSheet;
        readonly List<KeyRow> _keys = new List<KeyRow>();
        readonly List<RuleRow> _rules = new List<RuleRow>();

        public MergeColumnsPage()
        {
            InitializeComponent();
            PreviewBox.Items.Add(200);
            PreviewBox.Items.Add(1000);
            PreviewBox.Items.Add(5000);
            PreviewBox.SelectedItem = 1000;
            ExcelGrid.Hook(GridA, GridB, GridOut);
            AddKey();
            AddRule();
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
            if (!_loadingSheet)
                ReloadSheet(true);
        }

        void SheetB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_loadingSheet)
                ReloadSheet(false);
        }

        void PreviewBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewBox.SelectedItem == null)
                return;
            var rows = (int)PreviewBox.SelectedItem;
            if (rows == _previewRows)
                return;
            _previewRows = rows;
            if (!string.IsNullOrEmpty(PathA.Text))
                ReloadSheet(true);
            if (!string.IsNullOrEmpty(PathB.Text))
                ReloadSheet(false);
        }

        void AddKey_Click(object sender, RoutedEventArgs e)
        {
            AddKey();
        }

        void AddRule_Click(object sender, RoutedEventArgs e)
        {
            AddRule();
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            var extras = new List<VariableHelp>
            {
                new VariableHelp { Name = "$A_Name", Description = "Column Name (or letter, $A_C) of the matched A row. Blank when no A row matched.", Example = "" },
                new VariableHelp { Name = "$B_Name", Description = "Column Name (or letter) of the B row being filled, including results of earlier rules.", Example = "" },
                new VariableHelp { Name = "$Name", Description = "Same as $B_Name. A name B does not have falls back to A.", Example = "" },
                new VariableHelp { Name = "$Value", Description = "The target cell's current value.", Example = "" },
                new VariableHelp { Name = "$Matched", Description = "true when an A row matched this B row.", Example = "" }
            };
            ExpressionHelp.Show(Window.GetWindow(this),
                "Example: Target Price, When $Matched && $Value == \"\", Value $A_Price. Add a second Price rule below it with Value $Value * 1.1 to change the rows the first rule skipped.",
                extras);
        }

        void Save_Click(object sender, RoutedEventArgs e)
        {
            if (_tableA == null || _tableB == null)
            {
                Alert("Open both Excel files first.", "Nothing to save", MessageBoxImage.Warning);
                return;
            }
            if (!FormulasOk())
            {
                Alert("Fix the formulas marked in red first.", "Formula error", MessageBoxImage.Warning);
                return;
            }

            try
            {
                var result = MergeColumnsWork.Merge(FullTable(true), FullTable(false), Keys(), Rules(), 0);
                var loaded = new List<KeyValuePair<string, string>>();
                loaded.Add(new KeyValuePair<string, string>("File B", PathB.Text));
                loaded.Add(new KeyValuePair<string, string>("File A", PathA.Text));
                var name = Path.GetFileNameWithoutExtension(PathB.Text) + "_merged.xlsx";
                var done = ExcelSaveDialog.Show(Window.GetWindow(this), result.Table, name, loaded);
                if (done != null)
                    Alert(done, "Done", MessageBoxImage.Information);
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
                Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                Title = isA ? "Open File A" : "Open File B"
            };
            if (dlg.ShowDialog() != true)
                return;
            try
            {
                ApplyLoad(isA, MergeColumnsWork.Load(dlg.FileName, null, _previewRows), dlg.FileName);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not open file", MessageBoxImage.Error);
            }
        }

        void ReloadSheet(bool isA)
        {
            var path = isA ? PathA.Text : PathB.Text;
            var sheet = (isA ? SheetA.SelectedItem : SheetB.SelectedItem) as string;
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(sheet))
                return;
            try
            {
                ApplyLoad(isA, MergeColumnsWork.Load(path, sheet, _previewRows), path);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not load sheet", MessageBoxImage.Error);
            }
        }

        void ApplyLoad(bool isA, ExcelLoadResult loaded, string path)
        {
            _loadingSheet = true;
            try
            {
                if (isA)
                {
                    _tableA = loaded.Table;
                    _fullA = null;
                    _cutA = loaded.Truncated;
                }
                else
                {
                    _tableB = loaded.Table;
                    _fullB = null;
                    _cutB = loaded.Truncated;
                }
                (isA ? PathA : PathB).Text = path;
                var sheetBox = isA ? SheetA : SheetB;
                sheetBox.ItemsSource = loaded.Sheets;
                sheetBox.SelectedItem = loaded.Sheet;
                (isA ? GridA : GridB).ItemsSource = loaded.Table.DefaultView;
                (isA ? InfoA : InfoB).Text = RowInfo(loaded);
                if (!isA)
                    RefreshTargets();
            }
            finally
            {
                _loadingSheet = false;
            }
            RefreshPreview();
        }

        static string RowInfo(ExcelLoadResult loaded)
        {
            var rows = loaded.Table.Rows.Count.ToString("N0", CultureInfo.InvariantCulture);
            if (!loaded.Truncated)
                return rows + " rows";
            return "Showing " + rows + " of " + loaded.TotalRows.ToString("N0", CultureInfo.InvariantCulture) +
                   " rows (all rows are used when you save)";
        }

        // The grids show a capped preview; matching needs every A row, so load it once and cache it.
        DataTable FullTable(bool isA)
        {
            var cached = isA ? _fullA : _fullB;
            if (cached != null)
                return cached;
            if (!(isA ? _cutA : _cutB))
                return isA ? _tableA : _tableB;

            var path = isA ? PathA.Text : PathB.Text;
            var sheet = (isA ? SheetA.SelectedItem : SheetB.SelectedItem) as string;
            var table = MergeColumnsWork.Load(path, sheet, 0).Table;
            if (isA)
                _fullA = table;
            else
                _fullB = table;
            return table;
        }

        void RefreshPreview()
        {
            if (_loadingSheet)
                return;
            GridOut.ItemsSource = null;
            if (_tableA == null || _tableB == null)
            {
                Status.Text = "Open both files and set a match rule.";
                return;
            }
            if (!FormulasOk())
            {
                Status.Text = "Fix the formulas marked in red. Hover one to see why.";
                return;
            }

            try
            {
                var a = FullTable(true);
                var keys = Keys();
                var rules = Rules();
                var result = MergeColumnsWork.Merge(a, _tableB, keys, rules, _previewRows);
                GridOut.ItemsSource = result.Table.DefaultView;

                var text = result.Matched.ToString("N0", CultureInfo.InvariantCulture) + " matched, " +
                           result.Unmatched.ToString("N0", CultureInfo.InvariantCulture) + " not matched, " +
                           result.Filled.ToString("N0", CultureInfo.InvariantCulture) + " cells filled";
                if (_cutB)
                    text += " (preview rows only)";
                text += ".";
                if (result.SharedKeys > 0)
                    text += " " + result.SharedKeys.ToString("N0", CultureInfo.InvariantCulture) +
                            " keys appear on more than one A row; the first A row is used.";
                var unknown = MergeColumnsWork.UnknownNames(a, result.Table, keys, rules);
                if (unknown.Length > 0)
                    text += " Unknown names read as blank: $" + string.Join(", $", unknown) + ".";
                Status.Text = text;
            }
            catch (Exception ex)
            {
                Status.Text = ex.Message;
            }
        }

        bool FormulasOk()
        {
            foreach (var k in _keys)
            {
                if (!FormulaField.Ok(k.ExprA) || !FormulaField.Ok(k.ExprB))
                    return false;
            }
            foreach (var r in _rules)
            {
                if (!FormulaField.Ok(r.When) || !FormulaField.Ok(r.Value))
                    return false;
            }
            return true;
        }

        List<MergeColumnsKey> Keys()
        {
            var list = new List<MergeColumnsKey>();
            foreach (var k in _keys)
                list.Add(new MergeColumnsKey { FormulaA = k.ExprA.Text, FormulaB = k.ExprB.Text });
            return list;
        }

        List<MergeColumnsRule> Rules()
        {
            var list = new List<MergeColumnsRule>();
            foreach (var r in _rules)
                list.Add(new MergeColumnsRule { Target = r.Target.Text, When = r.When.Text, Value = r.Value.Text });
            return list;
        }

        void AddKey()
        {
            var line = new Grid { Margin = new Thickness(0, 0, 8, 4) };
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(32) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var item = new KeyRow
            {
                ExprA = new TextBox { Text = "$A", ToolTip = "Formula on the A row, e.g. $ID or TRIM($Code)." },
                ExprB = new TextBox { Text = "$A", ToolTip = "Formula on the B row, e.g. $ID or TRIM($Code)." }
            };
            var eq = new TextBlock { Text = "==", HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            var del = new Button { Content = "Remove", Width = 64, Margin = new Thickness(4, 0, 0, 0) };
            Grid.SetColumn(eq, 1);
            Grid.SetColumn(item.ExprB, 2);
            Grid.SetColumn(del, 3);
            line.Children.Add(item.ExprA);
            line.Children.Add(eq);
            line.Children.Add(item.ExprB);
            line.Children.Add(del);

            FormulaField.Watch(item.ExprA);
            FormulaField.Watch(item.ExprB);
            item.ExprA.TextChanged += Formula_Changed;
            item.ExprB.TextChanged += Formula_Changed;
            del.Click += delegate
            {
                _keys.Remove(item);
                KeyPanel.Children.Remove(line);
                if (_keys.Count == 0)
                    AddKey();
                else
                    RefreshPreview();
            };

            _keys.Add(item);
            KeyPanel.Children.Add(line);
            RefreshPreview();
        }

        void AddRule()
        {
            var line = new Grid { Margin = new Thickness(0, 0, 8, 4) };
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var item = new RuleRow
            {
                Target = new ComboBox { IsEditable = true, Margin = new Thickness(0, 0, 4, 0), ToolTip = "Pick a B column to fill, or type a new column name." },
                When = new TextBox { Text = "$Matched", Margin = new Thickness(0, 0, 4, 0), ToolTip = "Optional condition, e.g. $Matched && $Value == \"\". Blank = always." },
                Value = new TextBox { Text = "$Value", Margin = new Thickness(0, 0, 4, 0), ToolTip = "Value formula, e.g. $A_Price or IF($A_Qty > 0, $A_Price, $Value)." }
            };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal };
            var up = new Button { Content = "Up", Width = 52, Margin = new Thickness(0, 0, 4, 0), ToolTip = "Run this rule earlier." };
            var del = new Button { Content = "Remove", Width = 64 };
            buttons.Children.Add(up);
            buttons.Children.Add(del);
            Grid.SetColumn(item.When, 1);
            Grid.SetColumn(item.Value, 2);
            Grid.SetColumn(buttons, 3);
            line.Children.Add(item.Target);
            line.Children.Add(item.When);
            line.Children.Add(item.Value);
            line.Children.Add(buttons);

            FillTargets(item.Target);
            FormulaField.WatchOptional(item.When);
            FormulaField.Watch(item.Value);
            item.Target.SelectionChanged += delegate { Dispatcher.BeginInvoke(new Action(RefreshPreview)); };
            item.Target.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler(Formula_Changed));
            item.When.TextChanged += Formula_Changed;
            item.Value.TextChanged += Formula_Changed;
            up.Click += delegate
            {
                var i = _rules.IndexOf(item);
                if (i <= 0)
                    return;
                _rules.RemoveAt(i);
                _rules.Insert(i - 1, item);
                RulePanel.Children.Remove(line);
                RulePanel.Children.Insert(i - 1, line);
                RefreshPreview();
            };
            del.Click += delegate
            {
                _rules.Remove(item);
                RulePanel.Children.Remove(line);
                if (_rules.Count == 0)
                    AddRule();
                else
                    RefreshPreview();
            };

            _rules.Add(item);
            RulePanel.Children.Add(line);
            RefreshPreview();
        }

        void FillTargets(ComboBox box)
        {
            var keep = box.Text;
            var names = new List<string>();
            if (_tableB != null)
            {
                foreach (DataColumn col in _tableB.Columns)
                    names.Add(ExcelFile.Header(col));
            }
            box.ItemsSource = names;
            box.Text = keep;
        }

        void RefreshTargets()
        {
            foreach (var r in _rules)
                FillTargets(r.Target);
        }

        void Formula_Changed(object sender, TextChangedEventArgs e)
        {
            RefreshPreview();
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }

        sealed class KeyRow
        {
            public TextBox ExprA;
            public TextBox ExprB;
        }

        sealed class RuleRow
        {
            public ComboBox Target;
            public TextBox When;
            public TextBox Value;
        }
    }
}
