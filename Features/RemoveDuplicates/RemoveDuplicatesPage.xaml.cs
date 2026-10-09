using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.RemoveDuplicates
{
    public partial class RemoveDuplicatesPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<RemoveDuplicatesSettings> _use;
        readonly List<Pick> _picks = new List<Pick>();

        public RemoveDuplicatesPage(DataTable input, string label, RemoveDuplicatesSettings settings, Action<RemoveDuplicatesSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - " + input.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows";

            var s = settings ?? new RemoveDuplicatesSettings();
            var keys = new HashSet<int>();
            foreach (var k in s.Keys)
            {
                int i = ExcelColumns.IndexOf(input, k);
                if (i >= 0)
                    keys.Add(i);
            }
            for (var i = 0; i < input.Columns.Count; i++)
                _picks.Add(new Pick { Letter = ExcelColumns.Letter(i), Label = ExcelColumns.Label(input, i), Include = keys.Contains(i) });
            KeyList.ItemsSource = _picks;
            FormulaField.Watch(FormulaBox, FormulaRequired);
            FormulaBox.Text = s.KeyFormula ?? "";
            ByFormulaBox.IsChecked = s.UseFormula;
            ByColumnsBox.IsChecked = !s.UseFormula;
            IgnoreCaseBox.IsChecked = s.IgnoreCase;
            TrimBox.IsChecked = s.TrimSpaces;
            KeepBox.SelectedIndex = (int)s.Keep;
            CountBox.Text = s.CountHeader ?? "";
        }

        bool FormulaRequired()
        {
            return ByFormulaBox != null && ByFormulaBox.IsChecked == true;
        }

        void Mode_Changed(object sender, RoutedEventArgs e)
        {
            if (FormulaPanel == null || ColumnsPanel == null)
                return;
            bool formula = ByFormulaBox.IsChecked == true;
            FormulaPanel.Visibility = formula ? Visibility.Visible : Visibility.Collapsed;
            ColumnsPanel.Visibility = formula ? Visibility.Collapsed : Visibility.Visible;
            FormulaField.Refresh(FormulaBox);
        }

        void Help_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "The formula is worked out for each row; rows that give the same result count as duplicates. $Name or $A is a column of the row (A = 1st).");
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            if (ByFormulaBox.IsChecked == true && !FormulaField.Check(FormulaBox))
                return;
            var s = new RemoveDuplicatesSettings
            {
                UseFormula = ByFormulaBox.IsChecked == true,
                KeyFormula = FormulaBox.Text.Trim(),
                IgnoreCase = IgnoreCaseBox.IsChecked == true,
                TrimSpaces = TrimBox.IsChecked == true,
                Keep = (DuplicateKeep)Math.Max(0, KeepBox.SelectedIndex),
                CountHeader = CountBox.Text.Trim()
            };
            foreach (var p in _picks)
            {
                if (p.Include)
                    s.Keys.Add(p.Letter);
            }
            DataTable result;
            try
            {
                result = RemoveDuplicatesWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Remove duplicates", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = _input.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows in, " +
                result.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows out.";
            _use(s);
        }

        public sealed class Pick
        {
            public string Letter { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }
    }
}
