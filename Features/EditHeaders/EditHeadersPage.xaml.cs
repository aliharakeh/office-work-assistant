using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.EditHeaders
{
    public partial class EditHeadersPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<EditHeadersSettings> _use;
        readonly List<EditRow> _rows = new List<EditRow>();

        public EditHeadersPage(DataTable input, string label, EditHeadersSettings settings, Action<EditHeadersSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - " + input.Columns.Count + " columns, " + input.Rows.Count + " rows";
            AllFormula.Text = settings == null ? "" : settings.Formula;
            FormulaField.WatchOptional(AllFormula);
            for (var i = 0; i < input.Columns.Count; i++)
            {
                HeaderEdit found = null;
                if (settings != null)
                    found = settings.Columns.Find(c => ExcelColumns.IndexOf(input, c.Source) == i);
                AddRow(found ?? new HeaderEdit { Source = ExcelColumns.Letter(i) });
            }
            if (settings != null)
                foreach (var edit in settings.Columns)
                    if (string.IsNullOrWhiteSpace(edit.Source)) AddRow(edit);
        }

        void AddRow(HeaderEdit edit)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };
            bool added = string.IsNullOrWhiteSpace(edit.Source);
            panel.Children.Add(new TextBlock { Text = added ? "New column" : ExcelColumns.Label(_input, ExcelColumns.Index(_input, edit.Source)), TextWrapping = TextWrapping.Wrap });
            var line = new StackPanel();
            line.Children.Add(new TextBlock { Text = "Header action", Margin = new Thickness(0, 4, 0, 2) });
            var mode = new ComboBox { ItemsSource = Enum.GetValues(typeof(HeaderEditMode)), SelectedItem = edit.Mode, Margin = new Thickness(0, 0, 0, 4) };
            line.Children.Add(mode);
            var row = new EditRow { Source = edit.Source, Panel = panel, Mode = mode, Value = new TextBox { Text = edit.Value ?? "", Margin = new Thickness(0, 2, 0, 0) } };
            if (added)
            {
                var remove = new Button { Content = "Remove", Margin = new Thickness(6, 2, 0, 0) };
                remove.Click += (s, e) => { _rows.Remove(row); Rows.Children.Remove(panel); };
                line.Children.Add(remove);
            }
            line.Children.Insert(2, new TextBlock { Text = "Header text or formula", Margin = new Thickness(0, 4, 0, 2) });
            line.Children.Insert(3, row.Value);
            panel.Children.Add(line);
            mode.SelectionChanged += (s, e) => UpdateRow(row);
            UpdateRow(row);
            _rows.Add(row);
            Rows.Children.Add(panel);
        }

        static void UpdateRow(EditRow row)
        {
            var mode = (HeaderEditMode)row.Mode.SelectedItem;
            row.Value.IsEnabled = mode == HeaderEditMode.Text || mode == HeaderEditMode.Formula;
        }

        void Add_Click(object sender, RoutedEventArgs e) { AddRow(new HeaderEdit { Mode = HeaderEditMode.Text }); }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var settings = new EditHeadersSettings { Formula = AllFormula.Text };
            foreach (var row in _rows)
            {
                var mode = (HeaderEditMode)row.Mode.SelectedItem;
                if (mode != HeaderEditMode.Keep || string.IsNullOrWhiteSpace(row.Source))
                    settings.Columns.Add(new HeaderEdit { Source = row.Source, Mode = mode, Value = row.Value.Text });
            }
            try
            {
                var output = EditHeadersWork.Run(_input, settings);
                Status.Text = output.Columns.Count + " columns; data rows preserved.";
                _use(settings);
            }
            catch (Exception ex) { Notices.Show(this, ex.Message, "Edit headers", MessageBoxButton.OK, MessageBoxImage.Warning); }
        }

        sealed class EditRow
        {
            public string Source;
            public StackPanel Panel;
            public ComboBox Mode;
            public TextBox Value;
        }
    }
}
