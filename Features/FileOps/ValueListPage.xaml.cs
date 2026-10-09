using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FileOps
{
    // Pipeline-only: edits a List / set / map step. input is null when no step is connected.
    public partial class ValueListPage : UserControl
    {
        const int PreviewRows = 1000;

        readonly DataTable _input;
        readonly Action<ValueListSettings> _use;
        readonly ObservableCollection<ValueListItem> _items = new ObservableCollection<ValueListItem>();
        bool _filling;

        public ValueListPage(DataTable input, string label, ValueListSettings settings, Action<ValueListSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            FormulaField.Watch(KeyBox, NeedKey);
            FormulaField.Watch(ValueBox, NeedValue);
            StoredGrid.ItemsSource = _items;
            InputInfo.Text = input == null
                ? "No input step is connected: type the values below, or connect a step and edit again."
                : label + " (" + input.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows; columns: " + Headers(input) + ")";

            var s = settings ?? new ValueListSettings
            {
                Source = input == null ? ValueListSource.Stored : ValueListSource.Input,
                KeyFormula = input != null && input.Columns.Count > 0 ? "$" + ExcelFile.Header(input.Columns[0]) : ""
            };
            _filling = true;
            SrcInput.IsChecked = s.Source == ValueListSource.Input;
            SrcStored.IsChecked = s.Source == ValueListSource.Stored;
            ModeBox.SelectedIndex = (int)s.Mode;
            KeyBox.Text = s.KeyFormula ?? "";
            ValueBox.Text = s.ValueFormula ?? "";
            TrimBox.IsChecked = s.Trim;
            SkipBox.IsChecked = s.SkipBlanks;
            DupBox.SelectedIndex = (int)s.OnDuplicate;
            foreach (var item in s.Items)
                _items.Add(new ValueListItem { Key = item.Key, Value = item.Value });
            _filling = false;
            UpdateEnabled();
        }

        static string Headers(DataTable table)
        {
            var names = new string[Math.Min(table.Columns.Count, 12)];
            for (var i = 0; i < names.Length; i++)
                names[i] = ExcelFile.Header(table.Columns[i]);
            return string.Join(", ", names) + (table.Columns.Count > names.Length ? ", ..." : "");
        }

        bool NeedKey()
        {
            return SrcInput.IsChecked == true;
        }

        bool NeedValue()
        {
            return SrcInput.IsChecked == true && ModeBox.SelectedIndex == (int)ValueListMode.Map;
        }

        void Changed(object sender, RoutedEventArgs e)
        {
            if (_filling)
                return;
            UpdateEnabled();
        }

        void UpdateEnabled()
        {
            var input = SrcInput.IsChecked == true;
            var map = ModeBox.SelectedIndex == (int)ValueListMode.Map;
            KeyBox.IsEnabled = input;
            ValueBox.IsEnabled = input && map;
            DupBox.IsEnabled = map;
            DupLabel.IsEnabled = map;
            CaptureBtn.IsEnabled = _input != null;
            StoredGroup.IsEnabled = !input;
            FormulaField.Refresh(KeyBox);
            FormulaField.Refresh(ValueBox);
        }

        ValueListSettings Current()
        {
            StoredGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var s = new ValueListSettings
            {
                Source = SrcStored.IsChecked == true ? ValueListSource.Stored : ValueListSource.Input,
                Mode = (ValueListMode)Math.Max(0, ModeBox.SelectedIndex),
                KeyFormula = KeyBox.Text.Trim(),
                ValueFormula = ValueBox.Text.Trim(),
                Trim = TrimBox.IsChecked == true,
                SkipBlanks = SkipBox.IsChecked == true,
                OnDuplicate = (MapDuplicate)Math.Max(0, DupBox.SelectedIndex)
            };
            foreach (var item in _items)
            {
                if (item != null && (!string.IsNullOrEmpty(item.Key) || !string.IsNullOrEmpty(item.Value)))
                    s.Items.Add(new ValueListItem { Key = item.Key ?? "", Value = item.Value ?? "" });
            }
            return s;
        }

        void Capture_Click(object sender, RoutedEventArgs e)
        {
            if (_input == null)
                return;
            try
            {
                var items = FileOpsWork.Capture(_input, Current());
                _items.Clear();
                foreach (var item in items)
                    _items.Add(item);
                SrcStored.IsChecked = true;
                UpdateEnabled();
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not capture", MessageBoxImage.Warning);
            }
        }

        void Use_Click(object sender, RoutedEventArgs e)
        {
            var s = Current();
            var error = FileOpsWork.Check(s);
            if (error != null)
            {
                Alert(error, "Not ready", MessageBoxImage.Warning);
                return;
            }
            _use(s);
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }
    }
}
