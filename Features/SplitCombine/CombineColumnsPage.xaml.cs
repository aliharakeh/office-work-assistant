using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.SplitCombine
{
    public partial class CombineColumnsPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<CombineColumnsSettings> _use;
        // Column indexes in join order.
        readonly List<int> _parts = new List<int>();

        public CombineColumnsPage(DataTable input, string label, CombineColumnsSettings settings, Action<CombineColumnsSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label;
            ColumnBox.ItemsSource = ExcelColumns.Labels(input);
            if (input.Columns.Count > 0)
                ColumnBox.SelectedIndex = 0;

            var s = settings ?? new CombineColumnsSettings();
            foreach (var c in s.Columns)
            {
                int i = ExcelColumns.IndexOf(input, c);
                if (i >= 0)
                    _parts.Add(i);
            }
            SeparatorBox.Text = s.Separator;
            HeaderBox.Text = s.Header;
            SkipBox.IsChecked = s.SkipBlanks;
            RemoveBox.IsChecked = s.RemoveSources;
            ShowParts(-1);
        }

        void ShowParts(int select)
        {
            var labels = new List<string>();
            foreach (var i in _parts)
                labels.Add(ExcelColumns.Label(_input, i));
            PartList.ItemsSource = labels;
            PartList.SelectedIndex = select;
        }

        void Add_Click(object sender, RoutedEventArgs e)
        {
            if (ColumnBox.SelectedIndex < 0)
                return;
            _parts.Add(ColumnBox.SelectedIndex);
            ShowParts(_parts.Count - 1);
            if (ColumnBox.SelectedIndex < ColumnBox.Items.Count - 1)
                ColumnBox.SelectedIndex++;
        }

        void Remove_Click(object sender, RoutedEventArgs e)
        {
            int index = PartList.SelectedIndex;
            if (index < 0)
                return;
            _parts.RemoveAt(index);
            ShowParts(Math.Min(index, _parts.Count - 1));
        }

        void Up_Click(object sender, RoutedEventArgs e)
        {
            Move(-1);
        }

        void Down_Click(object sender, RoutedEventArgs e)
        {
            Move(1);
        }

        void Move(int delta)
        {
            int index = PartList.SelectedIndex;
            int target = index + delta;
            if (index < 0 || target < 0 || target >= _parts.Count)
                return;
            int item = _parts[index];
            _parts[index] = _parts[target];
            _parts[target] = item;
            ShowParts(target);
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var s = new CombineColumnsSettings
            {
                Separator = SeparatorBox.Text,
                Header = HeaderBox.Text.Trim(),
                SkipBlanks = SkipBox.IsChecked == true,
                RemoveSources = RemoveBox.IsChecked == true
            };
            foreach (var i in _parts)
                s.Columns.Add(ExcelColumns.Letter(i));
            try
            {
                SplitCombineWork.Combine(_input, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Combine columns", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = "Joins " + s.Columns.Count + " columns into \"" + s.Header + "\".";
            _use(s);
        }
    }
}
