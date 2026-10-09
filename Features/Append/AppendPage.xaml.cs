using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace OfficeWorkAssistant.Features.Append
{
    public partial class AppendPage : UserControl
    {
        readonly IList<DataTable> _inputs;
        readonly IList<string> _names;
        readonly Action<AppendSettings> _use;

        // labels: shown beside each input. names: what the source column says for its rows.
        public AppendPage(IList<DataTable> inputs, IList<string> labels, IList<string> names, AppendSettings settings, Action<AppendSettings> use)
        {
            InitializeComponent();
            _inputs = inputs;
            _names = names;
            _use = use;

            var lines = new List<string>();
            for (var i = 0; i < inputs.Count; i++)
                lines.Add((i + 1).ToString(CultureInfo.InvariantCulture) + ". " + labels[i] + " - " +
                    inputs[i].Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows, " +
                    inputs[i].Columns.Count.ToString(CultureInfo.InvariantCulture) + " columns");
            InputList.ItemsSource = lines;

            var s = settings ?? new AppendSettings();
            ByHeaderBox.IsChecked = s.Match == AppendMatch.ByHeader;
            ByPositionBox.IsChecked = s.Match == AppendMatch.ByPosition;
            CommonBox.IsChecked = s.CommonOnly;
            SourceBox.Text = s.SourceHeader ?? "";
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var s = new AppendSettings
            {
                Match = ByPositionBox.IsChecked == true ? AppendMatch.ByPosition : AppendMatch.ByHeader,
                CommonOnly = CommonBox.IsChecked == true,
                SourceHeader = SourceBox.Text.Trim()
            };
            DataTable result;
            try
            {
                result = AppendWork.Run(_inputs, _names, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Append tables", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = result.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows, " +
                result.Columns.Count.ToString(CultureInfo.InvariantCulture) + " columns.";
            _use(s);
        }
    }
}
