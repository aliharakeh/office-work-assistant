using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.FormatSheet
{
    public partial class FormatSheetPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<FormatSheetSettings> _use;
        readonly List<ColumnFormat> _formats = new List<ColumnFormat>();

        public FormatSheetPage(DataTable input, string label, FormatSheetSettings settings, Action<FormatSheetSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - the look Save writes. Highlight colours from an earlier step stay on top.";
            foreach (var name in NamedColors.FillNames)
                BandColor.Items.Add(name);

            var s = settings ?? new FormatSheetSettings();
            HeaderBox.IsChecked = s.StyleHeader;
            HeaderPicker.Value = s.Header;
            FreezeBox.IsChecked = s.FreezeHeader;
            FreezeColsBox.Text = s.FreezeColumns.ToString(CultureInfo.InvariantCulture);
            FilterBox.IsChecked = s.AutoFilter;
            FitBox.IsChecked = s.AutoFit;
            BordersBox.IsChecked = s.Borders;
            BandBox.IsChecked = s.BandRows;
            BandColor.Text = NamedColors.NameOf(s.BandFill);
            _formats.AddRange(s.Columns);
            RebuildFormats();
        }

        void RebuildFormats()
        {
            FormatPanel.Children.Clear();
            if (_formats.Count == 0)
            {
                FormatPanel.Children.Add(new TextBlock { Text = "No column formats yet.", Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 0, 4) });
                return;
            }
            for (var i = 0; i < _formats.Count; i++)
            {
                var index = i;
                var text = Describe(_formats[i]);
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.Click += (s, e) => EditFormat(index);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveFormat(index);
                var buttons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                buttons.Children.Add(edit);
                buttons.Children.Add(remove);
                var row = new DockPanel { Margin = new Thickness(0, 0, 0, 4) };
                DockPanel.SetDock(buttons, Dock.Right);
                row.Children.Add(buttons);
                row.Children.Add(new TextBlock
                {
                    Text = text,
                    ToolTip = text,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                });
                FormatPanel.Children.Add(row);
            }
        }

        // "C - City: width 20, format #,##0" reads better than the raw letter.
        string Describe(ColumnFormat f)
        {
            var at = ExcelColumns.IndexOf(_input, f.Column);
            var label = at >= 0 ? ExcelColumns.Label(_input, at) : f.Column;
            var parts = new List<string>();
            if (f.Width > 0)
                parts.Add("width " + f.Width.ToString("0.#", CultureInfo.InvariantCulture));
            if (!string.IsNullOrWhiteSpace(f.NumberFormat))
                parts.Add("format " + f.NumberFormat);
            if (f.Align != CellAlign.General)
                parts.Add(f.Align.ToString().ToLowerInvariant());
            if (f.Wrap)
                parts.Add("wrap");
            return label + " - " + (parts.Count == 0 ? "nothing" : string.Join(", ", parts.ToArray()));
        }

        void Add_Click(object sender, RoutedEventArgs e)
        {
            EditFormat(-1);
        }

        // index -1 adds a new format.
        void EditFormat(int index)
        {
            var seed = index >= 0 ? _formats[index] : null;
            var f = ColumnFormatWindow.Edit(Window.GetWindow(this), _input, seed);
            if (f == null)
                return;
            if (index >= 0)
                _formats[index] = f;
            else
                _formats.Add(f);
            RebuildFormats();
            // The pipeline applies this editor by itself after a click inside it; the dialog's clicks never reach it.
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void RemoveFormat(int index)
        {
            _formats.RemoveAt(index);
            RebuildFormats();
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            int freeze;
            if (!int.TryParse(FreezeColsBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out freeze) || freeze < 0)
            {
                Status.Text = "Columns to freeze must be 0 or a whole number above 0.";
                FreezeColsBox.Focus();
                return;
            }
            var s = new FormatSheetSettings
            {
                StyleHeader = HeaderBox.IsChecked == true,
                Header = HeaderPicker.Value,
                FreezeHeader = FreezeBox.IsChecked == true,
                FreezeColumns = freeze,
                AutoFilter = FilterBox.IsChecked == true,
                AutoFit = FitBox.IsChecked == true,
                Borders = BordersBox.IsChecked == true,
                BandRows = BandBox.IsChecked == true,
                BandFill = NamedColors.ToHex(BandColor.Text) ?? BandColor.Text.Trim()
            };
            s.Columns.AddRange(_formats);
            try
            {
                FormatSheetWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                Notices.Show(this, ex.Message, "Format sheet", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = FormatSheetWork.Summary(s) + ".";
            _use(s);
        }
    }
}
