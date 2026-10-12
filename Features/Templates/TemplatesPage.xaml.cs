using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.Templates
{
    public partial class TemplatesPage : UserControl
    {
        DataTable _source;
        int _previewRows = 1000;
        List<TemplateColumn> _columns = new List<TemplateColumn>();
        List<SourceColumnPick> _sourcePicks;
        bool _syncSource;
        // The input comes from an earlier pipeline step, and "Use in pipeline"
        // hands the template back.
        DataTable _pipeInput;
        string _pipeLabel;
        Action<TemplateDefinition> _use;

        public TemplatesPage(DataTable input, string label, TemplateDefinition template, Action<TemplateDefinition> use)
        {
            // KeepSource.Checked fires during InitializeComponent before later
            // controls exist, so block handlers until load finishes.
            _syncSource = true;
            try
            {
                InitializeComponent();
            }
            finally
            {
                _syncSource = false;
            }
            _pipeInput = input;
            _pipeLabel = label;
            _use = use;
            PreviewBox.Items.Add(200);
            PreviewBox.Items.Add(1000);
            PreviewBox.Items.Add(5000);
            PreviewBox.SelectedItem = 1000;
            RebuildColumns();
            SyncSelectAll();
            ExcelGrid.Hook(GridSource, GridOut);
            InjectInput();
            if (template != null)
                ApplyTemplateDefinition(template);
        }

        // The grid shows a capped preview of the pipeline input.
        void InjectInput()
        {
            bool cut = _pipeInput.Rows.Count > _previewRows;
            _source = Capped(_pipeInput, _previewRows);
            SourcePath.Text = _pipeLabel;
            GridSource.ItemsSource = _source.DefaultView;
            InfoSource.Text = RowInfo(_source.Rows.Count, _pipeInput.Rows.Count, cut);
            RebuildSourcePicks(null);
        }

        void PreviewBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewBox.SelectedItem == null)
                return;
            var rows = (int)PreviewBox.SelectedItem;
            if (rows == _previewRows)
                return;
            _previewRows = rows;
            InjectInput();
        }

        static string RowInfo(int shown, int total, bool cut)
        {
            var rows = shown.ToString("N0", CultureInfo.InvariantCulture);
            if (!cut)
                return rows + " rows";
            return "Showing " + rows + " of " + total.ToString("N0", CultureInfo.InvariantCulture) +
                   " rows (all rows are used when you preview or save)";
        }

        // Generating output needs every row, not just the capped preview.
        DataTable FullSource()
        {
            return _pipeInput;
        }

        // Keep the output grid light; Save writes the full table.
        static DataTable Capped(DataTable table, int maxRows)
        {
            if (maxRows <= 0 || table.Rows.Count <= maxRows)
                return table;
            var copy = table.Clone();
            for (var i = 0; i < maxRows; i++)
                copy.ImportRow(table.Rows[i]);
            return copy;
        }

        void RebuildSourcePicks(IList<string> keep)
        {
            _syncSource = true;
            try
            {
                _sourcePicks = new List<SourceColumnPick>();
                if (_source != null)
                {
                    HashSet<string> wantedNames = null;
                    HashSet<int> wantedLetters = null;
                    if (keep != null && keep.Count > 0)
                    {
                        wantedNames = new HashSet<string>(keep, StringComparer.OrdinalIgnoreCase);
                        wantedLetters = new HashSet<int>();
                        for (var k = 0; k < keep.Count; k++)
                        {
                            int li;
                            if (TemplatesWork.TryParseColumnLetter(TemplatesWork.ExtractLetter(keep[k]), out li))
                                wantedLetters.Add(li);
                        }
                    }
                    for (var c = 0; c < _source.Columns.Count; c++)
                    {
                        DataColumn col = _source.Columns[c];
                        bool include = wantedNames == null ||
                            wantedNames.Contains(col.ColumnName) ||
                            wantedNames.Contains(ExcelFile.Header(col)) ||
                            (wantedLetters != null && wantedLetters.Contains(c));
                        _sourcePicks.Add(new SourceColumnPick
                        {
                            Name = col.ColumnName,
                            Label = TemplatesWork.SourceLabel(_source, c),
                            Include = include
                        });
                    }
                }
                SourceColumnList.ItemsSource = null;
                SourceColumnList.ItemsSource = _sourcePicks;
                SyncSelectAll();
            }
            finally
            {
                _syncSource = false;
            }
        }

        void KeepSource_Changed(object sender, RoutedEventArgs e)
        {
            if (_syncSource || KeepSource == null || SourceColumnList == null || SelectAllSource == null)
                return;
            SyncSelectAll();
        }

        void SourceColumn_Changed(object sender, RoutedEventArgs e)
        {
            if (_syncSource || KeepSource == null || SourceColumnList == null || SelectAllSource == null)
                return;
            SyncSelectAll();
        }

        void SelectAllSource_Click(object sender, RoutedEventArgs e)
        {
            if (_syncSource || _sourcePicks == null)
                return;
            bool on = SelectAllSource.IsChecked != false;
            _syncSource = true;
            try
            {
                for (var i = 0; i < _sourcePicks.Count; i++)
                    _sourcePicks[i].Include = on;
                SourceColumnList.ItemsSource = null;
                SourceColumnList.ItemsSource = _sourcePicks;
                KeepSource.IsChecked = true;
            }
            finally
            {
                _syncSource = false;
            }
            SyncSelectAll();
        }

        void SyncSelectAll()
        {
            if (KeepSource == null || SourceColumnList == null || SelectAllSource == null)
                return;
            _syncSource = true;
            try
            {
                bool enabled = KeepSource.IsChecked != false && _sourcePicks != null && _sourcePicks.Count > 0;
                SourceColumnList.IsEnabled = enabled;
                SelectAllSource.IsEnabled = enabled;
                if (!enabled)
                {
                    SelectAllSource.IsChecked = false;
                    return;
                }
                var all = true;
                var none = true;
                for (var i = 0; i < _sourcePicks.Count; i++)
                {
                    if (_sourcePicks[i].Include)
                        none = false;
                    else
                        all = false;
                }
                if (all)
                    SelectAllSource.IsChecked = true;
                else if (none)
                    SelectAllSource.IsChecked = false;
                else
                    SelectAllSource.IsChecked = null;
            }
            finally
            {
                _syncSource = false;
            }
        }

        List<string> SelectedSourceColumns()
        {
            var list = new List<string>();
            if (KeepSource.IsChecked == false || _sourcePicks == null || _source == null)
                return list;
            // Store Excel letters (A, B, ...) so templates follow position, not header names.
            for (var i = 0; i < _sourcePicks.Count; i++)
            {
                if (!_sourcePicks[i].Include)
                    continue;
                int at = _source.Columns.IndexOf(_sourcePicks[i].Name);
                if (at < 0)
                {
                    // Fallback: pick order matches source order.
                    at = i;
                }
                list.Add(TemplatesWork.ColumnLetter(at));
            }
            return list;
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            DataTable output;
            if (!TryBuildOutput(out output))
                return;
            GridOut.ItemsSource = Capped(output, _previewRows).DefaultView;
        }

        void ApplyTemplateDefinition(TemplateDefinition template)
        {
            TemplateName.Text = template.Name;
            KeepSource.IsChecked = template.KeepSourceColumns;
            _columns = new List<TemplateColumn>(template.Columns ?? new List<TemplateColumn>());
            RebuildColumns();
            RebuildSourcePicks(template.SourceColumns);
        }

        void UseInPipeline_Click(object sender, RoutedEventArgs e)
        {
            DataTable output;
            if (!TryBuildOutput(out output))
                return;
            TemplateDefinition template = BuildTemplate();
            // All source columns ticked: store none, which means "all", so columns
            // that an earlier step adds later still come through.
            if (template.KeepSourceColumns && _sourcePicks != null && template.SourceColumns.Count == _sourcePicks.Count)
                template.SourceColumns = new List<string>();
            _use(template);
        }

        bool TryBuildOutput(out DataTable output)
        {
            output = null;
            if (_source == null)
            {
                Alert("Choose a data file first.", "Missing file", MessageBoxImage.Warning);
                return false;
            }
            TemplateDefinition template = BuildTemplate();
            string error = TemplatesWork.Validate(template, true);
            if (error != null)
            {
                Alert(error, "Add columns", MessageBoxImage.Warning);
                return false;
            }
            try
            {
                output = TemplatesWork.ApplyTemplate(FullSource(), template);
                return true;
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not build preview", MessageBoxImage.Error);
                return false;
            }
        }

        TemplateDefinition BuildTemplate()
        {
            var template = new TemplateDefinition();
            template.Name = TemplateName.Text != null ? TemplateName.Text.Trim() : "";
            template.KeepSourceColumns = KeepSource.IsChecked != false;
            template.SourceColumns = SelectedSourceColumns();
            template.Columns = new List<TemplateColumn>(_columns);
            return template;
        }

        void AddUpdate_Click(object sender, RoutedEventArgs e)
        {
            EditColumn(-1);
        }

        void RebuildColumns()
        {
            ColumnPanel.Children.Clear();
            if (_columns.Count == 0)
            {
                ColumnPanel.Children.Add(new TextBlock { Text = "No added columns yet.", Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 0, 4) });
                return;
            }
            for (var i = 0; i < _columns.Count; i++)
            {
                var index = i;
                var text = (i + 1).ToString(CultureInfo.InvariantCulture) + ". " + ColumnText(_columns[i]);
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.Click += (s, e) => EditColumn(index);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveColumn(index);
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
                ColumnPanel.Children.Add(row);
            }
        }

        // "Price [Math]: $A * $B" reads better than the stored kind number.
        static string ColumnText(TemplateColumn col)
        {
            switch (col.Kind)
            {
                case TemplateColumnKind.Copy: return "Copy " + col.SourceColumn;
                case TemplateColumnKind.Concat: return "Text " + col.Pattern;
                case TemplateColumnKind.Math: return "Math " + col.Expression;
                case TemplateColumnKind.Conditional: return "If " + col.Condition + " -> " + col.TrueExpression + " else " + col.FalseExpression;
                default: return "Fixed " + col.FixedValue;
            }
        }

        // index -1 adds a new column.
        void EditColumn(int index)
        {
            var seed = index >= 0 ? _columns[index] : null;
            var col = TemplateColumnWindow.Edit(Window.GetWindow(this), _pipeInput, seed, _columns, index);
            if (col == null)
                return;
            if (index >= 0)
                _columns[index] = col;
            else
                _columns.Add(col);
            RebuildColumns();
            // The pipeline applies this editor by itself after a click inside it; the dialog's clicks never reach it.
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void RemoveColumn(int index)
        {
            _columns.RemoveAt(index);
            RebuildColumns();
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }
        public sealed class SourceColumnPick
        {
            public string Name { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            Notices.Show(this, message, title, MessageBoxButton.OK, icon);
        }
    }
}
