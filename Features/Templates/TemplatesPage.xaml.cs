using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.Templates
{
    public partial class TemplatesPage : Page
    {
        DataTable _source;
        int _previewRows = 1000;
        List<TemplateColumn> _columns = new List<TemplateColumn>();
        List<SourceColumnPick> _sourcePicks;
        int _editingIndex = -1;
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
            ColKind.SelectedIndex = 0;
            FormulaField.Watch(MathExpression, MathRequired);
            FormulaField.Watch(CondCondition, CondRequired);
            FormulaField.WatchOptional(CondTrue);
            FormulaField.WatchOptional(CondFalse);
            RefreshColumnList();
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
            RefreshSourceCombos();
        }

        void Back_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
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

        void RefreshSourceCombos()
        {
            string[] labels = TemplatesWork.SourceLabels(_source);
            CopySource.ItemsSource = labels;
            if (CopySource.Items.Count > 0 && CopySource.SelectedIndex < 0)
                CopySource.SelectedIndex = 0;
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

        void Kind_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PanelFixed == null || PanelCopy == null || PanelConcat == null || PanelMath == null || PanelConditional == null || ColKind == null)
                return;
            int kind = ColKind.SelectedIndex;
            PanelFixed.Visibility = kind == 0 ? Visibility.Visible : Visibility.Collapsed;
            PanelCopy.Visibility = kind == 1 ? Visibility.Visible : Visibility.Collapsed;
            PanelConcat.Visibility = kind == 2 ? Visibility.Visible : Visibility.Collapsed;
            PanelMath.Visibility = kind == 3 ? Visibility.Visible : Visibility.Collapsed;
            PanelConditional.Visibility = kind == 4 ? Visibility.Visible : Visibility.Collapsed;
            FormulaField.Refresh(MathExpression);
            FormulaField.Refresh(CondCondition);
            FormulaField.Refresh(CondTrue);
            FormulaField.Refresh(CondFalse);
        }

        bool MathRequired()
        {
            return ColKind.SelectedIndex == 3;
        }

        bool CondRequired()
        {
            return ColKind.SelectedIndex == 4;
        }

        void Columns_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = ColumnList.SelectedIndex;
            if (index < 0 || index >= _columns.Count)
                return;
            _editingIndex = index;
            LoadEditor(_columns[index]);
            AddUpdateBtn.Content = "Update column";
        }

        void AddUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (ColKind.SelectedIndex == 3 && !FormulaField.Check(MathExpression))
                return;
            if (ColKind.SelectedIndex == 4 &&
                (!FormulaField.Check(CondCondition) || !FormulaField.Check(CondTrue) || !FormulaField.Check(CondFalse)))
                return;

            TemplateColumn col;
            try
            {
                col = BuildColumnFromEditor();
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Column", MessageBoxImage.Warning);
                return;
            }

            string error = TemplatesWork.ValidateColumn(col);
            if (error != null)
            {
                Alert(error, "Column", MessageBoxImage.Warning);
                return;
            }

            if (_editingIndex >= 0 && _editingIndex < _columns.Count)
            {
                for (var i = 0; i < _columns.Count; i++)
                {
                    if (i != _editingIndex &&
                        string.Equals(_columns[i].Name.Trim(), col.Name.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        Alert("Duplicate column name: " + col.Name.Trim() + ".", "Column", MessageBoxImage.Warning);
                        return;
                    }
                }
                _columns[_editingIndex] = col;
            }
            else
            {
                for (var i = 0; i < _columns.Count; i++)
                {
                    if (string.Equals(_columns[i].Name.Trim(), col.Name.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        Alert("Duplicate column name: " + col.Name.Trim() + ".", "Column", MessageBoxImage.Warning);
                        return;
                    }
                }
                _columns.Add(col);
            }

            ClearEditor();
            RefreshColumnList();
        }

        void Remove_Click(object sender, RoutedEventArgs e)
        {
            int index = ColumnList.SelectedIndex;
            if (index < 0 || index >= _columns.Count)
            {
                Alert("Select a column first.", "Remove", MessageBoxImage.Warning);
                return;
            }
            _columns.RemoveAt(index);
            ClearEditor();
            RefreshColumnList();
        }

        void ClearEditor_Click(object sender, RoutedEventArgs e)
        {
            ClearEditor();
            RefreshColumnList();
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "Type the Name exactly, e.g. $Today. Works in Fixed, Combine, Math and Conditional. Week starts Monday. Dates have no time except $Now. Use $A $B for columns (A = 1st, B = 2nd). FIRSTWORD() and LASTWORD() pick the first/last word of a text.");
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            DataTable output;
            if (!TryBuildOutput(out output))
                return;
            GridOut.ItemsSource = Capped(output, _previewRows).DefaultView;
        }

        void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            TemplateDefinition template = BuildTemplate();
            string error = TemplatesWork.Validate(template, true);
            if (error != null)
            {
                Alert(error, "Template", MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Template files (*.xml)|*.xml",
                FileName = (string.IsNullOrWhiteSpace(template.Name) ? "template" : template.Name.Trim()) + ".xml"
            };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                TemplatesWork.SaveTemplate(template, dlg.FileName);
                Alert("Template saved.", "Done", MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not save template", MessageBoxImage.Error);
            }
        }

        void LoadTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Template files (*.xml)|*.xml",
                Title = "Load template"
            };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                ApplyTemplateDefinition(TemplatesWork.LoadTemplate(dlg.FileName));
                DataTable output;
                if (_source != null && TryBuildOutput(out output))
                    GridOut.ItemsSource = Capped(output, _previewRows).DefaultView;
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not load template", MessageBoxImage.Error);
            }
        }

        void ApplyTemplateDefinition(TemplateDefinition template)
        {
            TemplateName.Text = template.Name;
            KeepSource.IsChecked = template.KeepSourceColumns;
            _columns = new List<TemplateColumn>(template.Columns ?? new List<TemplateColumn>());
            ClearEditor();
            RefreshColumnList();
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
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
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
                Alert(error, "Template", MessageBoxImage.Warning);
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

        TemplateColumn BuildColumnFromEditor()
        {
            string name = ColName.Text != null ? ColName.Text.Trim() : "";
            if (name.Length == 0)
                throw new InvalidOperationException("Enter a column name.");

            var col = new TemplateColumn();
            col.Name = name;
            int kind = ColKind.SelectedIndex;
            if (kind == 1)
            {
                col.Kind = TemplateColumnKind.Copy;
                string picked = CopySource.SelectedItem as string;
                if (string.IsNullOrWhiteSpace(picked))
                    picked = CopySource.Text != null ? CopySource.Text.Trim() : "";
                // Store the Excel letter (A, B, ...) so templates use position, not header names.
                col.SourceColumn = TemplatesWork.ExtractLetter(picked);
            }
            else if (kind == 2)
            {
                col.Kind = TemplateColumnKind.Concat;
                col.Pattern = ConcatPattern.Text;
            }
            else if (kind == 3)
            {
                col.Kind = TemplateColumnKind.Math;
                col.Expression = MathExpression.Text;
                // Keep legacy fields filled so old readers still see something.
                col.Left = MathExpression.Text;
                col.Operator = "+";
                col.Right = "";
            }
            else if (kind == 4)
            {
                col.Kind = TemplateColumnKind.Conditional;
                col.Condition = CondCondition.Text;
                col.TrueExpression = CondTrue.Text;
                col.FalseExpression = CondFalse.Text;
            }
            else
            {
                col.Kind = TemplateColumnKind.Fixed;
                col.FixedValue = FixedValue.Text;
            }
            return col;
        }

        void LoadEditor(TemplateColumn col)
        {
            ColName.Text = col.Name;
            if (col.Kind == TemplateColumnKind.Copy)
            {
                ColKind.SelectedIndex = 1;
                CopySource.Text = "";
                CopySource.SelectedIndex = -1;
                string wantLetter = TemplatesWork.ExtractLetter(col.SourceColumn);
                for (var i = 0; i < CopySource.Items.Count; i++)
                {
                    string item = CopySource.Items[i] as string;
                    // Match by letter first (new templates store "A"), then by old header name.
                    if (string.Equals(TemplatesWork.ExtractLetter(item), wantLetter, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(item, col.SourceColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        CopySource.SelectedIndex = i;
                        break;
                    }
                }
                if (CopySource.SelectedIndex < 0)
                    CopySource.Text = col.SourceColumn;
            }
            else if (col.Kind == TemplateColumnKind.Concat)
            {
                ColKind.SelectedIndex = 2;
                ConcatPattern.Text = col.Pattern;
            }
            else if (col.Kind == TemplateColumnKind.Math)
            {
                ColKind.SelectedIndex = 3;
                string expr = col.Expression;
                if (string.IsNullOrWhiteSpace(expr) && !string.IsNullOrWhiteSpace(col.Left))
                    expr = col.Left + " " + col.Operator + " " + col.Right;
                MathExpression.Text = expr;
            }
            else if (col.Kind == TemplateColumnKind.Conditional)
            {
                ColKind.SelectedIndex = 4;
                CondCondition.Text = col.Condition;
                CondTrue.Text = col.TrueExpression;
                CondFalse.Text = col.FalseExpression;
            }
            else
            {
                ColKind.SelectedIndex = 0;
                FixedValue.Text = col.FixedValue;
            }
        }

        void ClearEditor()
        {
            _editingIndex = -1;
            ColumnList.SelectedIndex = -1;
            AddUpdateBtn.Content = "Add column";
            ColName.Text = "";
            FixedValue.Text = "";
            ConcatPattern.Text = "";
            MathExpression.Text = "";
            CondCondition.Text = "";
            CondTrue.Text = "";
            CondFalse.Text = "";
            if (CopySource.Items.Count > 0)
                CopySource.SelectedIndex = 0;
        }

        void RefreshColumnList()
        {
            ColumnList.ItemsSource = null;
            ColumnList.ItemsSource = _columns;
        }

        public sealed class SourceColumnPick
        {
            public string Name { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }
    }
}
