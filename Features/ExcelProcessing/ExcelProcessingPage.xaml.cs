using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.ExcelProcessing
{
    public partial class ExcelProcessingPage : UserControl
    {
        DataTable _tableA;
        DataTable _tableB;
        int _previewRows = 1000;
        ExcelProcessingCompareResult _preview;
        List<ColumnPick> _picks;
        readonly List<ExcelProcessingMatch> _rules = new List<ExcelProcessingMatch>();
        bool _loading;
        bool _syncPicks;
        bool _restoring;
        // Inputs come from earlier pipeline steps, and "Use in pipeline"
        // hands the settings back.
        DataTable _pipeA;
        DataTable _pipeB;
        string _pipeLabelA;
        string _pipeLabelB;
        Action<ExcelProcessingSettings> _use;

        public ExcelProcessingPage(DataTable a, string labelA, DataTable b, string labelB,
            ExcelProcessingSettings settings, Action<ExcelProcessingSettings> use)
        {
            InitializeComponent();
            _pipeA = a;
            _pipeB = b;
            _pipeLabelA = labelA;
            _pipeLabelB = labelB;
            _use = use;
            PreviewBox.Items.Add(200);
            PreviewBox.Items.Add(1000);
            PreviewBox.Items.Add(5000);
            PreviewBox.SelectedItem = 1000;
            RebuildRules();
            ExcelGrid.Hook(GridA, GridB, GridOut);
            InjectInputs();
            if (settings != null)
                ApplySettings(settings);
        }

        void InjectInputs()
        {
            ApplyLoad(true, _pipeA, _pipeLabelA);
            ApplyLoad(false, _pipeB, _pipeLabelB);
        }

        void PreviewBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewBox.SelectedItem == null)
                return;
            var rows = (int)PreviewBox.SelectedItem;
            if (rows == _previewRows)
                return;
            _previewRows = rows;
            InjectInputs();
        }

        void AddCond_Click(object sender, RoutedEventArgs e)
        {
            EditRule(-1);
        }

        void Op_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshPreview(true);
        }

        void UseInPipeline_Click(object sender, RoutedEventArgs e)
        {
            if (_preview == null)
            {
                // Not built yet: rebuild with alerts so the user sees what is missing.
                if (SelectedOp() == null)
                    Alert("Choose an output operation first.", "Not ready", MessageBoxImage.Warning);
                else
                    RefreshPreview(true);
                return;
            }
            var selected = SelectedPicks();
            if (selected.Count == 0)
            {
                Alert("Select at least one column.", "No columns", MessageBoxImage.Warning);
                return;
            }

            var s = new ExcelProcessingSettings();
            s.Operation = SelectedOp();
            s.Matches = CollectedConds();
            s.AllColumns = BothState() == true;
            if (!s.AllColumns)
            {
                for (var i = 0; i < selected.Count; i++)
                    s.Columns.Add(new ExcelProcessingColumn { Side = selected[i].Source.ToString(), Header = selected[i].OriginalName });
            }
            _use(s);
        }

        void ApplySettings(ExcelProcessingSettings s)
        {
            _restoring = true;
            try
            {
                _rules.Clear();
                if (s.Matches != null)
                {
                    foreach (var m in s.Matches)
                        _rules.Add(CopyRule(m));
                }
                RebuildRules();
                foreach (var item in OpList.Items)
                {
                    var op = item as ListBoxItem;
                    if (op != null && (op.Tag as string) == s.Operation)
                        OpList.SelectedItem = op;
                }
            }
            finally
            {
                _restoring = false;
            }
            RefreshPreview(false);

            if (s.AllColumns || _picks == null || s.Columns == null)
                return;
            var used = new HashSet<ColumnPick>();
            foreach (var p in _picks)
                p.Include = false;
            foreach (var want in s.Columns)
            {
                foreach (var p in _picks)
                {
                    if (!used.Contains(p) && p.Source.ToString() == want.Side &&
                        string.Equals(p.OriginalName, want.Header, StringComparison.OrdinalIgnoreCase))
                    {
                        p.Include = true;
                        used.Add(p);
                        break;
                    }
                }
            }
            AfterUtility();
        }

        // The grids show a capped preview of the pipeline input.
        void ApplyLoad(bool isA, DataTable full, string label)
        {
            var preview = full;
            if (full.Rows.Count > _previewRows)
            {
                preview = full.Clone();
                for (var i = 0; i < _previewRows; i++)
                    preview.ImportRow(full.Rows[i]);
            }
            var info = RowInfo(preview.Rows.Count, full.Rows.Count, full.Rows.Count > _previewRows);

            _loading = true;
            try
            {
                if (isA)
                {
                    _tableA = preview;
                    PathA.Text = label;
                    GridA.ItemsSource = preview.DefaultView;
                    InfoA.Text = info;
                }
                else
                {
                    _tableB = preview;
                    PathB.Text = label;
                    GridB.ItemsSource = preview.DefaultView;
                    InfoB.Text = info;
                }
            }
            finally
            {
                _loading = false;
            }

            RefreshPreview(false);
        }

        static string RowInfo(int shown, int total, bool cut)
        {
            var rows = shown.ToString("N0", CultureInfo.InvariantCulture);
            if (!cut)
                return rows + " rows";
            return "Showing " + rows + " of " + total.ToString("N0", CultureInfo.InvariantCulture) +
                   " rows (all rows are used when you run an operation)";
        }

        // Operations need every row, not just the capped preview.
        DataTable FullTable(bool isA)
        {
            return isA ? _pipeA : _pipeB;
        }

        internal static string[] ColumnNames(DataTable table)
        {
            var names = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                names[i] = ExcelFile.Header(table.Columns[i]);
            return names;
        }

        void RefreshPreview(bool alert)
        {
            if (_loading || _restoring)
                return;

            var tag = SelectedOp();
            if (tag == null)
            {
                ClearPreview();
                return;
            }

            DataTable a;
            DataTable b;
            if (!TryReady(out a, out b, alert))
            {
                ClearPreview();
                return;
            }

            if (!FormulasOk())
                return;

            try
            {
                _preview = ExcelProcessingWork.Compare(a, b, tag, CollectedConds());

                _picks = new List<ColumnPick>();
                for (var i = 0; i < _preview.Table.Columns.Count; i++)
                {
                    _picks.Add(new ColumnPick
                    {
                        InternalName = _preview.Table.Columns[i].ColumnName,
                        OriginalName = _preview.Originals[i],
                        Source = _preview.Sources[i],
                        Include = true
                    });
                }
                ApplyPickLabels();

                ColList.ItemsSource = _picks;
                ShowProjected();
                SyncUtility();
            }
            catch (Exception ex)
            {
                ClearPreview();
                if (alert)
                    Alert(ex.Message, "Could not build preview", MessageBoxImage.Error);
            }
        }

        bool FormulasOk()
        {
            for (var i = 0; i < _rules.Count; i++)
            {
                if (ExpressionEngine.Validate(_rules[i].ExpressionA) != null || ExpressionEngine.Validate(_rules[i].ExpressionB) != null)
                    return false;
            }
            return true;
        }

        void ClearPreview()
        {
            _preview = null;
            _picks = null;
            ColList.ItemsSource = null;
            GridOut.ItemsSource = null;
            SyncUtility();
        }

        string SelectedOp()
        {
            var item = OpList.SelectedItem as ListBoxItem;
            return item != null ? item.Tag as string : null;
        }

        List<ColumnPick> SelectedPicks()
        {
            var list = new List<ColumnPick>();
            if (_picks == null)
                return list;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Include)
                    list.Add(_picks[i]);
            }
            return list;
        }

        // _picks lines up with the preview's columns, so a pick's index is its column.
        DataTable Projected(List<ColumnPick> selected, int maxRows)
        {
            var columns = new List<int>();
            for (var i = 0; i < selected.Count; i++)
                columns.Add(_picks.IndexOf(selected[i]));
            return ExcelProcessingWork.ProjectColumns(_preview, columns, maxRows);
        }

        void ApplyPickLabels()
        {
            if (_picks == null)
                return;
            var originals = new string[_picks.Count];
            var sources = new char[_picks.Count];
            for (var i = 0; i < _picks.Count; i++)
            {
                originals[i] = _picks[i].OriginalName;
                sources[i] = _picks[i].Source;
            }
            for (var i = 0; i < _picks.Count; i++)
                _picks[i].Label = ExcelProcessingWork.OutputName(originals[i], sources[i], originals, sources);
        }

        void ShowProjected()
        {
            if (_preview == null)
            {
                GridOut.ItemsSource = null;
                return;
            }
            GridOut.ItemsSource = Projected(SelectedPicks(), _previewRows).DefaultView;
        }

        void Col_Changed(object sender, RoutedEventArgs e)
        {
            if (_syncPicks)
                return;
            ShowProjected();
            SyncUtility();
        }

        void PickA_Click(object sender, RoutedEventArgs e)
        {
            ToggleSource('A');
        }

        void PickB_Click(object sender, RoutedEventArgs e)
        {
            ToggleSource('B');
        }

        void PickBoth_Click(object sender, RoutedEventArgs e)
        {
            if (_syncPicks)
                return;
            var on = BothState() != true;
            SetSource('A', on);
            SetSource('B', on);
            AfterUtility();
        }

        void ToggleSource(char source)
        {
            if (_syncPicks)
                return;
            SetSource(source, SourceState(source) != true);
            AfterUtility();
        }

        void SetSource(char source, bool include)
        {
            if (_picks == null)
                return;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Source == source)
                    _picks[i].Include = include;
            }
        }

        void AfterUtility()
        {
            ColList.ItemsSource = null;
            ColList.ItemsSource = _picks;
            ShowProjected();
            SyncUtility();
        }

        void SyncUtility()
        {
            _syncPicks = true;
            try
            {
                PickA.IsChecked = SourceState('A');
                PickB.IsChecked = SourceState('B');
                PickBoth.IsChecked = BothState();
                PickA.IsEnabled = HasSource('A');
                PickB.IsEnabled = HasSource('B');
                PickBoth.IsEnabled = _picks != null && _picks.Count > 0;
            }
            finally
            {
                _syncPicks = false;
            }
        }

        bool HasSource(char source)
        {
            if (_picks == null)
                return false;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Source == source)
                    return true;
            }
            return false;
        }

        bool? SourceState(char source)
        {
            if (_picks == null)
                return false;
            var any = false;
            var all = true;
            var none = true;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Source != source)
                    continue;
                any = true;
                if (_picks[i].Include)
                    none = false;
                else
                    all = false;
            }
            if (!any)
                return false;
            if (all)
                return true;
            if (none)
                return false;
            return null;
        }

        bool? BothState()
        {
            if (_picks == null || _picks.Count == 0)
                return false;
            var all = true;
            var none = true;
            for (var i = 0; i < _picks.Count; i++)
            {
                if (_picks[i].Include)
                    none = false;
                else
                    all = false;
            }
            if (all)
                return true;
            if (none)
                return false;
            return null;
        }

        void RebuildRules()
        {
            CondPanel.Children.Clear();
            if (_rules.Count == 0)
            {
                CondPanel.Children.Add(new TextBlock
                {
                    Text = "No match rules yet.",
                    Foreground = Brushes.Gray,
                    Margin = new Thickness(0, 0, 0, 4)
                });
                return;
            }
            for (var i = 0; i < _rules.Count; i++)
            {
                var index = i;
                var text = RuleText(_rules[i]);
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.Click += (s, e) => EditRule(index);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveRule(index);
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
                CondPanel.Children.Add(row);
            }
        }

        static ExcelProcessingMatch CopyRule(ExcelProcessingMatch m)
        {
            return new ExcelProcessingMatch
            {
                ColumnA = m.ColumnA,
                ColumnB = m.ColumnB,
                ExpressionA = m.ExpressionA,
                ExpressionB = m.ExpressionB,
                Operator = m.Operator,
                SplitA = m.SplitA,
                SplitB = m.SplitB
            };
        }

        static string RuleText(ExcelProcessingMatch m)
        {
            return Side(m.ColumnA, m.SplitA, m.ExpressionA) + "   " + (m.Operator ?? "==") + "   " +
                Side(m.ColumnB, m.SplitB, m.ExpressionB);
        }

        static string Side(string column, string split, string formula)
        {
            var head = column ?? "?";
            if (!string.IsNullOrEmpty(split))
                head += " (split " + SplitName(split) + ")";
            return head + ": " + formula;
        }

        static string SplitName(string split)
        {
            switch (split)
            {
                case " ": return "space";
                case "_": return "underscore";
                case "-": return "dash";
                default: return split;
            }
        }

        // index -1 adds a new rule.
        void EditRule(int index)
        {
            if (_tableA == null || _tableB == null)
            {
                Alert("Open both Excel files first.", "Missing file", MessageBoxImage.Warning);
                return;
            }
            var seed = index >= 0 ? _rules[index] : null;
            var rule = MatchRuleWindow.Edit(Window.GetWindow(this), _tableA, _tableB, seed);
            if (rule == null)
                return;
            if (index >= 0)
                _rules[index] = rule;
            else
                _rules.Add(rule);
            RulesChanged();
        }

        void RemoveRule(int index)
        {
            _rules.RemoveAt(index);
            RulesChanged();
        }

        // The pipeline applies this editor by itself after a click inside it; the dialog's clicks never reach it.
        void RulesChanged()
        {
            RebuildRules();
            RefreshPreview(false);
            RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        }

        List<ExcelProcessingMatch> CollectedConds()
        {
            return new List<ExcelProcessingMatch>(_rules);
        }
        static bool HasJoin(IList<ExcelProcessingMatch> extra)
        {
            if (extra == null)
                return false;
            for (var i = 0; i < extra.Count; i++)
            {
                var m = extra[i];
                if (!string.IsNullOrWhiteSpace(m.ExpressionA) && !string.IsNullOrWhiteSpace(m.ExpressionB))
                    return true;
            }
            return false;
        }

        bool TryReady(out DataTable a, out DataTable b, bool alert)
        {
            a = null;
            b = null;

            if (_tableA == null || _tableB == null)
            {
                if (alert)
                    Alert("Open both Excel files first.", "Missing file", MessageBoxImage.Warning);
                return false;
            }
            if (!HasJoin(CollectedConds()))
            {
                if (alert)
                    Alert("Set a match rule that uses both A and B.", "Missing match", MessageBoxImage.Warning);
                return false;
            }

            try
            {
                a = FullTable(true);
                b = FullTable(false);
                return true;
            }
            catch (Exception ex)
            {
                if (alert)
                    Alert(ex.Message, "Could not read all rows", MessageBoxImage.Error);
                return false;
            }
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            Notices.Show(this, message, title, MessageBoxButton.OK, icon);
        }

        public sealed class ColumnPick
        {
            public string InternalName { get; set; }
            public string OriginalName { get; set; }
            public char Source { get; set; }
            public bool Include { get; set; }
            public string Label { get; set; }
        }
    }
}
