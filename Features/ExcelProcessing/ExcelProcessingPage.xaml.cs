using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
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
        readonly List<CondRow> _conds = new List<CondRow>();
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
            AddCond(null);
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
            AddCond(null);
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
                _conds.Clear();
                CondPanel.Children.Clear();
                if (s.Matches == null || s.Matches.Count == 0)
                    AddCond(null);
                else
                {
                    foreach (var m in s.Matches)
                        AddCond(m);
                }
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
                RefreshCondCombos();
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

        static string[] ColumnNames(DataTable table)
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
            for (var i = 0; i < _conds.Count; i++)
            {
                if (!FormulaField.Ok(_conds[i].ExprA) || !FormulaField.Ok(_conds[i].ExprB))
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

        void AddCond(ExcelProcessingMatch seed)
        {
            var root = new StackPanel { Margin = new Thickness(0, 0, 8, 12) };
            var line = new Grid();
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(48) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            line.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var panelA = new StackPanel { Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Bottom };
            var opBox = new ComboBox { Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Bottom };
            opBox.Items.Add("==");
            opBox.Items.Add("!=");
            opBox.Items.Add(">");
            opBox.Items.Add(">=");
            opBox.Items.Add("<");
            opBox.Items.Add("<=");
            var opKeep = seed != null && !string.IsNullOrWhiteSpace(seed.Operator) ? seed.Operator : "==";
            opBox.SelectedItem = opKeep;
            if (opBox.SelectedIndex < 0)
                opBox.SelectedIndex = 0;

            var panelB = new StackPanel { Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Bottom };
            var del = new Button { Content = "Remove", Width = 64, VerticalAlignment = VerticalAlignment.Bottom };

            Grid.SetColumn(opBox, 1);
            Grid.SetColumn(panelB, 2);
            Grid.SetColumn(del, 3);
            line.Children.Add(panelA);
            line.Children.Add(opBox);
            line.Children.Add(panelB);
            line.Children.Add(del);
            root.Children.Add(line);

            var item = new CondRow
            {
                Root = root,
                Op = opBox,
                PanelA = panelA,
                PanelB = panelB,
                ColA = new ComboBox { MinWidth = 120, Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Bottom, ToolTip = "Column for $Value and Split" },
                ColB = new ComboBox { MinWidth = 120, Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Bottom, ToolTip = "Column for $Value and Split" },
                SepA = NewSepBox(),
                SepB = NewSepBox(),
                CustomA = NewCustomSep(),
                CustomB = NewCustomSep(),
                ExprA = new TextBox { Margin = new Thickness(0, 2, 0, 0), Text = seed != null && seed.ExpressionA != null ? seed.ExpressionA : "$Value", ToolTip = "Formula. $Value is the column." },
                ExprB = new TextBox { Margin = new Thickness(0, 2, 0, 0), Text = seed != null && seed.ExpressionB != null ? seed.ExpressionB : "$Value", ToolTip = "Formula. $Value is the column." }
            };
            if (seed != null)
            {
                ApplySep(item.SepA, item.CustomA, seed.SplitA);
                ApplySep(item.SepB, item.CustomB, seed.SplitB);
            }

            FillCondCombo(item.ColA, _tableA, seed != null ? seed.ColumnA : null);
            FillCondCombo(item.ColB, _tableB, seed != null ? seed.ColumnB : null);
            FormulaField.Watch(item.ExprA);
            FormulaField.Watch(item.ExprB);
            WireCond(item.ColA, item.ColB, item.Op, item.SepA, item.SepB, item.CustomA, item.CustomB,
                item.ExprA, item.ExprB);
            item.SepA.SelectionChanged += (s, e) => item.CustomA.IsEnabled = item.SepA.SelectedIndex == 4;
            item.SepB.SelectionChanged += (s, e) => item.CustomB.IsEnabled = item.SepB.SelectedIndex == 4;

            del.Click += (s, e) =>
            {
                _conds.Remove(item);
                CondPanel.Children.Remove(root);
                if (_conds.Count == 0)
                    AddCond(null);
                else
                    RefreshPreview(false);
            };

            FillSidePanels(item);
            _conds.Add(item);
            CondPanel.Children.Add(root);
            RefreshPreview(false);
        }

        void WireCond(params Control[] controls)
        {
            for (var i = 0; i < controls.Length; i++)
            {
                var box = controls[i] as ComboBox;
                if (box != null)
                    box.SelectionChanged += Cond_Changed;
                var text = controls[i] as TextBox;
                if (text != null)
                    text.TextChanged += Cond_TextChanged;
            }
        }

        static ComboBox NewSepBox()
        {
            var sep = new ComboBox { Width = 92, Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Bottom, ToolTip = "Split $Value into $1 $2" };
            sep.Items.Add("None");
            sep.Items.Add("Space");
            sep.Items.Add("Underscore");
            sep.Items.Add("Dash");
            sep.Items.Add("Custom");
            sep.SelectedIndex = 0;
            return sep;
        }

        static TextBox NewCustomSep()
        {
            return new TextBox { Width = 28, MaxLength = 1, IsEnabled = false, VerticalAlignment = VerticalAlignment.Bottom };
        }

        static void ApplySep(ComboBox sep, TextBox custom, string stored)
        {
            if (stored == " ")
                sep.SelectedIndex = 1;
            else if (stored == "_")
                sep.SelectedIndex = 2;
            else if (stored == "-")
                sep.SelectedIndex = 3;
            else if (!string.IsNullOrEmpty(stored))
            {
                sep.SelectedIndex = 4;
                custom.Text = stored.Substring(0, 1);
                custom.IsEnabled = true;
            }
            else
                sep.SelectedIndex = 0;
        }

        static string ReadSep(ComboBox sep, TextBox custom)
        {
            switch (sep.SelectedIndex)
            {
                case 1: return " ";
                case 2: return "_";
                case 3: return "-";
                case 4: return custom.Text.Length > 0 ? custom.Text.Substring(0, 1) : "";
                default: return "";
            }
        }

        static void FillSidePanels(CondRow item)
        {
            item.PanelA.Children.Clear();
            item.PanelB.Children.Clear();
            item.PanelA.Children.Add(SideRow(item.ColA, item.SepA, item.CustomA));
            item.PanelA.Children.Add(item.ExprA);
            item.PanelB.Children.Add(SideRow(item.ColB, item.SepB, item.CustomB));
            item.PanelB.Children.Add(item.ExprB);
        }

        static StackPanel SideRow(ComboBox col, ComboBox sep, TextBox custom)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Bottom };
            row.Children.Add(col);
            row.Children.Add(sep);
            row.Children.Add(custom);
            return row;
        }

        void FillCondCombo(ComboBox box, DataTable table, string keep)
        {
            var names = table == null ? new string[0] : ColumnNames(table);
            box.ItemsSource = names;
            // keep is a header (Caption), so look it up in names, not in Columns (ColumnName ids).
            var at = -1;
            for (var i = 0; keep != null && i < names.Length && at < 0; i++)
            {
                if (string.Equals(names[i], keep, StringComparison.OrdinalIgnoreCase))
                    at = i;
            }
            if (at >= 0)
                box.SelectedIndex = at;
            else if (names.Length > 0)
                box.SelectedIndex = 0;
            else
                box.SelectedIndex = -1;
        }

        void RefreshCondCombos()
        {
            for (var i = 0; i < _conds.Count; i++)
            {
                var c = _conds[i];
                FillCondCombo(c.ColA, _tableA, c.ColA.SelectedItem as string);
                FillCondCombo(c.ColB, _tableB, c.ColB.SelectedItem as string);
            }
        }

        void Cond_Changed(object sender, SelectionChangedEventArgs e)
        {
            RefreshPreview(false);
        }

        void Cond_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPreview(false);
        }

        List<ExcelProcessingMatch> CollectedConds()
        {
            var list = new List<ExcelProcessingMatch>();
            for (var i = 0; i < _conds.Count; i++)
            {
                var c = _conds[i];
                list.Add(new ExcelProcessingMatch
                {
                    ColumnA = c.ColA.SelectedItem as string,
                    ColumnB = c.ColB.SelectedItem as string,
                    ExpressionA = c.ExprA.Text,
                    ExpressionB = c.ExprB.Text,
                    Operator = c.Op.SelectedItem as string,
                    SplitA = ReadSep(c.SepA, c.CustomA),
                    SplitB = ReadSep(c.SepB, c.CustomB)
                });
            }
            return list;
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

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "Type $Today in a formula. $Value is the chosen column on that side. After Split, $1 $2 are parts of it. IF(), FIRSTWORD() and LASTWORD() work inside the formula.");
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }

        public sealed class ColumnPick
        {
            public string InternalName { get; set; }
            public string OriginalName { get; set; }
            public char Source { get; set; }
            public bool Include { get; set; }
            public string Label { get; set; }
        }

        sealed class CondRow
        {
            public StackPanel Root;
            public ComboBox Op;
            public StackPanel PanelA;
            public StackPanel PanelB;
            public ComboBox ColA;
            public ComboBox ColB;
            public ComboBox SepA;
            public ComboBox SepB;
            public TextBox CustomA;
            public TextBox CustomB;
            public TextBox ExprA;
            public TextBox ExprB;
        }
    }
}
