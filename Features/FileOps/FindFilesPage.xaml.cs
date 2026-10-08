using System;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Forms = System.Windows.Forms;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FileOps
{
    // Pipeline-only: edits a Find files step (keys from the input) or a List folder step (no input).
    public partial class FindFilesPage : Page
    {
        const int PreviewRows = 1000;

        readonly DataTable _keys;
        readonly Action<FindFilesSettings> _useFind;
        readonly Action<FolderScanSettings> _useList;

        // Find files: keys is the input step's result.
        public FindFilesPage(DataTable keys, string label, FindFilesSettings settings, Action<FindFilesSettings> use)
        {
            InitializeComponent();
            _keys = keys;
            _useFind = use;
            var keyCol = keys.Columns.Count == 0 ? null : ExcelFile.FindColumn(keys, "Key") ?? keys.Columns[0];
            InputInfo.Text = label + ": " + keys.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows. Keys come from column " +
                (keyCol == null ? "(none)" : "'" + ExcelFile.Header(keyCol) + "'") + (ExcelFile.FindColumn(keys, "Value") != null ? ", values from 'Value'." : ".");
            var s = settings ?? new FindFilesSettings();
            Fill(s.Scan);
            MatchBox.SelectedIndex = (int)s.Match;
            FileFormulaBox.Text = s.FileFormula ?? "";
            FolderFormulaBox.Text = s.FolderFormula ?? "";
            UnmatchedBox.IsChecked = s.KeepUnmatched;
            InsideBox.IsChecked = s.SkipInsideMatchedFolders;
            Hook();
        }

        // List folder: no input.
        public FindFilesPage(FolderScanSettings settings, Action<FolderScanSettings> use)
        {
            InitializeComponent();
            _useList = use;
            HeaderTitle.Text = "List folder";
            Title = "List folder";
            InputInfo.Text = "Lists every file and/or folder that passes the filter. Connect it to Compare, Filter & Sort or File action.";
            MatchPanel.Visibility = Visibility.Collapsed;
            Fill(settings ?? new FolderScanSettings());
            Hook();
        }

        void Hook()
        {
            ExcelGrid.Hook(ResultGrid);
            FormulaField.WatchOptional(FileFilterBox);
            FormulaField.WatchOptional(FolderFilterBox);
            FormulaField.Watch(FileFormulaBox, NeedFileFormula);
            FormulaField.Watch(FolderFormulaBox, NeedFolderFormula);
            Match_Changed(null, null);
        }

        void Fill(FolderScanSettings s)
        {
            FolderBox.Text = s.Folder ?? "";
            IncludeBox.SelectedIndex = (int)s.Include;
            RecurseBox.IsChecked = s.Recursive;
            SepBox.Text = string.IsNullOrEmpty(s.Separator) ? "_" : s.Separator;
            FileFilterBox.Text = s.FileFilter ?? "";
            FolderFilterBox.Text = s.FolderFilter ?? "";
        }

        bool LooksForFiles()
        {
            return IncludeBox.SelectedIndex != (int)FileInclude.Folders;
        }

        bool LooksForFolders()
        {
            return IncludeBox.SelectedIndex != (int)FileInclude.Files;
        }

        bool FormulaMode()
        {
            return _useFind != null && MatchBox.SelectedIndex == (int)FileMatchMode.Formula;
        }

        bool NeedFileFormula()
        {
            return FormulaMode() && LooksForFiles();
        }

        bool NeedFolderFormula()
        {
            return FormulaMode() && LooksForFolders();
        }

        void Include_Changed(object sender, SelectionChangedEventArgs e)
        {
            Match_Changed(null, null);
        }

        // Only the boxes for the kinds of item being looked for are in use.
        void Match_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (FileFormulaBox == null || FolderFormulaBox == null || FileFilterBox == null || FolderFilterBox == null)
                return;
            FileFilterBox.IsEnabled = LooksForFiles();
            FolderFilterBox.IsEnabled = LooksForFolders();
            FileFormulaBox.IsEnabled = NeedFileFormula();
            FolderFormulaBox.IsEnabled = NeedFolderFormula();
            FormulaField.Refresh(FileFormulaBox);
            FormulaField.Refresh(FolderFormulaBox);
        }

        FolderScanSettings CurrentScan()
        {
            return new FolderScanSettings
            {
                Folder = FolderBox.Text.Trim(),
                Include = (FileInclude)Math.Max(0, IncludeBox.SelectedIndex),
                Recursive = RecurseBox.IsChecked == true,
                Separator = SepBox.Text.Length > 0 ? SepBox.Text.Substring(0, 1) : "_",
                FileFilter = FileFilterBox.Text.Trim(),
                FolderFilter = FolderFilterBox.Text.Trim()
            };
        }

        FindFilesSettings CurrentFind()
        {
            return new FindFilesSettings
            {
                Scan = CurrentScan(),
                Match = (FileMatchMode)Math.Max(0, MatchBox.SelectedIndex),
                FileFormula = FileFormulaBox.Text.Trim(),
                FolderFormula = FolderFormulaBox.Text.Trim(),
                KeepUnmatched = UnmatchedBox.IsChecked == true,
                SkipInsideMatchedFolders = InsideBox.IsChecked == true
            };
        }

        string Check()
        {
            return _useFind != null ? FileOpsWork.Check(CurrentFind()) : FileOpsWork.Check(CurrentScan());
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            var error = Check();
            if (error != null)
            {
                ResultGrid.ItemsSource = null;
                ResultInfo.Text = error;
                return;
            }
            DataTable table;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                table = _useFind != null ? FileOpsWork.FindFiles(_keys, CurrentFind()) : FileOpsWork.ListFolder(CurrentScan());
            }
            catch (Exception ex)
            {
                ResultGrid.ItemsSource = null;
                ResultInfo.Text = ex.Message;
                return;
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
            var shown = table;
            if (table.Rows.Count > PreviewRows)
            {
                shown = table.Clone();
                for (var i = 0; i < PreviewRows; i++)
                    shown.ImportRow(table.Rows[i]);
            }
            ResultGrid.ItemsSource = shown.DefaultView;
            ResultInfo.Text = table.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows" +
                (shown != table ? " (showing the first " + PreviewRows.ToString("N0", CultureInfo.InvariantCulture) + ")" : "") +
                ". Nothing is changed on disk.";
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "File formulas see each file, folder formulas see each folder. $1 $2 split the name without extension by the split character.",
                FileOpsWork.FileVariableHelp(_useFind != null, false));
        }

        void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = "Folder to search";
                dlg.ShowNewFolderButton = false;
                if (FolderBox.Text.Trim().Length > 0)
                    dlg.SelectedPath = FolderBox.Text.Trim();
                if (dlg.ShowDialog() == Forms.DialogResult.OK)
                    FolderBox.Text = dlg.SelectedPath;
            }
        }

        void Use_Click(object sender, RoutedEventArgs e)
        {
            var error = Check();
            if (error != null)
            {
                Alert(error, "Not ready", MessageBoxImage.Warning);
                return;
            }
            if (_useFind != null)
                _useFind(CurrentFind());
            else
                _useList(CurrentScan());
            Back_Click(null, null);
        }

        void Back_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }
    }
}
