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
    // Pipeline-only: edits a File action step. Shows the plan; never touches the disk.
    public partial class FileActionPage : UserControl
    {
        const int PreviewRows = 1000;

        readonly DataTable _input;
        readonly Action<FileActionSettings> _use;
        bool _filling = true;

        public FileActionPage(DataTable input, string label, FileActionSettings settings, Action<FileActionSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
                        FormulaField.WatchOptional(FileSubBox);
            FormulaField.WatchOptional(FileNameBox);
            FormulaField.WatchOptional(FolderSubBox);
            FormulaField.WatchOptional(FolderNameBox);
            InputInfo.Text = label + ": " + input.Rows.Count.ToString("N0", CultureInfo.InvariantCulture) + " rows.";
            foreach (DataColumn col in input.Columns)
                PathColBox.Items.Add(ExcelFile.Header(col));

            var s = settings ?? new FileActionSettings();
            CopyRadio.IsChecked = s.Action == FileActionKind.Copy;
            MoveRadio.IsChecked = s.Action == FileActionKind.Move;
            DeleteRadio.IsChecked = s.Action == FileActionKind.Delete;
            PathColBox.Text = s.PathColumn ?? "Path";
            DestBox.Text = s.DestFolder ?? "";
            FileSubBox.Text = s.FileSubfolderFormula ?? "";
            FileNameBox.Text = s.FileNameFormula ?? "";
            FolderSubBox.Text = s.FolderSubfolderFormula ?? "";
            FolderNameBox.Text = s.FolderNameFormula ?? "";
            StructureBox.IsChecked = s.KeepStructure;
            ConflictBox.SelectedIndex = (int)s.Conflict;
            PermanentBox.IsChecked = s.Permanent;
            SepBox.Text = string.IsNullOrEmpty(s.Separator) ? "_" : s.Separator;
            _filling = false;
            Action_Changed(null, null);
        }

        void Action_Changed(object sender, RoutedEventArgs e)
        {
            if (_filling)
                return;
            var delete = DeleteRadio.IsChecked == true;
            DestPanel.Visibility = delete ? Visibility.Collapsed : Visibility.Visible;
            DeletePanel.Visibility = delete ? Visibility.Visible : Visibility.Collapsed;
            PermanentWarn.Visibility = PermanentBox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        }

        FileActionSettings Current()
        {
            return new FileActionSettings
            {
                Action = DeleteRadio.IsChecked == true ? FileActionKind.Delete
                    : MoveRadio.IsChecked == true ? FileActionKind.Move : FileActionKind.Copy,
                PathColumn = PathColBox.Text.Trim(),
                DestFolder = DestBox.Text.Trim(),
                FileSubfolderFormula = FileSubBox.Text.Trim(),
                FileNameFormula = FileNameBox.Text.Trim(),
                FolderSubfolderFormula = FolderSubBox.Text.Trim(),
                FolderNameFormula = FolderNameBox.Text.Trim(),
                KeepStructure = StructureBox.IsChecked == true,
                Conflict = (FileConflict)Math.Max(0, ConflictBox.SelectedIndex),
                Permanent = PermanentBox.IsChecked == true,
                Separator = SepBox.Text.Length > 0 ? SepBox.Text.Substring(0, 1) : "_"
            };
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "The file formulas are used for rows whose path is a file, the folder formulas for rows whose path is a folder. Both see the item plus every column of the row ($Key, $Value, $Relative ...).",
                FileOpsWork.FileVariableHelp(false, true));
        }

        void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = "Destination folder";
                dlg.ShowNewFolderButton = true;
                if (DestBox.Text.Trim().Length > 0)
                    dlg.SelectedPath = DestBox.Text.Trim();
                if (dlg.ShowDialog() == Forms.DialogResult.OK)
                    DestBox.Text = dlg.SelectedPath;
            }
        }

        void Use_Click(object sender, RoutedEventArgs e)
        {
            var s = Current();
            var error = FileOpsWork.Check(s);
            if (error == null && ExcelFile.FindColumn(_input, s.PathColumn) == null)
                error = "Column '" + s.PathColumn + "' is not in this step's input.";
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
