using System;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Forms = System.Windows.Forms;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.MergeDuplicates
{
    // Pipeline step editor for Merge folders. Shows the plan; never touches the disk.
    public partial class MergeDuplicatesPage : Page
    {
        readonly Action<MergeFoldersSettings> _use;

        public MergeDuplicatesPage(MergeFoldersSettings settings, Action<MergeFoldersSettings> use)
        {
            InitializeComponent();
            _use = use;
            ExcelGrid.Hook(ResultGrid);
            FormulaField.Watch(ConditionBox);
            var s = settings ?? new MergeFoldersSettings();
            FolderBox.Text = s.Folder ?? "";
            ConditionBox.Text = s.Condition ?? "";
            SepBox.Text = string.IsNullOrEmpty(s.Separator) ? "_" : s.Separator;
            KeepBox.SelectedIndex = (int)s.Keep;
            if (settings != null)
                Preview_Click(null, null);
        }

        MergeFoldersSettings Current()
        {
            return new MergeFoldersSettings
            {
                Folder = FolderBox.Text.Trim(),
                Condition = ConditionBox.Text.Trim(),
                Separator = SepBox.Text.Length > 0 ? SepBox.Text.Substring(0, 1) : "_",
                Keep = (MergeKeep)Math.Max(0, KeepBox.SelectedIndex)
            };
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            DataTable plan;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                plan = MergeDuplicatesWork.Plan(Current());
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
            ResultGrid.ItemsSource = plan.DefaultView;
            ResultInfo.Text = plan.Rows.Count == 0
                ? "No folders belong together."
                : MergeDuplicatesWork.CountReady(plan).ToString(CultureInfo.InvariantCulture) + " folder(s) would be merged into the kept one of their group.";
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "Compares two subfolders A and B. They belong together when the formula is true.",
                new[]
                {
                    new VariableHelp { Name = "$A", Description = "Name of folder A.", Example = "Acme_2024" },
                    new VariableHelp { Name = "$B", Description = "Name of folder B.", Example = "Acme_old" },
                    new VariableHelp { Name = "$A1", Description = "Split part 1 of folder A. $A2 is part 2.", Example = "Acme" },
                    new VariableHelp { Name = "$B1", Description = "Split part 1 of folder B.", Example = "Acme" },
                });
        }

        void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = "Parent folder holding the duplicates";
                dlg.ShowNewFolderButton = false;
                if (FolderBox.Text.Trim().Length > 0)
                    dlg.SelectedPath = FolderBox.Text.Trim();
                if (dlg.ShowDialog() == Forms.DialogResult.OK)
                    FolderBox.Text = dlg.SelectedPath;
            }
        }

        void Use_Click(object sender, RoutedEventArgs e)
        {
            var s = Current();
            var error = MergeDuplicatesWork.Check(s);
            if (error != null)
            {
                MessageBox.Show(Window.GetWindow(this), error, "Not ready", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            _use(s);
            Back_Click(null, null);
        }

        void Back_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
        }
    }
}
