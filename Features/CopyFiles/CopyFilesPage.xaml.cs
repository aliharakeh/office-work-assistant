using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using WorkAssistant.Expressions;
using WorkAssistant.Views;

namespace WorkAssistant.Features.CopyFiles
{
    public partial class CopyFilesPage : Page
    {
        List<CopyFilesPlan> _plans;

        public CopyFilesPage()
        {
            InitializeComponent();
        }

        void Home_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
            else if (NavigationService != null)
                NavigationService.Navigate(new HomePage());
        }

        void BrowseSource_Click(object sender, RoutedEventArgs e)
        {
            PickFolder(PathSource, "Source folder");
        }

        void BrowseDest_Click(object sender, RoutedEventArgs e)
        {
            PickFolder(PathDest, "Destination folder");
        }

        void FileSep_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (FileCustomSep != null)
                FileCustomSep.IsEnabled = FileSepBox.SelectedIndex == 3;
        }

        void FolderSep_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (FolderCustomSep != null)
                FolderCustomSep.IsEnabled = FolderSepBox.SelectedIndex == 3;
        }

        void WrapBox_Changed(object sender, RoutedEventArgs e)
        {
            if (FolderPatternBox != null)
                FolderPatternBox.IsEnabled = WrapBox.IsChecked == true;
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "File filter, folder filter, file pattern and folder pattern all use the same formulas. File boxes use file variables; folder boxes use folder variables. $1 $2 are split parts.",
                new[]
                {
                    new VariableHelp { Name = "{Name}", Description = "File: full file name. Folder: folder name." },
                    new VariableHelp { Name = "{Stem}", Description = "File name without extension." },
                    new VariableHelp { Name = "{Ext}", Description = "File extension, including the dot." },
                    new VariableHelp { Name = "{Path}", Description = "Full path of the file or folder." },
                    new VariableHelp { Name = "{Folder}", Description = "Directory that holds the file." },
                    new VariableHelp { Name = "{FolderName}", Description = "Name of that directory." },
                    new VariableHelp { Name = "{Relative}", Description = "Folder path relative to Source." },
                    new VariableHelp { Name = "{Size}", Description = "File size in bytes." },
                    new VariableHelp { Name = "{Modified}", Description = "File last-write time." },
                    new VariableHelp { Name = "$1", Description = "First split part. $2 is the second." },
                });
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            char fileSep;
            char folderSep;
            if (!TryReady(out fileSep, out folderSep))
                return;

            var wrap = WrapBox.IsChecked == true;
            if (!OkExpr("File pattern", FilePatternBox.Text, true) ||
                !OkExpr("Folder pattern", FolderPatternBox.Text, wrap) ||
                !OkExpr("File filter", FileFilterBox.Text, false) ||
                !OkExpr("Folder filter", FolderFilterBox.Text, false))
                return;

            try
            {
                _plans = CopyFilesWork.PlanMoves(
                    PathSource.Text, PathDest.Text, fileSep, folderSep, FilePatternBox.Text,
                    RecurseBox.IsChecked == true, wrap,
                    FolderPatternBox.Text, FileFilterBox.Text, FolderFilterBox.Text);
                PreviewGrid.ItemsSource = _plans;
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not list files", MessageBoxImage.Error);
            }
        }

        void Copy_Click(object sender, RoutedEventArgs e)
        {
            if (_plans == null || _plans.Count == 0)
            {
                Alert("Preview first.", "No plan", MessageBoxImage.Warning);
                return;
            }

            CopyFilesWork.ApplyMoves(_plans);
            PreviewGrid.Items.Refresh();

            var moved = 0;
            var skipped = 0;
            var failed = 0;
            foreach (var p in _plans)
            {
                if (p.Status == "Copied")
                    moved++;
                else if (p.Status != null && p.Status.StartsWith("Failed", StringComparison.Ordinal))
                    failed++;
                else
                    skipped++;
            }

            Alert("Copied " + moved + ", skipped " + skipped + ", failed " + failed + ".", "Done",
                failed > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }

        bool TryReady(out char fileSep, out char folderSep)
        {
            fileSep = SelectedSeparator(FileSepBox, FileCustomSep);
            folderSep = SelectedSeparator(FolderSepBox, FolderCustomSep);
            if (string.IsNullOrWhiteSpace(PathSource.Text) || string.IsNullOrWhiteSpace(PathDest.Text))
            {
                Alert("Choose source and destination folders.", "Missing folder", MessageBoxImage.Warning);
                return false;
            }
            if (FileSepBox.SelectedIndex == 3 && FileCustomSep.Text.Length == 0)
            {
                Alert("Enter a custom file split character.", "Missing split", MessageBoxImage.Warning);
                return false;
            }
            if (FolderSepBox.SelectedIndex == 3 && FolderCustomSep.Text.Length == 0)
            {
                Alert("Enter a custom folder split character.", "Missing split", MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        bool OkExpr(string label, string expression, bool required)
        {
            var err = CopyFilesWork.ValidateExpression(expression, required);
            if (err == null)
                return true;
            Alert(label + ": " + err, "Bad formula", MessageBoxImage.Warning);
            return false;
        }

        char SelectedSeparator(ComboBox box, TextBox custom)
        {
            switch (box.SelectedIndex)
            {
                case 0: return ' ';
                case 2: return '-';
                case 3: return custom.Text.Length > 0 ? custom.Text[0] : '_';
                default: return '_';
            }
        }

        void PickFolder(TextBox target, string title)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = title;
                dlg.ShowNewFolderButton = true;
                if (dlg.ShowDialog() == Forms.DialogResult.OK)
                    target.Text = dlg.SelectedPath;
            }
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }
    }
}
