using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
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

        void SepBox_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (CustomSep != null)
                CustomSep.IsEnabled = SepBox.SelectedIndex == 3;
        }

        void WrapBox_Changed(object sender, RoutedEventArgs e)
        {
            if (FolderOnlyBox != null)
            {
                FolderOnlyBox.IsEnabled = WrapBox.IsChecked == true;
                if (WrapBox.IsChecked != true)
                    FolderOnlyBox.IsChecked = false;
            }
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            char sep;
            if (!TryReady(out sep))
                return;

            try
            {
                _plans = CopyFilesWork.PlanMoves(
                    PathSource.Text, PathDest.Text, sep, PatternBox.Text,
                    RecurseBox.IsChecked == true, WrapBox.IsChecked == true,
                    ExtFilterBox.Text, ContainsBox.Text,
                    FolderOnlyBox.IsChecked == true, FolderBox.Text);
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

        bool TryReady(out char separator)
        {
            separator = SelectedSeparator();
            if (string.IsNullOrWhiteSpace(PathSource.Text) || string.IsNullOrWhiteSpace(PathDest.Text))
            {
                Alert("Choose source and destination folders.", "Missing folder", MessageBoxImage.Warning);
                return false;
            }
            if (SepBox.SelectedIndex == 3 && CustomSep.Text.Length == 0)
            {
                Alert("Enter a custom separator.", "Missing separator", MessageBoxImage.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(PatternBox.Text))
            {
                Alert("Enter a rename pattern such as $1_$2.", "Missing pattern", MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        char SelectedSeparator()
        {
            switch (SepBox.SelectedIndex)
            {
                case 0: return ' ';
                case 2: return '-';
                case 3: return CustomSep.Text.Length > 0 ? CustomSep.Text[0] : '_';
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
