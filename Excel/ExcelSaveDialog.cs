using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace OfficeWorkAssistant.Excel
{
    // One shared "where should the result go" prompt: a new Excel file, or a new sheet
    // inside one of the .xlsx files the feature has loaded.
    public static class ExcelSaveDialog
    {
        sealed class Target
        {
            public string Label;
            public string Path;
            public override string ToString() { return Label; }
        }

        // loadedPaths: files currently loaded in the feature (label -> path). Blank or missing
        // files and .xlsm files (saving would drop their macros) are not offered.
        // Returns a short success message, or null when the user cancelled.
        public static string Show(Window owner, DataTable table, string defaultFileName,
            IList<KeyValuePair<string, string>> loadedPaths)
        {
            var targets = new List<Target>();
            if (loadedPaths != null)
            {
                foreach (var item in loadedPaths)
                {
                    if (string.IsNullOrWhiteSpace(item.Value) || !File.Exists(item.Value))
                        continue;
                    if (!string.Equals(System.IO.Path.GetExtension(item.Value), ".xlsx", StringComparison.OrdinalIgnoreCase))
                        continue;
                    targets.Add(new Target
                    {
                        Label = item.Key + ": " + System.IO.Path.GetFileName(item.Value),
                        Path = item.Value
                    });
                }
            }

            var win = new Window
            {
                Title = "Save result",
                Width = 480,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = owner
            };
            var stack = new StackPanel { Margin = new Thickness(14) };
            win.Content = stack;

            stack.Children.Add(new TextBlock
            {
                Text = "Save " + table.Rows.Count.ToString("N0") + (table.Rows.Count == 1 ? " row" : " rows") + " to:",
                Margin = new Thickness(0, 0, 0, 8)
            });

            var newFile = new RadioButton { Content = "A new Excel file", IsChecked = true, Margin = new Thickness(0, 0, 0, 6) };
            var newSheet = new RadioButton
            {
                Content = "A new sheet inside a loaded Excel file",
                IsEnabled = targets.Count > 0,
                ToolTip = targets.Count > 0 ? null : "Load an .xlsx file first. .xlsm files are not offered because saving would drop their macros."
            };
            stack.Children.Add(newFile);
            stack.Children.Add(newSheet);

            var sheetPanel = new StackPanel { Margin = new Thickness(20, 6, 0, 0), IsEnabled = false };
            var fileBox = new ComboBox { ItemsSource = targets, Margin = new Thickness(0, 0, 0, 6) };
            if (targets.Count > 0)
                fileBox.SelectedIndex = 0;
            var sheetRow = new DockPanel();
            sheetRow.Children.Add(new TextBlock { Text = "New sheet name", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0) });
            var sheetName = new TextBox { Text = "Result", MaxLength = 31 };
            sheetRow.Children.Add(sheetName);
            sheetPanel.Children.Add(fileBox);
            sheetPanel.Children.Add(sheetRow);
            sheetPanel.Children.Add(new TextBlock
            {
                Text = "The other sheets stay as they are. Close the file in Excel first. Charts and pivot tables in it may not be kept.",
                TextWrapping = TextWrapping.Wrap,
                Opacity = 0.7,
                Margin = new Thickness(0, 6, 0, 0)
            });
            stack.Children.Add(sheetPanel);

            newSheet.Checked += delegate { sheetPanel.IsEnabled = true; };
            newSheet.Unchecked += delegate { sheetPanel.IsEnabled = false; };

            var ok = new Button { Content = "Save", Width = 90, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            var cancel = new Button { Content = "Cancel", Width = 90, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 14, 0, 0) };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            stack.Children.Add(buttons);

            string message = null;
            ok.Click += delegate
            {
                try
                {
                    if (newSheet.IsChecked == true)
                    {
                        var target = fileBox.SelectedItem as Target;
                        if (target == null)
                            return;
                        ExcelFile.Save(table, target.Path, sheetName.Text, true);
                        message = "Added sheet \"" + sheetName.Text.Trim() + "\" with " + table.Rows.Count +
                            " rows to " + System.IO.Path.GetFileName(target.Path) + ".";
                    }
                    else
                    {
                        var dlg = new SaveFileDialog
                        {
                            Filter = "Excel files (*.xlsx)|*.xlsx",
                            FileName = defaultFileName
                        };
                        if (dlg.ShowDialog(win) != true)
                            return;
                        ExcelFile.Save(table, dlg.FileName);
                        message = "Saved " + table.Rows.Count + " rows.";
                    }
                    win.DialogResult = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(win, ex.Message, "Could not save", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            };

            return win.ShowDialog() == true ? message : null;
        }
    }
}
