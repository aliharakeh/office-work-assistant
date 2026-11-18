using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.Pdf
{
    public abstract class PdfEditor : UserControl
    {
        readonly StackPanel _panel = new StackPanel();

        protected PdfEditor()
        {
            Content = _panel;
        }

        protected void Note(string text)
        {
            if (!string.IsNullOrEmpty(text))
                _panel.Children.Add(new TextBlock
                {
                    Text = text, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 6)
                });
        }

        protected TextBox Field(string label, string value)
        {
            var box = new TextBox { Text = value ?? "", Margin = new Thickness(0, 0, 0, 6), MinWidth = 0 };
            Label(label, box);
            _panel.Children.Add(box);
            return box;
        }

        protected void Label(string text, FrameworkElement field)
        {
            _panel.Children.Add(new Label { Content = text, Target = field, Padding = new Thickness(0, 0, 0, 2) });
        }

        protected CheckBox Option(string label, bool value)
        {
            var box = new CheckBox { Content = label, IsChecked = value, Margin = new Thickness(0, 0, 0, 6) };
            _panel.Children.Add(box);
            return box;
        }

        protected void Finish(string title, Action apply)
        {
            var button = new Button
            {
                Content = "Apply to step", HorizontalAlignment = HorizontalAlignment.Left, Padding = new Thickness(8, 4, 8, 4)
            };
            button.Click += (sender, args) =>
            {
                try { apply(); }
                catch (Exception ex) { Notices.Show(this, ex.Message, title, MessageBoxButton.OK, MessageBoxImage.Warning); }
            };
            _panel.Children.Add(button);
        }

        protected void Column(ComboBox box)
        {
            _panel.Children.Add(box);
        }
    }

    public sealed class LoadPdfPage : PdfEditor
    {
        readonly TextBox _path;

        public LoadPdfPage(DataTable input, string label, LoadPdfSettings settings, Action<LoadPdfSettings> use)
        {
            var s = settings ?? new LoadPdfSettings();
            if (input == null)
            {
                _path = Field("PDF file", s.Path);
                Note("Or connect a table of PDF paths, such as the result of List folder.");
                Finish("Load PDF", () =>
                {
                    string path = PdfWork.InputPath(_path.Text);
                    use(new LoadPdfSettings { Path = path, PathColumn = s.PathColumn });
                });
                return;
            }
            Note(label);
            int selected = ExcelColumns.IndexOf(input, s.PathColumn);
            var pathColumn = ExcelFile.FindColumn(input, "Path");
            if (selected < 0 && string.IsNullOrEmpty(s.PathColumn) && pathColumn != null)
                selected = pathColumn.Ordinal;
            var columns = new ComboBox { ItemsSource = ExcelColumns.Labels(input), SelectedIndex = selected };
            Label("PDF path column", columns);
            Column(columns);
            Finish("Load PDF", () =>
            {
                if (columns.SelectedIndex < 0)
                    throw new InvalidOperationException("Choose the column containing the PDF paths.");
                use(new LoadPdfSettings { Path = s.Path, PathColumn = ExcelColumns.Letter(columns.SelectedIndex) });
            });
        }

        public void SetPath(string path)
        {
            if (_path != null)
                _path.Text = path;
        }
    }

    public sealed class ArrangePdfPage : PdfEditor
    {
        public ArrangePdfPage(DataTable input, string label, DataTable insert, string insertLabel,
            ArrangePdfSettings settings, Action<ArrangePdfSettings> use)
        {
            Note("A · " + label + " (" + input.Rows.Count + " pages)");
            Note(insert == null ? "Connect a second PDF to B to insert its pages." :
                "B · " + insertLabel + " (" + insert.Rows.Count + " pages)");
            var pages = Field("Pages in output order", (settings ?? new ArrangePdfSettings()).Pages);
            Note("Use commas, ranges and repeated page numbers. Use all, last, odd, even or blank (A4). " +
                "Example: 1-2, B:1, 3-last, blank, 2@90. Use last-1 to reverse all pages.");
            Finish("Arrange PDF pages", () =>
            {
                var value = new ArrangePdfSettings { Pages = pages.Text.Trim() };
                PdfWork.Arrange(input, insert, value);
                use(value);
            });
        }
    }

    public sealed class MergePdfPage : PdfEditor
    {
        public MergePdfPage(IList<DataTable> inputs, IList<string> labels, MergePdfSettings settings, Action<MergePdfSettings> use)
        {
            var s = settings ?? new MergePdfSettings();
            for (int i = 0; i < inputs.Count; i++)
                Note((i + 1) + ". " + labels[i] + " (" + inputs[i].Rows.Count + " pages)");
            Note("PDFs are appended in input port order. Interleave alternates pages from exactly two documents.");
            var interleave = Option("Interleave pages", s.Interleave);
            var reverse = Option("Reverse the second PDF", s.ReverseSecond);
            reverse.IsEnabled = s.Interleave;
            interleave.Checked += (sender, args) => reverse.IsEnabled = true;
            interleave.Unchecked += (sender, args) => reverse.IsEnabled = false;
            Finish("Merge PDFs", () =>
            {
                var value = new MergePdfSettings { Interleave = interleave.IsChecked == true, ReverseSecond = reverse.IsChecked == true };
                PdfWork.Merge(inputs, value);
                use(value);
            });
        }
    }

    public sealed class SplitPdfPage : PdfEditor
    {
        public SplitPdfPage(DataTable input, string label, SplitPdfSettings settings, Action<SplitPdfSettings> use)
        {
            var s = settings ?? new SplitPdfSettings();
            Note(label);
            var count = Field("Pages per output PDF", s.PagesPerFile.ToString(CultureInfo.InvariantCulture));
            var ranges = Field("Page groups (optional)", s.Ranges);
            Note("Separate groups with semicolons, for example 1-3; 4-last. With page groups filled in, the page count is ignored. Connect Save PDF to write the results.");
            Finish("Split PDF", () =>
            {
                var value = new SplitPdfSettings { Ranges = ranges.Text.Trim() };
                if (value.Ranges.Length == 0)
                {
                    int number;
                    if (!int.TryParse(count.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out number) || number < 1)
                        throw new InvalidOperationException("Enter at least one page per output PDF.");
                    value.PagesPerFile = number;
                }
                PdfWork.Split(input, value);
                use(value);
            });
        }
    }

    public sealed class SavePdfPage : PdfEditor
    {
        readonly TextBox _path;

        public SavePdfPage(DataTable input, string label, SavePdfSettings settings, Action<SavePdfSettings> use)
        {
            Note(label);
            _path = Field("Output PDF path", (settings ?? new SavePdfSettings()).Path);
            Note("Multiple documents use numbered names such as result-001.pdf. Preview lists the files to be written.");
            Finish("Save PDF", () =>
            {
                var value = new SavePdfSettings { Path = _path.Text.Trim() };
                PdfWork.PlanSave(input, value);
                use(value);
            });
        }

        public void SetPath(string path)
        {
            _path.Text = path;
        }
    }
}

