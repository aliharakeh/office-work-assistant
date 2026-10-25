using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using Forms = System.Windows.Forms;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // Run mode: the page for the person who did not build the pipeline. It shows the saved
    // pipelines as cards on the left, asks only for the files, folders and output names a run
    // needs, says in plain words what the pipeline does, and runs it. What is filled in here
    // applies to this run only: nothing on this page is ever saved back to the pipeline.
    public partial class RunnerPage : Page
    {
        // One choice on the form: the file or folder it belongs to, the box that shows the path,
        // its sheet list (Load file only) and the red line under it that says what is missing.
        sealed class InputBlock
        {
            public RunnerInput Input;
            public TextBox Box;
            public ComboBox Sheet;
            public TextBlock SheetProblem;
            public TextBlock Problem;
            public StackPanel Root;
            // The sheet combo is filled in here, so its SelectionChanged is not the user's.
            public bool Filling;
        }

        static readonly Brush CardBorderBrush = Frozen(Color.FromRgb(0xDD, 0xE3, 0xE8));
        static readonly Brush CardHoverBrush = Frozen(Color.FromRgb(0xF1, 0xF6, 0xFB));
        static readonly Brush OpenBorderBrush = Frozen(Color.FromRgb(0x1E, 0x88, 0xE5));
        static readonly Brush OpenBackgroundBrush = Frozen(Color.FromRgb(0xE3, 0xF2, 0xFD));
        static readonly Brush ErrorBrush = Frozen(Color.FromRgb(0xC6, 0x28, 0x28));
        static readonly Brush TitleBrush = Frozen(Color.FromRgb(0x37, 0x47, 0x4F));
        static readonly Brush GreyBrush = Frozen(Color.FromRgb(0x9E, 0x9E, 0x9E));
        static readonly Brush SeparatorBrush = Frozen(Color.FromRgb(0xEE, 0xF1, 0xF4));

        readonly Dictionary<string, DataTable> _cache = new Dictionary<string, DataTable>();
        readonly List<RunnerInput> _inputs = new List<RunnerInput>();
        readonly List<InputBlock> _blocks = new List<InputBlock>();
        List<string> _folders = new List<string>();

        PipelineDefinition _def;
        string _openPath;
        Border _openCard;
        Button _openFolder;

        public RunnerPage()
        {
            InitializeComponent();
            Loaded += Page_Loaded;
        }

        // ---------- the saved pipelines ----------

        // Builder mode sets its own window title, so it is set again every time this page is shown.
        void Page_Loaded(object sender, RoutedEventArgs e)
        {
            var window = Window.GetWindow(this);
            if (window != null)
                window.Title = "Office Work Assistant";

            var saved = RefreshList();
            // Pick up where the user left off: the pipeline saved most recently.
            if (_openPath == null && saved.Count > 0)
                OpenSaved(saved[0]);
        }

        // Reads the pipelines folder again and shows one card per saved pipeline; returns what
        // was found, newest first. The open pipeline stays open.
        List<SavedPipeline> RefreshList()
        {
            Cards.Items.Clear();
            _openCard = null;
            List<SavedPipeline> saved;
            try
            {
                saved = PipelineStore.List();
            }
            catch (Exception ex)
            {
                saved = new List<SavedPipeline>();
                SetStatus("Could not read the pipelines folder: " + ex.Message);
            }
            NoPipelines.Visibility = saved.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            foreach (var one in saved)
            {
                var card = Card(one);
                Cards.Items.Add(card);
                if (IsOpen(one.Path))
                    _openCard = card;
            }
            if (_openPath != null && _openCard == null)
            {
                // The open pipeline was deleted or renamed: back to the empty state.
                ClearForm();
                SetStatus("The pipeline that was open is no longer in the pipelines folder.");
            }
            return saved;
        }

        Border Card(SavedPipeline saved)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = saved.Name, FontWeight = FontWeights.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis });
            panel.Children.Add(new TextBlock { Text = saved.ModifiedText, FontSize = 10, Foreground = GreyBrush });
            panel.Children.Add(new TextBlock
            {
                Text = CardLine(saved),
                FontSize = 11,
                Foreground = Brushes.Gray,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 2, 0, 0)
            });

            var card = new Border
            {
                Child = panel,
                Tag = saved,
                Margin = new Thickness(0, 0, 0, 6),
                Padding = new Thickness(8),
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(4),
                Cursor = Cursors.Hand,
                ToolTip = "Click to run this pipeline."
            };
            card.MouseEnter += Card_MouseEnter;
            card.MouseLeave += Card_MouseLeave;
            card.MouseLeftButtonDown += (s, e) => OpenSaved(saved);
            PaintCard(card);
            return card;
        }

        // "5 steps - 2 inputs - 1 output". A pipeline that cannot be read says so instead.
        static string CardLine(SavedPipeline saved)
        {
            try
            {
                return RunnerWork.Headline(PipelineStore.Load(saved.Path));
            }
            catch (Exception ex)
            {
                return "Cannot be read: " + ex.Message;
            }
        }

        bool IsOpen(string path)
        {
            return _openPath != null && string.Equals(path, _openPath, StringComparison.OrdinalIgnoreCase);
        }

        void PaintCard(Border card)
        {
            var saved = card.Tag as SavedPipeline;
            var open = saved != null && IsOpen(saved.Path);
            card.BorderBrush = open ? OpenBorderBrush : CardBorderBrush;
            card.Background = open ? OpenBackgroundBrush : Brushes.White;
        }

        void PaintCards()
        {
            foreach (var item in Cards.Items)
            {
                var card = item as Border;
                if (card != null)
                    PaintCard(card);
            }
        }

        Border CardFor(string path)
        {
            foreach (var item in Cards.Items)
            {
                var card = item as Border;
                var saved = card != null ? card.Tag as SavedPipeline : null;
                if (saved != null && string.Equals(saved.Path, path, StringComparison.OrdinalIgnoreCase))
                    return card;
            }
            return null;
        }

        void Card_MouseEnter(object sender, MouseEventArgs e)
        {
            var card = (Border)sender;
            var saved = card.Tag as SavedPipeline;
            if (saved == null || !IsOpen(saved.Path))
                card.Background = CardHoverBrush;
        }

        void Card_MouseLeave(object sender, MouseEventArgs e)
        {
            PaintCard((Border)sender);
        }

        void Refresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshList();
            if (_openPath != null)
                SetStatus("Looked in the pipelines folder again. " + (_def.Name ?? "") + " is still open.");
        }

        void OpenSaved(SavedPipeline saved)
        {
            PipelineDefinition def;
            try
            {
                def = PipelineStore.Load(saved.Path);
            }
            catch (Exception ex)
            {
                SetStatus("Could not open " + saved.Name + ": " + ex.Message);
                return;
            }
            _def = def;
            _openPath = saved.Path;
            _openCard = CardFor(saved.Path);
            _cache.Clear();
            _folders.Clear();
            ResultBox.Visibility = Visibility.Collapsed;
            NoSelection.Visibility = Visibility.Collapsed;
            FormScroll.Visibility = Visibility.Visible;
            SetStatus("Opened " + saved.Name + ". Fill in the files and folders it needs, then press Run.");
            BuildForm();
            PaintCards();
            if (_openCard != null)
                _openCard.BringIntoView();
        }

        void ClearForm()
        {
            _def = null;
            _openPath = null;
            _openCard = null;
            _cache.Clear();
            _folders.Clear();
            _inputs.Clear();
            _blocks.Clear();
            _openFolder = null;
            Form.Children.Clear();
            FormScroll.Visibility = Visibility.Collapsed;
            NoSelection.Visibility = Visibility.Visible;
            ResultBox.Visibility = Visibility.Collapsed;
        }

        // ---------- the form ----------

        // Built from the loaded pipeline: what to fill in, where the results go, what it does.
        void BuildForm()
        {
            Form.Children.Clear();
            _inputs.Clear();
            _blocks.Clear();
            _openFolder = null;

            if (_def == null)
                return;

            Form.Children.Add(new TextBlock
            {
                Text = _def.Name ?? "",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap
            });

            string headline;
            List<RunnerInput> inputs;
            List<RunnerStepNote> notes;
            try
            {
                // A loop or a step with no settings cannot be listed, let alone run.
                headline = RunnerWork.Headline(_def);
                inputs = RunnerWork.Inputs(_def);
                notes = RunnerWork.Describe(_def);
            }
            catch (Exception ex)
            {
                SetStatus("This pipeline cannot be run. Tell whoever built it.");
                Form.Children.Add(new TextBlock
                {
                    Text = "This pipeline cannot be run: " + ex.Message + ". Tell whoever built it, so they can fix it in Builder mode.",
                    Foreground = ErrorBrush,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 6, 0, 0)
                });
                return;
            }

            Form.Children.Add(new TextBlock { Text = headline, FontSize = 11, Foreground = Brushes.Gray, Margin = new Thickness(0, 2, 0, 6) });

            if (!RunnerWork.WritesAnything(_def))
            {
                SetStatus("Nothing to run: this pipeline does not write or change anything.");
                Form.Children.Add(new TextBlock
                {
                    Text = "This pipeline does not write or change anything yet, so there is nothing to run. " +
                        "Ask whoever built it to add a Save file, Save text / JSON, File action or Merge folders step.",
                    Foreground = Brushes.Gray,
                    TextWrapping = TextWrapping.Wrap
                });
                return;
            }

            _inputs.AddRange(inputs);
            var reads = new List<RunnerInput>();
            var writes = new List<RunnerInput>();
            foreach (var input in inputs)
            {
                if (input.IsOutput)
                    writes.Add(input);
                else
                    reads.Add(input);
            }

            if (reads.Count > 0)
            {
                Form.Children.Add(SectionHeader("FILL IN"));
                foreach (var input in reads)
                    Form.Children.Add(InputRow(input));
            }
            if (writes.Count > 0)
            {
                Form.Children.Add(Separator());
                Form.Children.Add(SectionHeader("WHERE THE RESULTS GO"));
                foreach (var input in writes)
                    Form.Children.Add(InputRow(input));
            }

            Form.Children.Add(Separator());
            Form.Children.Add(SectionHeader("WHAT THIS PIPELINE DOES"));
            for (var i = 0; i < notes.Count; i++)
                Form.Children.Add(NoteRow(notes[i], i + 1));

            Form.Children.Add(Separator());
            Form.Children.Add(ButtonRow());
        }

        // One file or folder the person running the pipeline has to choose.
        FrameworkElement InputRow(RunnerInput input)
        {
            var block = new InputBlock { Input = input };
            var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };

            panel.Children.Add(new TextBlock { Text = input.Label, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap });
            if (!string.IsNullOrEmpty(input.Note))
                panel.Children.Add(new TextBlock
                {
                    Text = input.Note,
                    FontSize = 11,
                    Foreground = Brushes.Gray,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 1, 0, 4)
                });

            var row = new DockPanel { Margin = new Thickness(0, 4, 0, 4) };
            var browse = new Button { Content = "Browse...", Width = 84, Margin = new Thickness(6, 0, 0, 0) };
            browse.Click += (s, e) => Browse(block);
            DockPanel.SetDock(browse, Dock.Right);
            block.Box = new TextBox { IsReadOnly = true, Text = input.Value, TextWrapping = TextWrapping.NoWrap };
            row.Children.Add(browse);
            row.Children.Add(block.Box);
            panel.Children.Add(row);

            if (input.Kind == RunnerInputKind.ExcelFile)
                panel.Children.Add(SheetRow(block));

            block.Problem = new TextBlock { Foreground = ErrorBrush, FontSize = 11, TextWrapping = TextWrapping.Wrap };
            panel.Children.Add(block.Problem);

            block.Root = panel;
            _blocks.Add(block);
            return panel;
        }

        // Which sheet of the chosen workbook to read (Load file steps only).
        FrameworkElement SheetRow(InputBlock block)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 4) };
            panel.Children.Add(new TextBlock { Text = "Sheet", Margin = new Thickness(0, 0, 0, 2) });

            block.Sheet = new ComboBox { Margin = new Thickness(0, 0, 0, 2) };
            block.Sheet.SelectionChanged += (s, e) =>
            {
                if (block.Filling)
                    return;
                var name = block.Sheet.SelectedItem as string;
                if (name != null)
                    block.Input.Sheet = name;
            };
            panel.Children.Add(block.Sheet);

            block.SheetProblem = new TextBlock { Foreground = ErrorBrush, FontSize = 11, TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed };
            panel.Children.Add(block.SheetProblem);

            RefreshSheet(block);
            return panel;
        }

        // The sheet names of the chosen file, or why they cannot be read yet. A file that cannot
        // be read shows the reason instead of the list.
        void RefreshSheet(InputBlock block)
        {
            if (block.Sheet == null)
                return;
            var names = block.Input.SheetNames();
            var was = block.Filling;
            block.Filling = true;
            try
            {
                if (names == null || names.Length == 0)
                {
                    block.Sheet.Visibility = Visibility.Collapsed;
                    block.Sheet.ItemsSource = null;
                    var problem = names == null ? block.Input.SheetProblem : null;
                    block.SheetProblem.Text = problem ?? "";
                    block.SheetProblem.Visibility = string.IsNullOrEmpty(problem) ? Visibility.Collapsed : Visibility.Visible;
                    return;
                }
                block.SheetProblem.Text = "";
                block.SheetProblem.Visibility = Visibility.Collapsed;
                block.Sheet.Visibility = Visibility.Visible;
                block.Sheet.ItemsSource = names;
                var sheet = block.Input.Sheet;
                if (string.IsNullOrEmpty(sheet) || Array.IndexOf(names, sheet) < 0)
                {
                    sheet = names[0];
                    // The sheet the pipeline read before is not in this file.
                    block.Input.Sheet = sheet;
                }
                block.Sheet.SelectedItem = sheet;
            }
            finally
            {
                block.Filling = was;
            }
        }

        // One numbered row of "what this pipeline does": the step's icon, name and words.
        FrameworkElement NoteRow(RunnerStepNote note, int number)
        {
            var icon = PipelineIcons.Create(note.Kind, 22);
            icon.Margin = new Thickness(0, 0, 8, 0);
            icon.VerticalAlignment = VerticalAlignment.Top;
            DockPanel.SetDock(icon, Dock.Left);

            var text = new StackPanel();
            text.Children.Add(new TextBlock
            {
                Text = number.ToString(CultureInfo.InvariantCulture) + ". " + note.Label,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap
            });
            text.Children.Add(new TextBlock { Text = note.Text, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 1, 0, 0) });
            if (!string.IsNullOrEmpty(note.Hint))
                text.Children.Add(new TextBlock
                {
                    Text = note.Hint,
                    FontSize = 11,
                    Foreground = Brushes.Gray,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 1, 0, 0)
                });

            var row = new DockPanel { Margin = new Thickness(0, 0, 0, 10) };
            row.Children.Add(icon);
            row.Children.Add(text);
            return row;
        }

        FrameworkElement ButtonRow()
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 0) };

            var check = new Button { Content = "Check", Style = StyleOf("BarButton"), Margin = new Thickness(0) };
            check.Click += Check_Click;
            check.ToolTip = "Run the pipeline without writing anything, to see what it would do.";
            row.Children.Add(check);

            var run = new Button
            {
                Content = "Run",
                Style = StyleOf("PrimaryButton"),
                Padding = new Thickness(18, 6, 18, 6),
                FontSize = 14,
                FontWeight = FontWeights.SemiBold
            };
            run.Click += Run_Click;
            run.ToolTip = "Run the pipeline and write the results.";
            row.Children.Add(run);

            _openFolder = new Button { Content = "Open output folder", Style = StyleOf("BarButton"), Visibility = Visibility.Collapsed };
            _openFolder.Click += OpenFolder_Click;
            _openFolder.ToolTip = "Show the folder the results went to.";
            row.Children.Add(_openFolder);

            return row;
        }

        TextBlock SectionHeader(string text)
        {
            var header = new TextBlock { Text = text, Style = StyleOf("SideHeader") };
            header.Margin = new Thickness(0, 4, 0, 6);
            return header;
        }

        static Border Separator()
        {
            return new Border { Height = 1, Background = SeparatorBrush, Margin = new Thickness(0, 8, 0, 10) };
        }

        Style StyleOf(string key)
        {
            return (Style)FindResource(key);
        }

        // ---------- choosing a file or folder ----------

        void Browse(InputBlock block)
        {
            var input = block.Input;
            var chosen = "";
            if (input.Kind == RunnerInputKind.ExcelFile)
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                    Title = "Choose the Excel file for " + input.StepTitle,
                    CheckFileExists = true
                };
                if (dlg.ShowDialog() != true)
                    return;
                chosen = dlg.FileName;
            }
            else if (input.Kind == RunnerInputKind.ImageFile)
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "Image files (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff",
                    Title = "Choose the image for " + input.StepTitle,
                    CheckFileExists = true
                };
                if (dlg.ShowDialog() != true)
                    return;
                chosen = dlg.FileName;
            }
            else if (input.Kind == RunnerInputKind.Folder)
            {
                using (var dlg = new Forms.FolderBrowserDialog())
                {
                    dlg.Description = "Choose the folder for " + input.StepTitle;
                    dlg.ShowNewFolderButton = true;
                    if (Directory.Exists(input.Value))
                        dlg.SelectedPath = input.Value;
                    if (dlg.ShowDialog() != Forms.DialogResult.OK)
                        return;
                    chosen = dlg.SelectedPath;
                }
            }
            else if (input.Kind == RunnerInputKind.ExcelOutput)
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "Excel workbook (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx",
                    AddExtension = true,
                    // Replacing an output file is normal here; Run asks before it happens.
                    OverwritePrompt = false,
                    Title = "Where should the result be saved?",
                    FileName = FileNameOf(input, "result.xlsx")
                };
                if (dlg.ShowDialog() != true)
                    return;
                chosen = dlg.FileName;
            }
            else
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|JSON files (*.json)|*.json|All files (*.*)|*.*",
                    FilterIndex = input.IsJson ? 2 : 1,
                    DefaultExt = input.DefaultExtension,
                    AddExtension = true,
                    OverwritePrompt = false,
                    Title = "Where should the result be saved?",
                    FileName = FileNameOf(input, "result" + input.DefaultExtension)
                };
                if (dlg.ShowDialog() != true)
                    return;
                chosen = dlg.FileName;
            }

            input.Value = chosen;
            block.Box.Text = chosen;
            RefreshSheet(block);
            block.Problem.Text = RunnerWork.InputProblem(input) ?? "";
        }

        // The file name the save dialog starts with: the one already chosen, or a sensible default.
        static string FileNameOf(RunnerInput input, string fallback)
        {
            var current = input.Value.Trim();
            if (current.Length == 0)
                return fallback;
            try
            {
                return Path.GetFileName(current);
            }
            catch (Exception)
            {
                // A path typed in Builder mode with characters Windows does not allow.
                return fallback;
            }
        }

        // ---------- checking and running ----------

        // The boxes are read-only and filled by Browse, so the paths are already in the pipeline;
        // copy them back anyway, then show every problem under its own box.
        bool Validate()
        {
            foreach (var block in _blocks)
                block.Input.Value = block.Box.Text;

            InputBlock first = null;
            foreach (var block in _blocks)
            {
                var problem = RunnerWork.InputProblem(block.Input);
                block.Problem.Text = problem ?? "";
                if (problem != null && first == null)
                    first = block;
            }

            if (RunnerWork.CheckInputs(_inputs) == null)
                return true;

            if (first != null)
            {
                SetStatus(first.Problem.Text);
                first.Root.BringIntoView();
                first.Box.Focus();
            }
            return false;
        }

        void Check_Click(object sender, RoutedEventArgs e)
        {
            if (!Validate())
                return;

            // Start clean so every Load step reads its file again.
            _cache.Clear();
            var result = RunPipeline(false);
            if (result.Error != null)
            {
                ShowResult("Needs attention", RunnerWork.FailureText(_def, result), true);
                SetStatus("This pipeline needs a change before it can run.");
                return;
            }
            var lines = RunnerWork.OutputLines(_def, _cache);
            lines.AddRange(RunnerWork.ActionLines(_def, _cache));
            ShowResult("Ready to run", lines.Count == 0 ? "This pipeline does not write anything." : Bullets(lines), false);
            SetStatus("Everything looks fine. Press Run.");
        }

        void Run_Click(object sender, RoutedEventArgs e)
        {
            if (!Validate())
                return;

            var warnings = new List<string>();
            var replaced = RunnerWork.ReplacedFiles(_def);
            if (replaced.Count > 0)
                warnings.Add("These files will be replaced:\n" + string.Join("\n", replaced.ToArray()));

            if (HasActions())
            {
                // A dry run first, so the question can say how many items each action touches.
                _cache.Clear();
                var plan = RunPipeline(false);
                if (plan.Error != null)
                {
                    ShowResult("Needs attention", RunnerWork.FailureText(_def, plan), true);
                    SetStatus("This pipeline needs a change before it can run.");
                    return;
                }
                warnings.Add("These file actions will run:\n" + string.Join("\n", RunnerWork.ActionLines(_def, _cache).ToArray()));
            }

            if (warnings.Count > 0 &&
                MessageBox.Show(Window.GetWindow(this), string.Join("\n\n", warnings.ToArray()) + "\n\nContinue?", "Run the pipeline",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            // Read every file again, then carry out the file steps and write the results.
            _cache.Clear();
            var result = RunPipeline(true);
            if (result.Error != null)
            {
                ShowResult("Needs attention", RunnerWork.FailureText(_def, result), true);
                SetStatus("Stopped. See the message above.");
                return;
            }

            _folders = RunnerWork.Folders(_def, result);
            if (_openFolder != null)
                _openFolder.Visibility = _folders.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            ShowResult(RunnerWork.DoneHeadline(result), string.Join("\n", RunnerWork.DoneLines(result).ToArray()), false);
            SetStatus("Done. You can run it again with other files.");
        }

        // True when the pipeline has a step that changes files on disk.
        bool HasActions()
        {
            foreach (var node in _def.Nodes)
            {
                if (PipelineWork.IsActionKind(node.Kind))
                    return true;
            }
            return false;
        }

        // Runs the pipeline over what the form says. Files are only touched when writeFiles is true.
        PipelineRunResult RunPipeline(bool writeFiles)
        {
            SetStatus("Working... please wait.");
            // Show the new status before the work starts.
            Dispatcher.Invoke(new Action(() => { }), System.Windows.Threading.DispatcherPriority.Render);
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                return PipelineWork.Run(_def, _cache, null, writeFiles);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        // What a check or a run found, in the panel under the form.
        void ShowResult(string title, string text, bool error)
        {
            ResultTitle.Text = title;
            ResultTitle.Foreground = error ? ErrorBrush : TitleBrush;
            Result.Text = text ?? "";
            ResultBox.Visibility = Visibility.Visible;
        }

        static string Bullets(List<string> lines)
        {
            var marked = new List<string>();
            foreach (var line in lines)
                marked.Add("• " + line);
            return string.Join("\n", marked.ToArray());
        }

        void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            if (_folders.Count == 0)
                return;
            var folder = _folders[0];
            try
            {
                // Only the first folder opens; the others are named in the status line.
                System.Diagnostics.Process.Start("explorer.exe", "\"" + folder + "\"");
            }
            catch (Exception ex)
            {
                SetStatus("Could not open the folder: " + ex.Message);
                return;
            }
            var more = _folders.Count - 1;
            SetStatus(more > 0
                ? "Opened " + folder + " (and " + more.ToString(CultureInfo.InvariantCulture) + " more folder" + (more == 1 ? "" : "s") + ")."
                : "Opened " + folder + ".");
        }

        // ---------- helpers ----------

        void SetStatus(string text)
        {
            Status.Text = text;
        }

        static Brush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
