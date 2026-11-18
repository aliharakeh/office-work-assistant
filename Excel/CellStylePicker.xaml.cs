using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace OfficeWorkAssistant.Excel
{
    // Edits one CellStyle: colours by name or #RRGGBB, font, border, alignment, number format,
    // with a live sample. Used by the Highlight and Format sheet steps.
    public partial class CellStylePicker : UserControl
    {
        static readonly string[] Formats =
        {
            "", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "dd/MM/yyyy", "dd/MM/yyyy HH:mm", "@"
        };

        bool _loading;

        public CellStylePicker()
        {
            _loading = true;
            InitializeComponent();
            FillBox.Items.Add("");
            foreach (var name in NamedColors.FillNames)
                FillBox.Items.Add(name);
            FontBox.Items.Add("");
            foreach (var name in NamedColors.TextNames)
                FontBox.Items.Add(name);
            foreach (var f in Formats)
                FormatBox.Items.Add(f);
            BorderBox.SelectedIndex = 0;
            AlignBox.SelectedIndex = 0;
            VerticalAlignBox.SelectedIndex = 0;
            _loading = false;

            FillBox.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new RoutedEventHandler(Changed));
            FontBox.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent, new RoutedEventHandler(Changed));
            FillBox.SelectionChanged += Changed;
            FontBox.SelectionChanged += Changed;
            BorderBox.SelectionChanged += Changed;
            AlignBox.SelectionChanged += Changed;
            VerticalAlignBox.SelectionChanged += Changed;
            foreach (var box in new[] { BoldBox, ItalicBox, UnderlineBox, StrikeBox, WrapBox })
            {
                box.Checked += Changed;
                box.Unchecked += Changed;
            }
            UpdateSample();
        }

        public bool BasicOptions
        {
            get { return ItalicBox.Visibility == Visibility.Collapsed; }
            set
            {
                var visibility = value ? Visibility.Collapsed : Visibility.Visible;
                ItalicBox.Visibility = visibility;
                UnderlineBox.Visibility = visibility;
                StrikeBox.Visibility = visibility;
            }
        }

        public CellStyle Value
        {
            get
            {
                return new CellStyle
                {
                    Fill = NamedColors.ToHex(FillBox.Text) ?? FillBox.Text.Trim(),
                    FontColor = NamedColors.ToHex(FontBox.Text) ?? FontBox.Text.Trim(),
                    Bold = BoldBox.IsChecked == true,
                    Italic = ItalicBox.IsChecked == true,
                    Underline = UnderlineBox.IsChecked == true,
                    Strike = StrikeBox.IsChecked == true,
                    Wrap = WrapBox.IsChecked == true,
                    Border = (CellBorder)System.Math.Max(0, BorderBox.SelectedIndex),
                    Align = (CellAlign)System.Math.Max(0, AlignBox.SelectedIndex),
                    VerticalAlign = (CellVerticalAlign)System.Math.Max(0, VerticalAlignBox.SelectedIndex),
                    NumberFormat = (FormatBox.Text ?? "").Trim()
                };
            }
            set
            {
                var s = value ?? new CellStyle();
                _loading = true;
                FillBox.Text = string.IsNullOrWhiteSpace(s.Fill) ? "" : NamedColors.NameOf(s.Fill);
                FontBox.Text = string.IsNullOrWhiteSpace(s.FontColor) ? "" : NamedColors.NameOf(s.FontColor);
                BoldBox.IsChecked = s.Bold;
                ItalicBox.IsChecked = s.Italic;
                UnderlineBox.IsChecked = s.Underline;
                StrikeBox.IsChecked = s.Strike;
                WrapBox.IsChecked = s.Wrap;
                BorderBox.SelectedIndex = (int)s.Border;
                AlignBox.SelectedIndex = (int)s.Align;
                VerticalAlignBox.SelectedIndex = (int)s.VerticalAlign;
                FormatBox.Text = s.NumberFormat ?? "";
                _loading = false;
                UpdateSample();
            }
        }

        // Null when the colours make sense; otherwise the reason.
        public string Error
        {
            get { return Value.Check(); }
        }

        void Changed(object sender, RoutedEventArgs e)
        {
            if (!_loading)
                UpdateSample();
        }

        void UpdateSample()
        {
            if (SampleText == null)
                return;
            // A just-picked item is not in Text yet; read it on the next pass.
            Dispatcher.BeginInvoke(new System.Action(delegate
            {
                var s = Value;
                Sample.Background = BrushOf(s.Fill) ?? Brushes.White;
                SampleText.Foreground = BrushOf(s.FontColor) ?? Brushes.Black;
                SampleText.FontWeight = s.Bold ? FontWeights.Bold : FontWeights.Normal;
                SampleText.FontStyle = s.Italic ? FontStyles.Italic : FontStyles.Normal;
                SampleText.TextAlignment = s.Align == CellAlign.Center ? TextAlignment.Center :
                    s.Align == CellAlign.Right ? TextAlignment.Right : TextAlignment.Left;
                SampleText.VerticalAlignment = s.VerticalAlign == CellVerticalAlign.Top ? VerticalAlignment.Top :
                    s.VerticalAlign == CellVerticalAlign.Center ? VerticalAlignment.Center : VerticalAlignment.Bottom;
                var deco = new TextDecorationCollection();
                if (s.Underline)
                    deco.Add(TextDecorations.Underline);
                if (s.Strike)
                    deco.Add(TextDecorations.Strikethrough);
                SampleText.TextDecorations = deco;
                Sample.BorderBrush = s.Border == CellBorder.None ? new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)) : Brushes.Black;
                Sample.BorderThickness = new Thickness(s.Border == CellBorder.Thick ? 2 : 1);
                FillBox.ToolTip = NamedColors.IsValid(FillBox.Text) ? "A colour name, or #RRGGBB. Blank = unchanged." : "Not a colour: use a name from the list or #RRGGBB.";
                FontBox.ToolTip = NamedColors.IsValid(FontBox.Text) ? "A colour name, or #RRGGBB. Blank = unchanged." : "Not a colour: use a name from the list or #RRGGBB.";
            }));
        }

        static Brush BrushOf(string color)
        {
            byte r;
            byte g;
            byte b;
            if (!NamedColors.TryRgb(color, out r, out g, out b))
                return null;
            return new SolidColorBrush(Color.FromRgb(r, g, b));
        }
    }
}
