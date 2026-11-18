using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace OfficeWorkAssistant.Features.Pipeline
{
    public enum PipelineNodeState
    {
        Idle,
        Ok,
        Error
    }

    // Tag of every port ellipse, so a hit test can tell which step and port the mouse is on.
    public sealed class PipelinePort
    {
        public string NodeId;
        public string Port;
        public bool IsOutput;
    }

    // One step box on the canvas: a card with input ports on the left and the output port on
    // the right. Fixed width; an Append step grows taller with its inputs. Code-only, so it
    // needs no XAML entry.
    public sealed class PipelineNodeView
    {
        public const double Width = 72;
        // The usual height of the icon box; see BoxHeight.
        public const double Height = 72;
        const double LabelWidth = 150;
        const double PortSize = 14;
        const double PortGap = 22;

        // Two colours, shared with the palette: green for Excel steps, orange for files and folders.
        static readonly Brush ExcelAccent = Frozen(Color.FromRgb(0x2E, 0x7D, 0x32));
        static readonly Brush FilesAccent = Frozen(Color.FromRgb(0xF5, 0x7C, 0x00));

        public static Brush AccentOf(PipelineStepKind kind)
        {
            return PipelineWork.IsFileKind(kind) ? FilesAccent : ExcelAccent;
        }

        static readonly Brush IdleBrush = Frozen(Color.FromRgb(0xBD, 0xBD, 0xBD));
        static readonly Brush OkBrush = Frozen(Color.FromRgb(0x43, 0xA0, 0x47));
        static readonly Brush ErrorBrush = Frozen(Color.FromRgb(0xE5, 0x39, 0x35));
        static readonly Brush SelectedFill = Frozen(Color.FromRgb(0xE3, 0xF2, 0xFD));
        static readonly Brush PortFill = Frozen(Color.FromRgb(0x60, 0x7D, 0x8B));

        public readonly PipelineNode Node;
        public readonly Canvas Root;
        public readonly Border Box;
        public readonly Dictionary<string, Ellipse> Outputs = new Dictionary<string, Ellipse>();
        public readonly Dictionary<string, Ellipse> Inputs = new Dictionary<string, Ellipse>();
        public readonly double BoxHeight;
        readonly string[] _ports;
        readonly string[] _outputs;
        readonly TextBlock _title;
        readonly TextBlock _detail;
        PipelineNodeState _state;
        bool _selected;

        // ports: the step's input ports now (PipelineWork.InputPorts).
        public PipelineNodeView(PipelineNode node, string[] ports, string[] outputs)
        {
            Node = node;
            _ports = ports;
            _outputs = outputs;
            BoxHeight = Math.Max(Height, PortGap * Math.Max(ports.Length, outputs.Length) + 16);
            Root = new Canvas { Width = Width, Height = BoxHeight + 44 };
            Panel.SetZIndex(Root, 1);

            // The icon is the node. Its name and summary sit centred below it, outside the box.
            _title = new TextBlock { FontWeight = FontWeights.SemiBold, TextAlignment = TextAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
            _detail = new TextBlock { FontSize = 11, Foreground = Brushes.DimGray, TextAlignment = TextAlignment.Center, TextWrapping = TextWrapping.Wrap, TextTrimming = TextTrimming.CharacterEllipsis, MaxHeight = 30 };
            var caption = new StackPanel { Width = LabelWidth, IsHitTestVisible = false };
            caption.Children.Add(_title);
            caption.Children.Add(_detail);
            Canvas.SetLeft(caption, (Width - LabelWidth) / 2);
            Canvas.SetTop(caption, BoxHeight + 4);

            Box = new Border
            {
                Width = Width,
                Height = BoxHeight,
                CornerRadius = new CornerRadius(14),
                BorderThickness = new Thickness(2),
                Background = Brushes.White,
                Child = PipelineIcons.Glyph(node.Kind, 40),
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            Root.Children.Add(Box);
            Root.Children.Add(caption);

            for (int i = 0; i < ports.Length; i++)
            {
                var port = ports[i];
                var dot = Port(port);
                int number = port == PipelineWork.PortIn ? 1 : PipelineWork.AppendPortNumber(port);
                if (node.Kind == PipelineStepKind.FillColumns)
                    number = i + 1;
                dot.ToolTip = "Drop a link here: " + PipelineWork.PortLabel(node.Kind, port);
                var y = InputY(port);
                Canvas.SetLeft(dot, -PortSize / 2);
                Canvas.SetTop(dot, y - PortSize / 2);
                Root.Children.Add(dot);
                Inputs[port] = dot;
                if (node.Kind == PipelineStepKind.FillColumns || port != PipelineWork.PortIn || ports.Length > 1)
                {
                    var name = number > 0 ? number.ToString(System.Globalization.CultureInfo.InvariantCulture) : port;
                    var label = new TextBlock { Text = name, FontSize = 10, FontWeight = FontWeights.Bold, Foreground = PortFill, IsHitTestVisible = false };
                    Canvas.SetLeft(label, PortSize / 2 + 1);
                    Canvas.SetTop(label, y - 7);
                    Root.Children.Add(label);
                }
            }

            for (int i = 0; i < outputs.Length; i++)
            {
                var output = Port(outputs[i], true);
                var number = (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                output.ToolTip = outputs.Length == 1 ? "Drag from here to the next step's input" : "Result " + number;
                output.Cursor = System.Windows.Input.Cursors.Cross;
                var y = PortY(i, outputs.Length);
                Canvas.SetLeft(output, Width - PortSize / 2);
                Canvas.SetTop(output, y - PortSize / 2);
                Root.Children.Add(output);
                Outputs.Add(outputs[i], output);
                if (outputs.Length > 1)
                {
                    var label = new TextBlock { Text = number, FontSize = 10, FontWeight = FontWeights.Bold,
                        Foreground = PortFill, IsHitTestVisible = false };
                    Canvas.SetRight(label, PortSize / 2 + 1);
                    Canvas.SetTop(label, y - 7);
                    Root.Children.Add(label);
                }
            }

            Refresh();
        }

        Ellipse Port(string port, bool isOutput = false)
        {
            return new Ellipse
            {
                Width = PortSize,
                Height = PortSize,
                Fill = PortFill,
                Stroke = Brushes.White,
                StrokeThickness = 2,
                Tag = new PipelinePort { NodeId = Node.Id, Port = port, IsOutput = isOutput }
            };
        }

        double PortY(int index, int count)
        {
            return BoxHeight * (index + 1) / (count + 1);
        }

        double InputY(string port) { return PortY(Array.IndexOf(_ports, port), _ports.Length); }

        public Point InputPoint(string port)
        {
            return new Point(Node.X, Node.Y + InputY(port));
        }

        public Point OutputPoint(string port)
        {
            return new Point(Node.X + Width, Node.Y + PortY(Math.Max(0, Array.IndexOf(_outputs, port)), _outputs.Length));
        }

        public int OutputNumber(string port)
        {
            var index = Array.IndexOf(_outputs, port);
            return index < 0 ? 1 : index + 1;
        }

        public bool HasPorts(string[] inputs, string[] outputs)
        {
            return System.Linq.Enumerable.SequenceEqual(_ports, inputs) &&
                System.Linq.Enumerable.SequenceEqual(_outputs, outputs);
        }

        public void Place()
        {
            Canvas.SetLeft(Root, Node.X);
            Canvas.SetTop(Root, Node.Y);
        }

        public bool Selected
        {
            get { return _selected; }
            set { _selected = value; Refresh(); }
        }

        public void SetState(PipelineNodeState state, string tip)
        {
            _state = state;
            Box.ToolTip = string.IsNullOrEmpty(tip) ? null : tip;
            Refresh();
        }

        // Re-read the title and summary after the node changed.
        public void Refresh()
        {
            _title.Text = Node.Title;
            _detail.Text = (PipelineWork.IsActionKind(Node.Kind) ? "Changes files. " : "") + PipelineWork.Summary(Node);
            Box.BorderBrush = _state == PipelineNodeState.Ok ? OkBrush : _state == PipelineNodeState.Error ? ErrorBrush : IdleBrush;
            Box.BorderThickness = new Thickness(_selected ? 3 : 2);
            Box.Background = _selected ? SelectedFill : Brushes.White;
            Place();
        }

        static Brush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
