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
        public const string Out = "Out";
        public string NodeId;
        public string Port;
    }

    // One step box on the canvas: a card with input ports on the left and the output port on
    // the right. Fixed width; an Append step grows taller with its inputs. Code-only, so it
    // needs no XAML entry.
    public sealed class PipelineNodeView
    {
        public const double Width = 200;
        // The usual height; see BoxHeight.
        public const double Height = 80;
        const double PortSize = 14;
        const double PortGap = 22;

        // Step categories, shared with the palette: data in/out, Excel work, styling, files, disk changes.
        static readonly Brush InOutAccent = Frozen(Color.FromRgb(0x00, 0x89, 0x7B));
        static readonly Brush ExcelAccent = Frozen(Color.FromRgb(0x1E, 0x88, 0xE5));
        static readonly Brush StyleAccent = Frozen(Color.FromRgb(0x8E, 0x24, 0xAA));
        static readonly Brush FilesAccent = Frozen(Color.FromRgb(0xF5, 0x7C, 0x00));
        static readonly Brush ActionAccent = Frozen(Color.FromRgb(0xD8, 0x1B, 0x60));

        public static Brush AccentOf(PipelineStepKind kind)
        {
            if (PipelineWork.IsActionKind(kind))
                return ActionAccent;
            if (PipelineWork.IsFileKind(kind))
                return FilesAccent;
            if (PipelineWork.IsStyleKind(kind))
                return StyleAccent;
            if (kind == PipelineStepKind.Load || kind == PipelineStepKind.Save)
                return InOutAccent;
            return ExcelAccent;
        }

        static readonly Brush IdleBrush = Frozen(Color.FromRgb(0xBD, 0xBD, 0xBD));
        static readonly Brush OkBrush = Frozen(Color.FromRgb(0x43, 0xA0, 0x47));
        static readonly Brush ErrorBrush = Frozen(Color.FromRgb(0xE5, 0x39, 0x35));
        static readonly Brush SelectedFill = Frozen(Color.FromRgb(0xE3, 0xF2, 0xFD));
        static readonly Brush PortFill = Frozen(Color.FromRgb(0x60, 0x7D, 0x8B));

        public readonly PipelineNode Node;
        public readonly Canvas Root;
        public readonly Border Box;
        public readonly Ellipse Output;
        public readonly Dictionary<string, Ellipse> Inputs = new Dictionary<string, Ellipse>();
        public readonly double BoxHeight;
        readonly string[] _ports;
        readonly TextBlock _title;
        readonly TextBlock _kind;
        readonly TextBlock _detail;
        PipelineNodeState _state;
        bool _selected;

        // ports: the step's input ports now (PipelineWork.InputPorts).
        public PipelineNodeView(PipelineNode node, string[] ports)
        {
            Node = node;
            _ports = ports;
            BoxHeight = Math.Max(Height, PortGap * ports.Length + 16);
            Root = new Canvas { Width = Width, Height = BoxHeight };
            Panel.SetZIndex(Root, 1);

            var accent = AccentOf(node.Kind);
            _kind = new TextBlock { FontSize = 10, FontWeight = FontWeights.SemiBold, Foreground = accent };
            _title = new TextBlock { FontWeight = FontWeights.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis };
            _detail = new TextBlock { FontSize = 11, Foreground = Brushes.DimGray, TextWrapping = TextWrapping.Wrap, TextTrimming = TextTrimming.CharacterEllipsis, MaxHeight = 30 };
            var text = new StackPanel { Margin = new Thickness(10, 5, 12, 4) };
            text.Children.Add(_kind);
            text.Children.Add(_title);
            text.Children.Add(_detail);
            var stripe = new Border { Width = 6, Background = accent, CornerRadius = new CornerRadius(4, 0, 0, 4) };
            DockPanel.SetDock(stripe, Dock.Left);
            var inner = new DockPanel();
            inner.Children.Add(stripe);
            inner.Children.Add(text);
            Box = new Border
            {
                Width = Width,
                Height = BoxHeight,
                CornerRadius = new CornerRadius(6),
                BorderThickness = new Thickness(2),
                Background = Brushes.White,
                Child = inner,
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            Root.Children.Add(Box);

            foreach (var port in ports)
            {
                var dot = Port(port);
                int number = PipelineWork.AppendPortNumber(port);
                dot.ToolTip = number == ports.Length && node.Kind == PipelineStepKind.Append
                    ? "Drop a link here to add another table"
                    : "Drop a link here: " + PipelineWork.PortLabel(node.Kind, port);
                var y = InputY(port);
                Canvas.SetLeft(dot, -PortSize / 2);
                Canvas.SetTop(dot, y - PortSize / 2);
                Root.Children.Add(dot);
                Inputs[port] = dot;
                if (port != PipelineWork.PortIn)
                {
                    var name = number > 0 ? number.ToString(System.Globalization.CultureInfo.InvariantCulture) : port;
                    var label = new TextBlock { Text = name, FontSize = 10, FontWeight = FontWeights.Bold, Foreground = PortFill, IsHitTestVisible = false };
                    Canvas.SetLeft(label, PortSize / 2 + 1);
                    Canvas.SetTop(label, y - 7);
                    Root.Children.Add(label);
                }
            }

            Output = Port(PipelinePort.Out);
            Output.ToolTip = "Drag from here to the next step's input";
            Output.Cursor = System.Windows.Input.Cursors.Cross;
            Canvas.SetLeft(Output, Width - PortSize / 2);
            Canvas.SetTop(Output, BoxHeight / 2 - PortSize / 2);
            Root.Children.Add(Output);

            Refresh();
        }

        Ellipse Port(string port)
        {
            return new Ellipse
            {
                Width = PortSize,
                Height = PortSize,
                Fill = PortFill,
                Stroke = Brushes.White,
                StrokeThickness = 2,
                Tag = new PipelinePort { NodeId = Node.Id, Port = port }
            };
        }

        double InputY(string port)
        {
            if (port == PipelineWork.PortA)
                return BoxHeight * 0.3;
            if (port == PipelineWork.PortB)
                return BoxHeight * 0.7;
            int number = PipelineWork.AppendPortNumber(port);
            if (number > 0)
                return BoxHeight * number / (_ports.Length + 1);
            return BoxHeight / 2;
        }

        public Point InputPoint(string port)
        {
            return new Point(Node.X, Node.Y + InputY(port));
        }

        public Point OutputPoint()
        {
            return new Point(Node.X + Width, Node.Y + BoxHeight / 2);
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
            _kind.Text = PipelineWork.KindLabel(Node.Kind).ToUpperInvariant() +
                (PipelineWork.IsActionKind(Node.Kind) ? "  -  CHANGES FILES" : "");
            _title.Text = Node.Title;
            _detail.Text = PipelineWork.Summary(Node);
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
