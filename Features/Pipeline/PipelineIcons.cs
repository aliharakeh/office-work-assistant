using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace OfficeWorkAssistant.Features.Pipeline
{
    // One line icon per step kind, drawn as vector paths on a 24 x 24 grid so they stay sharp at any
    // size and need no icon font (Windows 7 has no Segoe MDL2). Code-only, no XAML entry.
    public static class PipelineIcons
    {
        static readonly Dictionary<PipelineStepKind, Geometry> Cache = new Dictionary<PipelineStepKind, Geometry>();

        static string Data(PipelineStepKind kind)
        {
            switch (kind)
            {
                // Excel file: a page with a folded corner and an X.
                case PipelineStepKind.Load: return "M5,2 H14 L19,7 V22 H5 Z M14,2 V7 H19 M8.5,10.5 L15.5,18.5 M15.5,10.5 L8.5,18.5";
                // Floppy disk.
                case PipelineStepKind.Save: return "M4,4 H17 L20,7 V20 H4 Z M8,4 V9 H15 V4 M8,20 V14 H16 V20";
                // Funnel.
                case PipelineStepKind.FilterSort: return "M3,4 H21 L14,12 V19 L10,21 V12 Z";
                // Plus sign: add columns.
                case PipelineStepKind.Templates: return "M12,4 V20 M4,12 H20";
                // Two arrows, opposite ways.
                case PipelineStepKind.Compare: return "M4,8 H19 M15,4 L19,8 L15,12 M20,16 H5 M9,12 L5,16 L9,20";
                // Column with an arrow filling down.
                case PipelineStepKind.FillColumns: return "M5,3 H19 V21 H5 Z M5,9 H19 M12,12 V19 M9,16 L12,19 L15,16";
                // Bulleted list.
                case PipelineStepKind.ValueList: return "M9,6 H20 M9,12 H20 M9,18 H20 M4,6 H4.5 M4,12 H4.5 M4,18 H4.5";
                // Folder with lines.
                case PipelineStepKind.ListFolder: return "M3,6 H10 L12,8 H21 V19 H3 Z M7,13 H17 M7,16 H14";
                // Magnifier.
                case PipelineStepKind.FindFiles: return "M10,3 A7,7 0 1 1 10,17 A7,7 0 1 1 10,3 Z M15,15 L21,21";
                // Lightning bolt: acts on the disk.
                case PipelineStepKind.FileAction: return "M13,2 L5,13 H11 L10,22 L19,10 H12.5 Z";
                // Two lines joining into one.
                case PipelineStepKind.MergeFolders: return "M4,5 H9 L15,12 H21 M4,19 H9 L15,12 M18,9 L21,12 L18,15";
                // Broom.
                case PipelineStepKind.CleanText: return "M20,4 L11,13 M8,10 L14,16 L9,21 H4 V15 Z M6,17 L8,19";
                // One line branching into two outputs, arrows at the right.
                case PipelineStepKind.SplitColumn: return "M3,12 H8 L13,6 H20 M8,12 L13,18 H20 M17,3.5 L20,6 L17,8.5 M17,15.5 L20,18 L17,20.5";
                // Two lines joining into one output, arrow at the right.
                case PipelineStepKind.CombineColumns: return "M3,6 H8 L13,12 H20 M3,18 H8 L13,12 M17,9.5 L20,12 L17,14.5";
                // Two overlapping squares.
                case PipelineStepKind.RemoveDuplicates: return "M8,8 H20 V20 H8 Z M16,8 V4 H4 V16 H8";
                // Three columns with a move arrow.
                case PipelineStepKind.ArrangeColumns: return "M3,3 H8 V14 H3 Z M10,3 H15 V14 H10 Z M17,3 H21 V14 H17 Z M5,19 H19 M16,16 L19,19 L16,22";
                case PipelineStepKind.EditHeaders: return "M3,4 H21 V20 H3 Z M3,10 H21 M9,10 V20 M15,10 V20 M6,7 H10 M14,7 H18";
                // Two tables stacked, plus.
                case PipelineStepKind.Append: return "M4,3 H20 V8 H4 Z M4,10 H20 V15 H4 Z M12,17 V23 M9,20 H15";
                // Marker pen.
                case PipelineStepKind.Highlight: return "M15,4 L20,9 L10,19 L5,19 L5,14 Z M4,22 H20";
                // Table with a header row.
                case PipelineStepKind.FormatSheet: return "M3,4 H21 V20 H3 Z M3,9 H21 M9,9 V20 M15,9 V20";
                // Text page with a save arrow.
                case PipelineStepKind.SaveText: return "M4,2 H15 L20,7 V13 M15,2 V7 H20 M4,2 V22 H11 M7,10 H15 M7,14 H11 M17,14 V22 M13,18 L17,22 L21,18";
                case PipelineStepKind.LoadPdf: return "M5,2 H14 L19,7 V22 H5 Z M14,2 V7 H19 M8,13 H16 M12,9 V17 M9,14 L12,17 L15,14";
                case PipelineStepKind.ArrangePdf: return "M3,4 H15 M3,12 H15 M3,20 H15 M19,5 V19 M17,7 L19,5 L21,7 M17,17 L19,19 L21,17";
                case PipelineStepKind.MergePdf: return "M3,3 H10 V10 H3 Z M14,3 H21 V10 H14 Z M6,10 V14 H12 M18,10 V14 H12 M12,14 V21 M9,18 L12,21 L15,18";
                case PipelineStepKind.SplitPdf: return "M8,3 H16 V10 H8 Z M12,10 V14 H5 V20 M12,14 H19 V20 M3,18 L5,20 L7,18 M17,18 L19,20 L21,18";
                case PipelineStepKind.SavePdf: return "M4,2 H13 L18,7 V12 M13,2 V7 H18 M4,2 V22 H11 M15,13 V22 M11,18 L15,22 L19,18";
                default: return "M4,4 H20 V20 H4 Z";
            }
        }

        static Geometry Shape(PipelineStepKind kind)
        {
            Geometry g;
            if (!Cache.TryGetValue(kind, out g))
            {
                g = Geometry.Parse(Data(kind));
                g.Freeze();
                Cache[kind] = g;
            }
            return g;
        }

        // The icon for a step kind: a tinted round badge holding the line drawing, `size` wide.
        // Just the line drawing, `size` wide, in the step's category colour.
        public static FrameworkElement Glyph(PipelineStepKind kind, double size)
        {
            var path = new Path
            {
                Data = Shape(kind),
                Stroke = PipelineNodeView.AccentOf(kind),
                StrokeThickness = 1.8,
                StrokeLineJoin = PenLineJoin.Round,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round,
                Width = 24,
                Height = 24
            };
            return new Viewbox { Width = size, Height = size, Child = path, IsHitTestVisible = false };
        }

        public static FrameworkElement Create(PipelineStepKind kind, double size)
        {
            var accent = PipelineNodeView.AccentOf(kind);
            var tint = new SolidColorBrush(((SolidColorBrush)accent).Color) { Opacity = 0.14 };
            tint.Freeze();
            var box = Glyph(kind, size * 0.58);
            return new Border
            {
                Width = size,
                Height = size,
                CornerRadius = new CornerRadius(size / 2),
                Background = tint,
                Child = box,
                SnapsToDevicePixels = false,
                IsHitTestVisible = false
            };
        }
    }
}
