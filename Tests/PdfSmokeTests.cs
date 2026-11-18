using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.Pdf;
using OfficeWorkAssistant.Features.Pipeline;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

internal static class PdfSmokeTests
{
    static int _checks;

    static void Check(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
        _checks++;
    }

    static PipelineNode Step(PipelineDefinition def, PipelineStepKind kind, object settings, PipelineNode input, string port)
    {
        var node = PipelineWork.AddNode(def, kind, 0, 0);
        node.Settings = settings;
        if (input != null)
            Link(def, input, node, port);
        return node;
    }

    static void Link(PipelineDefinition def, PipelineNode from, PipelineNode to, string port)
    {
        def.Links.Add(new PipelineLink { From = from.Id, To = to.Id, Port = port });
    }

    static void CreatePdf(string path, params int[] widths)
    {
        using (var pdf = new PdfDocument())
        {
            foreach (int width in widths)
            {
                var page = pdf.AddPage();
                page.Width = XUnit.FromPoint(width);
                page.Height = XUnit.FromPoint(640);
                using (var graphics = XGraphics.FromPdfPage(page))
                    graphics.DrawRectangle(XBrushes.Black, 10, 10, 20, 20);
            }
            pdf.Save(path);
        }
    }

    static void Output(string path, params int[] widths)
    {
        using (var pdf = PdfReader.Open(path, PdfDocumentOpenMode.Import))
        {
            Check(pdf.PageCount == widths.Length, "Wrong page count in " + path);
            for (int i = 0; i < widths.Length; i++)
            {
                Check(Math.Round(pdf.Pages[i].Width.Point) == widths[i], "Page order or page size changed in " + path);
                if (widths[i] != 595)
                    Check(pdf.Pages[i].Contents.Elements.Count > 0, "A copied page lost its content.");
            }
        }
    }

    static void TestPipeline(string folder, string first, string second)
    {
        var def = new PipelineDefinition();
        var a = Step(def, PipelineStepKind.LoadPdf, new LoadPdfSettings { Path = first }, null, null);
        var b = Step(def, PipelineStepKind.LoadPdf, new LoadPdfSettings { Path = second }, null, null);
        var arrange = Step(def, PipelineStepKind.ArrangePdf,
            new ArrangePdfSettings { Pages = "last-2, B:1, blank, 1@90, 1@-90" }, a, "A");
        Link(def, b, arrange, "B");
        var merge = Step(def, PipelineStepKind.MergePdf, new MergePdfSettings(), arrange, "In1");
        Link(def, b, merge, "In2");
        Check(PipelineWork.InputPorts(def, merge).SequenceEqual(new[] { "In1", "In2", "In3" }), "Missing open merge port.");
        var split = Step(def, PipelineStepKind.SplitPdf, new SplitPdfSettings { Ranges = "1-4; 5-last" }, merge, "In");

        string basePath = Path.Combine(folder, "result");
        string output1 = basePath + "-001.pdf";
        string output2 = basePath + "-002.pdf";
        var save = Step(def, PipelineStepKind.SavePdf, new SavePdfSettings { Path = basePath }, split, "In");
        CreatePdf(output1, 200);
        byte[] original = File.ReadAllBytes(first);
        byte[] existing = File.ReadAllBytes(output1);
        a.ShowInRunMode = true;
        save.ShowInRunMode = true;
        string xml = Path.Combine(folder, "workflow.xml");
        PipelineWork.SaveFile(def, xml);
        def = PipelineWork.LoadFile(xml);
        var choices = RunnerWork.Inputs(def);
        Check(choices.Count == 3 && choices[0].Kind == RunnerInputKind.PdfFile &&
            choices[2].Kind == RunnerInputKind.PdfOutput && choices[2].IsOutput &&
            RunnerWork.CheckInputs(choices) == null, "Run mode lost the PDF path choices.");
        Check(RunnerWork.WritesAnything(def), "Save PDF was not treated as an output.");
        Check(PipelineStepEditor.Prepare(def, PipelineWork.Find(def, arrange.Id), ignored => { }) is ArrangePdfPage &&
            PipelineStepEditor.Prepare(def, PipelineWork.Find(def, save.Id), ignored => { }) is SavePdfPage,
            "Run mode could not prepare the shared PDF editors.");
        Check(existing.SequenceEqual(File.ReadAllBytes(output1)) && !File.Exists(output2), "Preparing editors wrote a PDF.");

        var cache = new Dictionary<string, DataTable>();
        var check = PipelineWork.Run(def, cache, null, false);
        Check(check.Error == null && check.Written.Count == 0, "PDF preview failed: " + check.Error);
        Check(existing.SequenceEqual(File.ReadAllBytes(output1)) && !File.Exists(output2), "Preview wrote a PDF.");
        Check(RunnerWork.ReplacedFiles(def, cache).SequenceEqual(new[] { output1 }), "Wrong list of PDFs to replace.");
        var lines = RunnerWork.OutputLines(def, cache);
        Check(lines.Count == 2 && lines[0].Contains("4 pages") && lines[1].Contains("4 pages"), "Check omitted split files or page counts.");
        Check((int)cache[a.Id].Rows[0]["Rotation"] == 0, "Arrange mutated a source page.");

        var result = PipelineWork.Run(def, cache, null, true);
        Check(result.Error == null && result.Written.SequenceEqual(new[] { output1, output2 }), "PDF save failed: " + result.Error);
        Output(output1, 370, 360, 450, 595);
        Output(output2, 350, 350, 450, 460);
        using (var pdf = PdfReader.Open(output2, PdfDocumentOpenMode.Import))
            Check(pdf.Pages[0].Rotate == 90 && pdf.Pages[1].Rotate == 270, "A rotated duplicate changed the other occurrence.");
        Check(original.SequenceEqual(File.ReadAllBytes(first)), "Save changed the source PDF.");
        Check(Directory.GetFiles(folder, "*.tmp").Length == 0, "Save left temporary files.");

        var conflict = Step(def, PipelineStepKind.SavePdf, new SavePdfSettings { Path = output2 }, a, "In");
        result = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, true);
        Check(result.Error != null && result.Error.Contains("both write") && result.Error.Contains(output2) &&
            result.Written.Count == 0, "Numbered output collision was not caught before writing: " + result.Error);
        PipelineWork.RemoveNode(def, conflict.Id);
        var overwrite = Step(def, PipelineStepKind.SavePdf, new SavePdfSettings { Path = first }, a, "In");
        result = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, true);
        Check(result.FailedNodeId == overwrite.Id && result.Error.Contains("source PDF") &&
            original.SequenceEqual(File.ReadAllBytes(first)), "A save could overwrite a source PDF.");
    }

    static void BatchAndFailureChecks(string folder, string first, string second)
    {
        var def = new PipelineDefinition();
        var fileList = Step(def, PipelineStepKind.ValueList,
            new ValueListSettings { Source = ValueListSource.Stored, Mode = ValueListMode.List }, null, null);
        var values = (ValueListSettings)fileList.Settings;
        values.Items.Add(new ValueListItem { Key = first });
        values.Items.Add(new ValueListItem { Key = second });
        var load = Step(def, PipelineStepKind.LoadPdf, new LoadPdfSettings { PathColumn = "Key" }, fileList, "In");
        var merge = Step(def, PipelineStepKind.MergePdf,
            new MergePdfSettings { Interleave = true, ReverseSecond = true }, load, "In1");
        string interleaved = Path.Combine(folder, "interleaved.pdf");
        var save = Step(def, PipelineStepKind.SavePdf, new SavePdfSettings { Path = interleaved }, merge, "In");
        var cache = new Dictionary<string, DataTable>();
        var result = PipelineWork.Run(def, cache, null, true);
        Check(result.Error == null, result.Error);
        Output(interleaved, 350, 460, 360, 450, 370);
        Check(RunnerWork.Inputs(def).Count == 1, "Linked Load PDF showed an unnecessary file picker.");

        var split = PdfWork.Split(cache[load.Id], new SplitPdfSettings { PagesPerFile = 2 });
        Check(split.Rows.Count == 5 && (int)split.Rows[4]["Document"] == 3, "Batch splitting mixed pages from different documents.");
        var bad = Step(def, PipelineStepKind.ArrangePdf, new ArrangePdfSettings { Pages = "99" }, merge, "A");
        byte[] before = File.ReadAllBytes(interleaved);
        result = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, true);
        Check(result.FailedNodeId == bad.Id && result.Error.Contains("outside the range") &&
            result.Written.Count == 0 && before.SequenceEqual(File.ReadAllBytes(interleaved)), "An invalid page range allowed a save.");
        PipelineWork.RemoveNode(def, bad.Id);

        cache.Clear();
        result = PipelineWork.Run(def, cache, null, false);
        Check(result.Error == null && cache.ContainsKey(save.Id), "Could not preview before checking source changes.");
        File.AppendAllText(second, "\n% changed after preview\n");
        result = PipelineWork.Run(def, cache, null, true);
        Check(result.FailedNodeId == save.Id && result.Error.Contains("source PDF changed") &&
            result.Written.Count == 0 && before.SequenceEqual(File.ReadAllBytes(interleaved)) &&
            Directory.GetFiles(folder, "*.tmp").Length == 0, "Source changes damaged an existing output.");
    }

    static void SeparatePdfResults(string folder, string first, string second)
    {
        var def = new PipelineDefinition();
        var a = Step(def, PipelineStepKind.LoadPdf, new LoadPdfSettings { Path = first }, null, null);
        var b = Step(def, PipelineStepKind.LoadPdf, new LoadPdfSettings { Path = second }, null, null);
        var split = Step(def, PipelineStepKind.SplitPdf, new SplitPdfSettings { PagesPerFile = 2 }, a, "In");
        Link(def, b, split, "In2");
        var save = Step(def, PipelineStepKind.SavePdf, new SavePdfSettings { Path = Path.Combine(folder, "parts.pdf") }, split, "In");
        var expected = new[] { "parts-first-001.pdf", "parts-first-002.pdf", "parts-second.pdf" }
            .Select(name => Path.Combine(folder, name)).ToArray();
        var cache = new Dictionary<string, DataTable>();
        var check = PipelineWork.Run(def, cache, null, false);
        Check(check.Error == null && PipelineWork.SaveTargetPaths(save, cache).SequenceEqual(expected) &&
            expected.All(path => !File.Exists(path)), "Preview did not plan the separate split PDF outputs.");
        var run = PipelineWork.Run(def, cache, null, true);
        Check(run.Error == null && run.Written.SequenceEqual(expected), "Writing separate PDFs failed: " + run.Error);
        Output(expected[0], 350, 360);
        Output(expected[1], 370);
        Output(expected[2], 450, 460);
    }

    [STAThread]
    static int Main()
    {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeWorkAssistant-Pdf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            string first = Path.Combine(folder, "first.pdf");
            string second = Path.Combine(folder, "second.pdf");
            CreatePdf(first, 350, 360, 370);
            CreatePdf(second, 450, 460);
            TestPipeline(folder, first, second);
            SeparatePdfResults(folder, first, second);
            BatchAndFailureChecks(folder, first, second);
            Console.WriteLine("Passed {0} PDF checks in a {1}-bit process.", _checks, IntPtr.Size * 8);
            Directory.Delete(folder, true);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Console.Error.WriteLine("Fixtures kept in " + folder);
            return 1;
        }
    }
}

