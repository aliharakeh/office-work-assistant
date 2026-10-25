using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.Ocr;
using OfficeWorkAssistant.Features.LoadImage;
using OfficeWorkAssistant.Features.Pipeline;

// Standalone Framework console checks; RunOcrTests.ps1 compiles this beside the Release build.
internal static class OcrSmokeTests
{
    static int _checks;

    static void Check(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
        _checks++;
    }

    static void Fails(Action action, string message)
    {
        try { action(); }
        catch (InvalidOperationException ex)
        {
            Check(ex.Message.IndexOf(message, StringComparison.OrdinalIgnoreCase) >= 0,
                "Expected error containing '" + message + "', got: " + ex.Message);
            return;
        }
        throw new Exception("Expected failure: " + message);
    }

    static DataTable Input(string path)
    {
        var table = new DataTable();
        ExcelFile.AddColumn(table, "Image path (full)");
        ExcelFile.AddColumn(table, "Item / number");
        table.Rows.Add(path, 42);
        return table;
    }

    static OcrSettings Settings()
    {
        return new OcrSettings { PathColumn = "A" };
    }

    static Bitmap Image(string heading)
    {
        var image = new Bitmap(1600, 320);
        image.SetResolution(300, 300);
        using (var graphics = Graphics.FromImage(image))
        using (var font = new Font("Arial", 48, FontStyle.Regular, GraphicsUnit.Pixel))
        {
            graphics.Clear(Color.White);
            graphics.DrawString(heading, font, Brushes.Black, 50, 50);
            graphics.DrawString("Invoice number 12345", font, Brushes.Black, 50, 150);
        }
        return image;
    }

    static void MultiPageTiff(string path)
    {
        ImageCodecInfo encoder = null;
        foreach (var codec in ImageCodecInfo.GetImageEncoders())
        {
            if (codec.FormatID == ImageFormat.Tiff.Guid)
                encoder = codec;
        }
        using (var first = Image("OFFICE OCR TEST"))
        using (var second = Image("SECOND PAGE 67890"))
        using (var parameters = new EncoderParameters(1))
        {
            using (var flag = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame))
            {
                parameters.Param[0] = flag;
                first.Save(path, encoder, parameters);
            }
            using (var flag = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage))
            {
                parameters.Param[0] = flag;
                first.SaveAdd(second, parameters);
            }
            using (var flag = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush))
            {
                parameters.Param[0] = flag;
                first.SaveAdd(parameters);
            }
        }
    }

    static void LogicChecks()
    {
        Fails(() => LoadImageWork.Run(null), "Choose an image");
        Fails(() => LoadImageWork.Run(new LoadImageSettings()), "Choose an image");
        Fails(() => LoadImageWork.Run(new LoadImageSettings { Path = "relative.png" }), "full image path");
        Fails(() => LoadImageWork.Run(new LoadImageSettings { Path = @"C:relative.png" }), "full image path");
        Fails(() => LoadImageWork.Run(new LoadImageSettings { Path = @"\relative.png" }), "full image path");
        Fails(() => LoadImageWork.Run(new LoadImageSettings { Path = @"C:\missing.pdf" }), "PNG, JPEG, BMP or TIFF");
        Fails(() => LoadImageWork.Run(new LoadImageSettings { Path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png") }),
            "Image not found");
        Check(PipelineWork.Ports(PipelineStepKind.LoadImage).Length == 0, "Load image must be a source.");
        Check(PipelineWork.IsFileKind(PipelineStepKind.LoadImage) && !PipelineWork.IsActionKind(PipelineStepKind.LoadImage),
            "Load image must not be a disk action.");
        var input = Input("");
        input.Rows.Add(DBNull.Value, 7);
        var settings = Settings();
        var output = OcrWork.Run(input, settings);
        Check(output.Rows.Count == 2 && output.Columns.Count == 3, "Output shape changed.");
        Check(output.Rows[0].IsNull(2) && output.Rows[1].IsNull(2), "Blank paths must stay blank.");
        Check(input.Columns.Count == 2 && (int)input.Rows[0][1] == 42, "Input was mutated.");
        Check(ExcelFile.Header(output.Columns[0]) == "Image path (full)", "Caption was lost.");
        Fails(() => OcrWork.Run(null, settings), "Connect an input");
        Fails(() => OcrWork.Run(input, null), "Set up OCR");
        settings.PathColumn = "Z";
        Fails(() => OcrWork.Validate(input, settings), "not in this step");
        settings = Settings();
        settings.TextHeader = "image PATH (full)";
        Fails(() => OcrWork.Validate(input, settings), "already exists");
        settings.TextHeader = " ";
        Fails(() => OcrWork.Validate(input, settings), "Enter a name");
        settings = Settings();
        settings.Languages = "../eng";
        Fails(() => OcrWork.Validate(input, settings), "language codes");
        settings.Languages = "eng+";
        Fails(() => OcrWork.Validate(input, settings), "language codes");
        settings.Languages = "eng+ara";
        OcrWork.Validate(input, settings);
        settings = Settings();
        Fails(() => OcrWork.Run(Input("relative.png"), settings), "full image path");
        Fails(() => OcrWork.Run(Input(@"C:relative.png"), settings), "full image path");
        Fails(() => OcrWork.Run(Input(@"\relative.png"), settings), "full image path");
        Fails(() => OcrWork.Run(Input(@"C:\missing.pdf"), settings), "PDFs must be converted");
        Fails(() => OcrWork.Run(Input(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png")), settings), "OCR row 1");

        var names = PipelineWork.OcrToNames(settings, input);
        Check(names.PathColumn == "Image path (full)" && settings.PathColumn == "A", "Name conversion mutated settings.");
        var reordered = new DataTable();
        ExcelFile.AddColumn(reordered, "New first column");
        ExcelFile.AddColumn(reordered, "Image path (full)");
        Check(PipelineWork.OcrToLetters(names, reordered, true).PathColumn == "B", "Header reference did not survive reorder.");
        Fails(() => PipelineWork.OcrToLetters(names, new DataTable(), true), "Image path (full)");
        Check(PipelineWork.OcrToLetters(names, new DataTable(), false).PathColumn == "", "Missing editor column was silently retargeted.");

        var definition = new PipelineDefinition();
        var source = PipelineWork.AddNode(definition, PipelineStepKind.Load, 0, 0);
        var ocr = PipelineWork.AddNode(definition, PipelineStepKind.Ocr, 200, 0);
        ocr.Settings = names;
        definition.Links.Add(new PipelineLink { From = source.Id, To = ocr.Id, Port = PipelineWork.PortIn });
        var saved = PipelineWork.Clone(definition);
        Check(saved.Nodes[1].Kind == PipelineStepKind.Ocr && saved.Nodes[1].Settings is OcrSettings, "Pipeline XML round-trip failed.");
        Check(PipelineWork.IsFileKind(PipelineStepKind.Ocr) && !PipelineWork.IsActionKind(PipelineStepKind.Ocr), "OCR must not be a disk action.");
        Check(PipelineWork.Ports(PipelineStepKind.Ocr).Length == 1, "OCR needs one input.");
        var cache = new Dictionary<string, DataTable> { { source.Id, input } };
        var result = PipelineWork.Run(saved, cache, ocr.Id, false);
        Check(result.Error == null && cache[ocr.Id].Columns.Count == 3, "Pipeline OCR dispatch failed: " + result.Error);
    }

    static void ArabicChecks(string folder)
    {
        string arabic = "\u0627\u0644\u0644\u063a\u0629 \u0627\u0644\u0639\u0631\u0628\u064a\u0629";
        string word = "\u0627\u0644\u0639\u0631\u0628\u064a\u0629";
        for (var mixed = 0; mixed < 2; mixed++)
        {
            string path = Path.Combine(folder, "arabic-" + mixed + ".png");
            using (var image = new Bitmap(1600, 320))
            using (var graphics = Graphics.FromImage(image))
            using (var font = new Font("Arial", 64, FontStyle.Regular, GraphicsUnit.Pixel))
            using (var rtl = new StringFormat(StringFormatFlags.DirectionRightToLeft))
            {
                image.SetResolution(300, 300);
                graphics.Clear(Color.White);
                graphics.DrawString(arabic, font, Brushes.Black, new RectangleF(50, 30, 1500, 110), rtl);
                if (mixed == 1)
                    graphics.DrawString("OFFICE OCR TEST", font, Brushes.Black, 50, 170);
                else
                    graphics.DrawString(arabic, font, Brushes.Black, new RectangleF(50, 170, 1500, 110), rtl);
                image.Save(path, ImageFormat.Png);
            }
            var settings = Settings();
            settings.Languages = mixed == 0 ? "ara" : "eng+ara";
            var output = OcrWork.Run(Input(path), settings);
            string text = Convert.ToString(output.Rows[0][2]);
            var codePoints = new System.Text.StringBuilder();
            foreach (char character in text)
                codePoints.Append("\\u").Append(((int)character).ToString("x4"));
            Check(text.Contains(word), "Arabic OCR failed for " + settings.Languages + ": " + codePoints);
            if (mixed == 1)
                Check(text.Contains("OFFICE"), "Mixed-language OCR lost English: " + text);
            string workbook = Path.Combine(folder, "arabic-" + mixed + ".xlsx");
            ExcelFile.Save(output, workbook);
            Check(Convert.ToString(ExcelFile.Load(workbook, null).Table.Rows[0][2]) == text,
                "Arabic text did not survive Excel save/load.");
        }
    }

    static void ImageSourceChecks(string path, string folder)
    {
        var settings = new LoadImageSettings { Path = path };
        var input = LoadImageWork.Run(settings);
        Check(input.Rows.Count == 1 && input.Columns.Count == 2, "Load image output shape changed.");
        Check(ExcelFile.Header(input.Columns[0]) == "Name" && ExcelFile.Header(input.Columns[1]) == "Path",
            "Load image headers must match the OCR input convention.");
        Check((string)input.Rows[0][0] == Path.GetFileName(path) && (string)input.Rows[0][1] == path,
            "Load image lost the filename or Unicode path.");
        Check(settings.Path == path, "Load image mutated its settings.");
        input.Rows.Clear();
        Check(LoadImageWork.Run(settings).Rows.Count == 1, "Load image reused a mutable result table.");

        var definition = new PipelineDefinition();
        var source = PipelineWork.AddNode(definition, PipelineStepKind.LoadImage, 0, 0);
        Check(source.Settings is LoadImageSettings, "Load image settings were not initialized.");
        source.Settings = settings;
        Check(PipelineWork.Summary(source) == Path.GetFileName(path), "Load image summary is missing.");
        var ocr = PipelineWork.AddNode(definition, PipelineStepKind.Ocr, 200, 0);
        ocr.Settings = new OcrSettings { PathColumn = "Path" };
        var save = PipelineWork.AddNode(definition, PipelineStepKind.Save, 400, 0);
        string workbook = Path.Combine(folder, "load-image-pipeline.xlsx");
        save.Settings = new SaveStepSettings { Path = workbook };
        definition.Links.Add(new PipelineLink { From = source.Id, To = ocr.Id, Port = PipelineWork.PortIn });
        definition.Links.Add(new PipelineLink { From = ocr.Id, To = save.Id, Port = PipelineWork.PortIn });
        var saved = PipelineWork.Clone(definition);
        Check(saved.Nodes[0].Kind == PipelineStepKind.LoadImage && ((LoadImageSettings)saved.Nodes[0].Settings).Path == path,
            "Load image pipeline XML round-trip failed.");
        var cache = new Dictionary<string, DataTable>();
        var result = PipelineWork.Run(saved, cache, null, false);
        Check(result.Error == null, "Load image -> OCR -> Save preview failed: " + result.Error);
        Check(!File.Exists(workbook), "Load image pipeline preview wrote an Excel file.");
        Check(Convert.ToString(cache[ocr.Id].Rows[0][2]).Contains("OFFICE"), "Load image did not feed OCR.");
    }

    static void NativeChecks(string folder)
    {
        var formats = new[] { ImageFormat.Png, ImageFormat.Jpeg, ImageFormat.Bmp, ImageFormat.Tiff };
        var extensions = new[] { ".png", ".jpg", ".bmp", ".tiff" };
        for (var i = 0; i < formats.Length; i++)
        {
            string path = Path.Combine(folder, "invoice-\u0639\u0631\u0628\u064a-" + i + extensions[i]);
            using (var image = Image("OFFICE OCR TEST"))
                image.Save(path, formats[i]);
            byte[] original = File.ReadAllBytes(path);
            var input = Input(path);
            var output = OcrWork.Run(input, Settings());
            string text = Convert.ToString(output.Rows[0][2]);
            Check(text.Contains("OFFICE") && text.Contains("12345"), "OCR failed for " + extensions[i] + ": " + text);
            Check(input.Columns.Count == 2 && (string)input.Rows[0][0] == path, "Native run changed its input.");
            Check(Convert.ToBase64String(original) == Convert.ToBase64String(File.ReadAllBytes(path)), "OCR changed the image file.");
            if (i == 0)
                ImageSourceChecks(path, folder);
        }

        string multi = Path.Combine(folder, "two-pages.tif");
        MultiPageTiff(multi);
        var both = OcrWork.Run(Input(multi), Settings());
        string combined = Convert.ToString(both.Rows[0][2]);
        Check(combined.Contains("OFFICE") && combined.Contains("SECOND") && combined.Contains("67890"), "TIFF pages were lost: " + combined);
        var batch = Input(multi);
        batch.Rows.Add("", 43);
        batch.Rows.Add(multi, 44);
        var batchOutput = OcrWork.Run(batch, Settings());
        Check(batchOutput.Rows[1].IsNull(2) && Convert.ToString(batchOutput.Rows[2][2]) == combined,
            "Engine reuse or blank rows changed batch results.");
        string workbook = Path.Combine(folder, "ocr.xlsx");
        ExcelFile.Save(both, workbook);
        Check(Convert.ToString(ExcelFile.Load(workbook, null).Table.Rows[0][2]) == combined, "OCR text did not survive Excel save/load.");

        var missing = Settings();
        missing.Languages = "not_installed";
        Fails(() => OcrWork.Run(Input(multi), missing), "not_installed.traineddata");
        var corrupt = Path.Combine(folder, "corrupt.png");
        File.WriteAllText(corrupt, "Not an image");
        Fails(() => OcrWork.Run(Input(corrupt), Settings()), "OCR row 1");
        batch.Rows.Add(corrupt, 45);
        Fails(() => OcrWork.Run(batch, Settings()), "OCR row 4");
        ArabicChecks(folder);
    }

    static int Main(string[] args)
    {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeWorkAssistant-Ocr-" + Guid.NewGuid().ToString("N"));
        bool passed = false;
        try
        {
            LogicChecks();
            if (args.Length == 0 || args[0] != "--skip-native")
            {
                Directory.CreateDirectory(folder);
                NativeChecks(folder);
            }
            Console.WriteLine("Passed " + _checks + " OCR checks (" + (IntPtr.Size * 8) + "-bit process).");
            passed = true;
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Console.Error.WriteLine("Test images kept in: " + folder);
            return 1;
        }
        finally
        {
            if (passed && Directory.Exists(folder))
                Directory.Delete(folder, true);
        }
    }
}
