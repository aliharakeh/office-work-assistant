using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeWorkAssistant.Excel;
using Tesseract;

namespace OfficeWorkAssistant.Features.Ocr
{
    public sealed class OcrSettings
    {
        public string PathColumn { get; set; }
        public string TextHeader { get; set; }
        public string Languages { get; set; }
        // Blank = tessdata beside the executable. Relative folders are relative to the executable.
        public string DataFolder { get; set; }

        public OcrSettings()
        {
            PathColumn = "";
            TextHeader = "OCR text";
            Languages = "eng";
            DataFolder = "";
        }
    }

    // Reads images without changing source files. No WPF or references to other features.
    public static class OcrWork
    {
        public static void Validate(DataTable input, OcrSettings settings)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input with image paths first.");
            if (settings == null)
                throw new InvalidOperationException("Set up OCR first.");
            ExcelColumns.Index(input, settings.PathColumn);
            if (string.IsNullOrWhiteSpace(settings.TextHeader))
                throw new InvalidOperationException("Enter a name for the OCR text column.");
            foreach (DataColumn column in input.Columns)
            {
                if (string.Equals(ExcelFile.Header(column), settings.TextHeader.Trim(), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The OCR text column already exists. Choose a new column name.");
            }
            if (string.IsNullOrWhiteSpace(settings.Languages) ||
                !Regex.IsMatch(settings.Languages.Trim(), @"\A[A-Za-z0-9_]+(\+[A-Za-z0-9_]+)*\z"))
                throw new InvalidOperationException("Enter language codes such as eng, ara, or eng+ara.");
        }

        public static DataTable Run(DataTable input, OcrSettings settings)
        {
            Validate(input, settings);
            int pathColumn = ExcelColumns.Index(input, settings.PathColumn);
            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            foreach (DataColumn column in input.Columns)
                ExcelFile.AddColumn(output, ExcelFile.Header(column));
            var textColumn = ExcelFile.AddColumn(output, settings.TextHeader);

            TesseractEngine engine = null;
            try
            {
                for (var r = 0; r < input.Rows.Count; r++)
                {
                    var source = input.Rows[r];
                    var row = output.NewRow();
                    for (var c = 0; c < input.Columns.Count; c++)
                        row[c] = source[c];
                    string path = Convert.ToString(source[pathColumn], CultureInfo.InvariantCulture).Trim();
                    if (path.Length > 0)
                    {
                        try
                        {
                            string extension = Path.GetExtension(path).ToLowerInvariant();
                            if (extension != ".png" && extension != ".jpg" && extension != ".jpeg" &&
                                extension != ".bmp" && extension != ".tif" && extension != ".tiff")
                                throw new InvalidOperationException("Use PNG, JPEG, BMP or TIFF images. PDFs must be converted to images first.");
                            if (!Path.IsPathRooted(path) || Path.GetPathRoot(path).Length < 3)
                                throw new InvalidOperationException("Use a full image path, for example from List folder.");
                            if (!File.Exists(path))
                                throw new FileNotFoundException("Image not found or not accessible.", path);
                            if (engine == null)
                                engine = OpenEngine(settings);
                            row[textColumn] = ReadImage(engine, path, extension);
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException("OCR row " + (r + 1).ToString(CultureInfo.InvariantCulture) +
                                " (" + path + "): " + ex.Message, ex);
                        }
                    }
                    output.Rows.Add(row);
                }
            }
            finally
            {
                if (engine != null)
                    engine.Dispose();
            }
            return output;
        }

        static string DataPath(string dataFolder)
        {
            string folder = string.IsNullOrWhiteSpace(dataFolder) ? "tessdata" : dataFolder.Trim();
            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, folder));
        }

        // Installed languages (display name, code), then every pair of them joined with "+".
        public static List<KeyValuePair<string, string>> LanguageOptions(string dataFolder)
        {
            var codes = new List<string>();
            string folder = DataPath(dataFolder);
            if (Directory.Exists(folder))
            {
                foreach (string file in Directory.GetFiles(folder, "*.traineddata"))
                {
                    string code = Path.GetFileNameWithoutExtension(file);
                    if (code != "osd" && Regex.IsMatch(code, @"\A[A-Za-z0-9_]+\z"))
                        codes.Add(code);
                }
            }
            codes.Sort(StringComparer.OrdinalIgnoreCase);
            var list = new List<KeyValuePair<string, string>>();
            foreach (string code in codes)
                list.Add(new KeyValuePair<string, string>(LanguageName(code), code));
            for (var i = 0; i < codes.Count; i++)
                for (var j = i + 1; j < codes.Count; j++)
                    list.Add(new KeyValuePair<string, string>(
                        LanguageName(codes[i]) + " + " + LanguageName(codes[j]), codes[i] + "+" + codes[j]));
            return list;
        }

        static string LanguageName(string code)
        {
            if (code == "eng") return "English (eng)";
            if (code == "ara") return "Arabic (ara)";
            return code;
        }

        static TesseractEngine OpenEngine(OcrSettings settings)
        {
            string folder = DataPath(settings.DataFolder);
            foreach (string language in settings.Languages.Trim().Split('+'))
            {
                string model = Path.Combine(folder, language + ".traineddata");
                if (!File.Exists(model))
                    throw new InvalidOperationException("OCR language data is missing: " + model +
                        ". Add the matching tessdata_fast or tessdata_best language file to this folder.");
            }
            try
            {
                // The bundled models support LSTM only, not the legacy engine.
                return new TesseractEngine(folder, settings.Languages.Trim(), EngineMode.LstmOnly);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not start Tesseract OCR. Check the language files, the x86/x64 native DLL folders " +
                    "beside the app, and the Visual C++ 2019 runtime matching the app's architecture. " +
                    "Windows 7 compatibility has not been verified. Details: " + ex.Message, ex);
            }
        }

        static string ReadImage(TesseractEngine engine, string path, string extension)
        {
            // Managed file reads also work for paths the native library's filename encoding cannot represent.
            byte[] bytes = File.ReadAllBytes(path);
            var text = new StringBuilder();
            using (var stream = new MemoryStream(bytes))
            using (var image = System.Drawing.Image.FromStream(stream))
            {
                int pages = image.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);
                for (var i = 0; i < pages; i++)
                {
                    image.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, i);
                    using (var prepared = Prepare(image))
                    using (var pageStream = new MemoryStream())
                    {
                        prepared.Save(pageStream, System.Drawing.Imaging.ImageFormat.Png);
                        using (var page = Pix.LoadFromMemory(pageStream.ToArray()))
                        {
                            if (i > 0)
                                text.Append("\n\n");
                            text.Append(ReadPage(engine, page));
                        }
                    }
                }
            }
            return text.ToString();
        }

        // Transparent pixels become white (they would read as black), and small images are enlarged,
        // since Arabic needs roughly 30+ px letter height to be recognized reliably.
        static System.Drawing.Bitmap Prepare(System.Drawing.Image source)
        {
            double scale = 1;
            if (source.Width < 1600 && source.Height < 1600)
                scale = Math.Min(3.0, 1600.0 / Math.Max(source.Width, source.Height));
            if (scale < 1.25)
                scale = 1;
            int width = Math.Max(1, (int)Math.Round(source.Width * scale));
            int height = Math.Max(1, (int)Math.Round(source.Height * scale));
            var result = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            result.SetResolution(300, 300);
            using (var g = System.Drawing.Graphics.FromImage(result))
            {
                g.Clear(System.Drawing.Color.White);
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(source, new System.Drawing.Rectangle(0, 0, width, height));
            }
            return result;
        }

        static string ReadPage(TesseractEngine engine, Pix image)
        {
            // Automatic layout analysis sometimes finds no text blocks; retry as one block, then sparse text.
            foreach (var mode in new[] { PageSegMode.Auto, PageSegMode.SingleBlock, PageSegMode.SparseText })
            {
                using (var page = engine.Process(image, mode))
                {
                    string text = (page.GetText() ?? "").TrimEnd('\r', '\n');
                    if (text.Trim().Length > 0)
                        return text;
                }
            }
            return "";
        }
    }
}
