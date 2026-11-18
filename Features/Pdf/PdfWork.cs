using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using OfficeWorkAssistant.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace OfficeWorkAssistant.Features.Pdf
{
    public sealed class LoadPdfSettings
    {
        public string Path { get; set; }
        public string PathColumn { get; set; }

        public LoadPdfSettings()
        {
            Path = "";
            PathColumn = "";
        }
    }

    public sealed class ArrangePdfSettings
    {
        public string Pages { get; set; }

        public ArrangePdfSettings()
        {
            Pages = "all";
        }
    }

    public sealed class MergePdfSettings
    {
        public bool Interleave { get; set; }
        public bool ReverseSecond { get; set; }
    }

    public sealed class SplitPdfSettings
    {
        public int PagesPerFile { get; set; }
        public string Ranges { get; set; }

        public SplitPdfSettings()
        {
            PagesPerFile = 1;
            Ranges = "";
        }
    }

    public sealed class SavePdfSettings
    {
        public string Path { get; set; }

        public SavePdfSettings()
        {
            Path = "";
        }
    }

    public static class PdfWork
    {
        const string DocumentsKey = "OfficeWorkAssistant.Pdf.Documents";
        const string OutputsKey = "OfficeWorkAssistant.Pdf.Outputs";

        sealed class Source
        {
            public readonly string Path;
            readonly long _length;
            readonly DateTime _modified;

            public Source(string path)
            {
                var file = new FileInfo(path);
                Path = path;
                _length = file.Length;
                _modified = file.LastWriteTimeUtc;
            }

            public void Check()
            {
                var file = new FileInfo(Path);
                if (!file.Exists || file.Length != _length || file.LastWriteTimeUtc != _modified)
                    throw new InvalidOperationException("The source PDF changed. Run the preview again: " + Path);
            }
        }

        sealed class Page
        {
            public readonly Source Source;
            public readonly int Number;
            public readonly double Width;
            public readonly double Height;
            public readonly int Rotation;

            public Page(Source source, int number, double width, double height, int rotation)
            {
                Source = source;
                Number = number;
                Width = width;
                Height = height;
                Rotation = ((rotation % 360) + 360) % 360;
            }

            public Page Rotate(int degrees)
            {
                return degrees == 0 ? this : new Page(Source, Number, Width, Height, Rotation + degrees);
            }
        }

        sealed class Output
        {
            public string Path;
            public bool Replace;
            public List<Page> Pages;
        }

        public static string InputPath(string path)
        {
            path = PdfPath(path, false);
            if (!File.Exists(path))
                throw new InvalidOperationException("The PDF file does not exist: " + path);
            return path;
        }

        public static string OutputPath(SavePdfSettings settings)
        {
            string path = PdfPath(settings.Path, true);
            if (!Directory.Exists(Path.GetDirectoryName(path)))
                throw new InvalidOperationException("The output folder does not exist: " + Path.GetDirectoryName(path));
            if (Directory.Exists(path))
                throw new InvalidOperationException("Choose a PDF file, not a folder: " + path);
            return path;
        }

        static string PdfPath(string path, bool addExtension)
        {
            path = (path ?? "").Trim();
            if (!Path.IsPathRooted(path) || Path.GetPathRoot(path).Length < 3 ||
                path.StartsWith(@"\\.\", StringComparison.Ordinal) || path.StartsWith(@"\\?\", StringComparison.Ordinal))
                throw new InvalidOperationException("Use a full PDF file path, including the drive.");
            if (string.IsNullOrEmpty(Path.GetFileName(path)) || path.EndsWith(".", StringComparison.Ordinal))
                throw new InvalidOperationException("Enter a PDF filename.");
            if (addExtension && string.IsNullOrEmpty(Path.GetExtension(path)))
                path += ".pdf";
            if (!string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Choose a .pdf file: " + path);
            return Path.GetFullPath(path);
        }

        public static DataTable Load(DataTable input, LoadPdfSettings settings)
        {
            var paths = new List<string>();
            if (input == null)
                paths.Add(InputPath(settings.Path));
            else
            {
                int column = ExcelColumns.Index(input, settings.PathColumn);
                foreach (DataRow row in input.Rows)
                    paths.Add(InputPath(Convert.ToString(row[column], CultureInfo.InvariantCulture)));
            }
            if (paths.Count == 0)
                throw new InvalidOperationException("The input has no PDF paths.");

            var documents = new List<List<Page>>();
            foreach (string path in paths)
            {
                var source = new Source(path);
                using (var document = Open(source))
                {
                    var pages = new List<Page>();
                    foreach (PdfPage page in document.Pages)
                        pages.Add(new Page(source, pages.Count + 1, page.Width.Point, page.Height.Point, page.Rotate));
                    if (pages.Count == 0)
                        throw new InvalidOperationException("The PDF has no pages: " + path);
                    documents.Add(pages);
                }
            }
            return PageTable(documents);
        }

        static PdfDocument Open(Source source)
        {
            source.Check();
            try
            {
                var document = PdfReader.Open(source.Path, PdfDocumentOpenMode.Import);
                if (document.SecuritySettings.IsEncrypted)
                {
                    document.Dispose();
                    throw new InvalidOperationException("Encrypted PDFs are not supported by these steps.");
                }
                return document;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not open the PDF " + source.Path + ". " + ex.Message, ex);
            }
        }

        public static DataTable Arrange(DataTable input, DataTable insert, ArrangePdfSettings settings)
        {
            var documents = Documents(input);
            var extra = insert == null ? null : Documents(insert);
            if (extra != null && (documents.Count != 1 || extra.Count != 1))
                throw new InvalidOperationException("Connect one document to each Arrange PDF pages input when inserting from B.");

            var arranged = new List<List<Page>>();
            foreach (var pages in documents)
                arranged.Add(Select(pages, extra == null ? null : extra[0], settings.Pages, true));
            return PageTable(arranged);
        }

        static List<Page> Select(List<Page> first, List<Page> second, string selection, bool allowEdits)
        {
            var result = new List<Page>();
            foreach (string entry in (selection ?? "").Split(','))
            {
                var token = entry.Trim().ToLowerInvariant();
                int rotation = 0;
                int at = token.IndexOf('@');
                if (at >= 0)
                {
                    if (!allowEdits || !int.TryParse(token.Substring(at + 1).Trim(), NumberStyles.AllowLeadingSign,
                        CultureInfo.InvariantCulture, out rotation) || rotation % 90 != 0 || rotation < -270 || rotation > 270)
                        throw new InvalidOperationException("Use @90, @180, @270 or a negative quarter turn after a page selection.");
                    token = token.Substring(0, at).Trim();
                }

                var pages = first;
                if (allowEdits && token.StartsWith("b:", StringComparison.Ordinal))
                {
                    if (second == null)
                        throw new InvalidOperationException("Connect the PDF to insert to input B.");
                    pages = second;
                    token = token.Substring(2).Trim();
                }
                else if (allowEdits && token.StartsWith("a:", StringComparison.Ordinal))
                    token = token.Substring(2).Trim();

                switch (token)
                {
                    case "blank":
                        if (!allowEdits)
                            throw new InvalidOperationException("Add blank pages with Arrange PDF pages.");
                        result.Add(new Page(null, 0, 595.28, 841.89, rotation));
                        break;
                    case "all":
                    case "odd":
                    case "even":
                        for (int i = token == "even" ? 1 : 0; i < pages.Count; i += token == "all" ? 1 : 2)
                            result.Add(pages[i].Rotate(rotation));
                        break;
                    default:
                        var range = token.Split('-');
                        if (range.Length > 2)
                            throw new InvalidOperationException("Invalid page range: " + entry.Trim());
                        int from = PageNumber(range[0], pages.Count);
                        int to = range.Length == 1 ? from : PageNumber(range[1], pages.Count);
                        int direction = from <= to ? 1 : -1;
                        for (int i = from; i != to + direction; i += direction)
                            result.Add(pages[i - 1].Rotate(rotation));
                        break;
                }
            }
            if (result.Count == 0)
                throw new InvalidOperationException("The selection must include at least one page.");
            return result;
        }

        static int PageNumber(string text, int last)
        {
            text = text.Trim();
            int page;
            if (text == "last")
                return last;
            if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out page) || page < 1 || page > last)
                throw new InvalidOperationException("Page '" + text + "' is outside the range 1 to " + last + ".");
            return page;
        }

        public static DataTable Merge(IList<DataTable> inputs, MergePdfSettings settings)
        {
            var documents = new List<List<Page>>();
            foreach (var input in inputs)
                documents.AddRange(Documents(input));
            if (documents.Count < 2)
                throw new InvalidOperationException("Connect at least two PDF documents, or a Load PDF step that reads several files.");

            var pages = new List<Page>();
            if (settings.Interleave)
            {
                if (documents.Count != 2)
                    throw new InvalidOperationException("Interleave needs exactly two PDF documents.");
                var first = documents[0];
                var second = documents[1];
                for (int i = 0; i < Math.Max(first.Count, second.Count); i++)
                {
                    if (i < first.Count)
                        pages.Add(first[i]);
                    if (i < second.Count)
                        pages.Add(second[settings.ReverseSecond ? second.Count - 1 - i : i]);
                }
            }
            else
            {
                foreach (var document in documents)
                    pages.AddRange(document);
            }
            return PageTable(new List<List<Page>> { pages });
        }

        public static DataTable Split(DataTable input, SplitPdfSettings settings)
        {
            var split = new List<List<Page>>();
            var ranges = (settings.Ranges ?? "").Trim();
            if (ranges.Length == 0 && settings.PagesPerFile < 1)
                throw new InvalidOperationException("Enter at least one page per output file.");
            foreach (var document in Documents(input))
            {
                if (ranges.Length > 0)
                {
                    foreach (string range in ranges.Split(';'))
                        split.Add(Select(document, null, range, false));
                    continue;
                }
                for (int start = 0; start < document.Count;)
                {
                    int count = Math.Min(settings.PagesPerFile, document.Count - start);
                    split.Add(document.GetRange(start, count));
                    start += count;
                }
            }
            return PageTable(split);
        }

        static List<List<Page>> Documents(DataTable input)
        {
            var documents = input == null ? null : input.ExtendedProperties[DocumentsKey] as List<List<Page>>;
            if (documents == null)
                throw new InvalidOperationException("Connect a Load PDF, Arrange PDF pages, Merge PDFs or Split PDF step.");
            return documents;
        }

        static DataTable PageTable(List<List<Page>> documents)
        {
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            foreach (string header in new[] { "Document", "Page", "Source", "Source page", "Rotation", "Width (pt)", "Height (pt)" })
                ExcelFile.AddColumn(table, header);
            for (int i = 0; i < documents.Count; i++)
            {
                int position = 0;
                foreach (var page in documents[i])
                    table.Rows.Add(i + 1, ++position, page.Source == null ? "(blank A4)" : page.Source.Path,
                        page.Source == null ? (object)DBNull.Value : page.Number, page.Rotation, page.Width, page.Height);
            }
            table.ExtendedProperties[DocumentsKey] = documents;
            return table;
        }

        // A generic table step may change rows without updating the PDF document list.
        public static void ClearDocuments(DataTable table)
        {
            if (table != null && table.ExtendedProperties.ContainsKey(DocumentsKey))
                table.ExtendedProperties.Remove(DocumentsKey);
        }

        public static DataTable PlanSave(DataTable input, SavePdfSettings settings)
        {
            var documents = Documents(input);
            string path = OutputPath(settings);
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            foreach (string header in new[] { "Path", "Pages", "Status" })
                ExcelFile.AddColumn(table, header);

            var outputs = new List<Output>();
            int digits = Math.Max(3, documents.Count.ToString(CultureInfo.InvariantCulture).Length);
            for (int i = 0; i < documents.Count; i++)
            {
                string target = documents.Count == 1 ? path : Path.Combine(Path.GetDirectoryName(path),
                    Path.GetFileNameWithoutExtension(path) + "-" + (i + 1).ToString("D" + digits, CultureInfo.InvariantCulture) + ".pdf");
                if (Directory.Exists(target))
                    throw new InvalidOperationException("A folder occupies the PDF output path: " + target);
                var output = new Output { Path = target, Pages = documents[i], Replace = File.Exists(target) };
                outputs.Add(output);
                table.Rows.Add(target, output.Pages.Count, output.Replace ? "Overwrite" : "Ready");
            }
            table.ExtendedProperties[OutputsKey] = outputs;
            return table;
        }

        public static IEnumerable<string> Sources(DataTable table)
        {
            var documents = table.ExtendedProperties[DocumentsKey] as List<List<Page>>;
            if (documents == null)
                yield break;
            foreach (var pages in documents)
                foreach (var page in pages)
                    if (page.Source != null)
                        yield return page.Source.Path;
        }

        public static IEnumerable<string> Targets(DataTable plan, bool replacementsOnly)
        {
            var outputs = plan.ExtendedProperties[OutputsKey] as List<Output>;
            if (outputs == null)
                throw new InvalidOperationException("Run the PDF preview before saving.");
            foreach (var output in outputs)
                if (!replacementsOnly || output.Replace)
                    yield return output.Path;
        }

        // Generate every part before publishing any part. Each replacement is atomic.
        public static DataTable Write(DataTable plan, IList<string> written)
        {
            var outputs = (List<Output>)plan.ExtendedProperties[OutputsKey];
            var staged = new List<string>();
            var opened = new Dictionary<Source, PdfDocument>();
            try
            {
                foreach (var output in outputs)
                {
                    string temp = Path.Combine(Path.GetDirectoryName(output.Path), "." + Guid.NewGuid().ToString("N") + ".tmp");
                    using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        staged.Add(temp);
                        WritePages(output.Pages, stream, opened);
                    }
                }
                var result = plan.Copy();
                for (int i = 0; i < outputs.Count; i++)
                {
                    var output = outputs[i];
                    if (output.Replace && File.Exists(output.Path))
                        File.Replace(staged[i], output.Path, null);
                    else
                        File.Move(staged[i], output.Path);
                    written.Add(output.Path);
                    result.Rows[i]["Status"] = "Saved";
                }
                return result;
            }
            finally
            {
                foreach (var source in opened.Values)
                    source.Dispose();
                foreach (string temp in staged)
                {
                    if (File.Exists(temp))
                        File.Delete(temp);
                }
            }
        }

        static void WritePages(List<Page> pages, Stream stream, IDictionary<Source, PdfDocument> opened)
        {
            using (var result = new PdfDocument())
            {
                foreach (var page in pages)
                {
                    PdfPage added;
                    if (page.Source == null)
                    {
                        added = result.AddPage();
                        added.Width = XUnit.FromPoint(page.Width);
                        added.Height = XUnit.FromPoint(page.Height);
                    }
                    else
                    {
                        PdfDocument source;
                        if (!opened.TryGetValue(page.Source, out source))
                        {
                            source = Open(page.Source);
                            opened.Add(page.Source, source);
                        }
                        result.SetRequiredVersion(source.Version);
                        added = result.AddPage(source.Pages[page.Number - 1]);
                    }
                    added.Rotate = page.Rotation;
                }
                result.Save(stream, false);
            }
        }
    }
}

