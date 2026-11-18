using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml.Serialization;
using ClosedXML.Excel;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.Append;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FormatSheet;
using OfficeWorkAssistant.Features.Highlight;
using OfficeWorkAssistant.Features.Pipeline;
using OfficeWorkAssistant.Features.SaveText;

internal static class FormatSheetSmokeTests
{
    static void Check(bool ok, string message)
    {
        if (!ok) throw new Exception(message);
    }

    static void Main()
    {
        var table = new DataTable();
        ExcelFile.AddColumn(table, "Amount");
        ExcelFile.AddColumn(table, "Amount");
        ExcelFile.AddColumn(table, "").Caption = "";
        table.Rows.Add(0, 12, DBNull.Value);
        table.Rows.Add("", 24, false);

        var s = new FormatSheetSettings
        {
            Header = new CellStyle { Bold = true, Align = CellAlign.Right },
            SheetCells = new CellStyle { Align = CellAlign.Center, VerticalAlign = CellVerticalAlign.Center, Border = CellBorder.Thin },
            BorderValuesOnly = true,
            DataRows = new CellStyle { Fill = "Light green", Wrap = true },
            BandRows = true,
            DefaultColumn = new CellStyle { NumberFormat = "0.00" },
            DefaultColumnWidth = 18
        };
        s.Columns.Add(new ColumnFormat { Column = "B", Width = 25, Style = new CellStyle { Fill = "Light blue", VerticalAlign = CellVerticalAlign.Top } });
        var saved = PipelineWork.Clone(PipelineWork.FormatSheetToNames(s, table));
        s = PipelineWork.FormatSheetToLetters(saved, table, true);
        Check(s.Columns[0].Column == "B", "The second column with the same header lost its style.");

        var highlight = new HighlightSettings();
        highlight.Rules.Add(new HighlightRule { Target = HighlightTarget.Cells, Condition = "$Value == 12",
            Columns = { "B" }, Style = new CellStyle { Fill = "Light red" } });
        var result = FormatSheetWork.Run(HighlightWork.Run(table, highlight), s);
        Check(ExcelFile.GetStyle(table) == null, "Formatting changed the input.");

        string output = Path.Combine(Path.GetTempPath(), "FormatSheet-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            ExcelFile.Save(result, output);
            using (var book = new XLWorkbook(output))
            {
                var sheet = book.Worksheet(1);
                var header = sheet.Cell("A1").Style;
                Check(header.Font.Bold && header.Alignment.Horizontal == XLAlignmentHorizontalValues.Right &&
                    header.Alignment.Vertical == XLAlignmentVerticalValues.Center, "Header style and sheet alignment did not combine.");
                var data = sheet.Cell("A2").Style;
                Check(data.Alignment.Horizontal == XLAlignmentHorizontalValues.Center && data.Alignment.WrapText &&
                    data.NumberFormat.Format == "0.00", "Sheet, row and column default styles did not combine.");
                Check(data.Fill.BackgroundColor.Equals(XLColor.FromHtml("#C6EFCE")) &&
                    sheet.Cell("A3").Style.Fill.BackgroundColor.Equals(XLColor.FromHtml("#F2F2F2")), "Data row fill or banding failed.");
                Check(sheet.Cell("B3").Style.Fill.BackgroundColor.Equals(XLColor.FromHtml("#DDEBF7")) &&
                    sheet.Cell("B3").Style.Alignment.Vertical == XLAlignmentVerticalValues.Top, "Individual column style did not override the defaults.");
                Check(sheet.Cell("B2").Style.Fill.BackgroundColor.Equals(XLColor.FromHtml("#FFC7CE")), "A column style hid a highlight.");
                Check(sheet.Column(1).Width == 18 && sheet.Column(2).Width == 25, "Default or individual width was lost.");
                Check(sheet.AutoFilter.IsEnabled, "Header filters were lost.");
                Check(data.Border.TopBorder == XLBorderStyleValues.Thin && sheet.Cell("C3").Style.Border.BottomBorder == XLBorderStyleValues.Thin,
                    "Value-only borders skipped zero or false.");
                Check(sheet.Cell("C1").Style.Border.TopBorder == XLBorderStyleValues.None &&
                    sheet.Cell("C2").Style.Border.RightBorder == XLBorderStyleValues.None &&
                    sheet.Cell("A3").Style.Border.LeftBorder == XLBorderStyleValues.None, "Value-only borders styled a blank header or data cell.");
            }

            s.BorderValuesOnly = false;
            ExcelFile.Save(FormatSheetWork.Run(table, s), output);
            using (var book = new XLWorkbook(output))
                Check(book.Worksheet(1).Cell("C2").Style.Border.RightBorder == XLBorderStyleValues.Thin, "Full-range borders skipped a blank cell.");

            var serializer = new XmlSerializer(typeof(FormatSheetSettings));
            using (var xml = new StringReader("<FormatSheetSettings><StyleHeader>false</StyleHeader><Borders>true</Borders>" +
                "<Columns><ColumnFormat><Column>A</Column><Align>Right</Align><Wrap>true</Wrap><NumberFormat>0.0</NumberFormat>" +
                "</ColumnFormat></Columns></FormatSheetSettings>"))
                s = (FormatSheetSettings)serializer.Deserialize(xml);
            ExcelFile.Save(FormatSheetWork.Run(table, s), output);
            using (var book = new XLWorkbook(output))
            {
                var sheet = book.Worksheet(1);
                var data = sheet.Cell("A2").Style;
                Check(data.Alignment.Horizontal == XLAlignmentHorizontalValues.Right && data.Alignment.WrapText &&
                    data.NumberFormat.Format == "0.0", "Legacy column formatting changed.");
                Check(data.Border.TopBorder == XLBorderStyleValues.Thin && sheet.Cell("A1").Style.Border.TopBorder == XLBorderStyleValues.None,
                    "Legacy borders changed.");
            }
            BatchPipeline();
            Console.WriteLine("Format sheet and batch pipeline smoke tests passed.");
        }
        finally { File.Delete(output); }
    }

    static void Link(PipelineDefinition pipeline, PipelineNode source, PipelineNode target, string port = PipelineWork.PortIn)
    {
        pipeline.Links.Add(new PipelineLink { From = source.Id, To = target.Id, Port = port });
    }

    static void BatchPipeline()
    {
        var folder = Path.Combine(Path.GetTempPath(), "PipelineBatch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            var pipeline = new PipelineDefinition();
            var january = PipelineWork.AddNode(pipeline, PipelineStepKind.Load, 0, 0);
            var february = PipelineWork.AddNode(pipeline, PipelineStepKind.Load, 0, 0);
            var januaryPath = Path.Combine(folder, "January.xlsx");
            var februaryPath = Path.Combine(folder, "February.xlsx");
            january.Settings = new LoadStepSettings { Path = januaryPath };
            february.Settings = new LoadStepSettings { Path = februaryPath };

            var januaryData = new DataTable();
            ExcelFile.AddColumn(januaryData, "Name");
            ExcelFile.AddColumn(januaryData, "Amount");
            januaryData.Rows.Add("First", 8);
            januaryData.Rows.Add("Second", 2);
            ExcelFile.Save(januaryData, januaryPath);
            var februaryData = new DataTable();
            ExcelFile.AddColumn(februaryData, "Amount");
            ExcelFile.AddColumn(februaryData, "Name");
            februaryData.Rows.Add(6, "Third");
            februaryData.Rows.Add(4, "Fourth");
            ExcelFile.Save(februaryData, februaryPath);

            var sort = PipelineWork.AddNode(pipeline, PipelineStepKind.FilterSort, 0, 0);
            sort.Settings = new FilterSortSettings { Sorts = { new SortKey { Column = "Amount" } } };
            Link(pipeline, january, sort);
            Check(PipelineWork.InputPorts(pipeline, sort)[1] == "In2", "A second input did not become available.");
            Link(pipeline, february, sort, "In2");
            Check(PipelineWork.InputPorts(pipeline, sort)[2] == "In3", "Another free port was not added.");

            var style = PipelineWork.AddNode(pipeline, PipelineStepKind.FormatSheet, 0, 0);
            style.Settings = new FormatSheetSettings
            {
                Header = new CellStyle { Bold = true },
                Columns = { new ColumnFormat { Column = "Amount", Style = new CellStyle { NumberFormat = "0.00" } } }
            };
            Link(pipeline, sort, style);

            var save = PipelineWork.AddNode(pipeline, PipelineStepKind.Save, 0, 0);
            var basePath = Path.Combine(folder, "result.xlsx");
            var januaryOutput = Path.Combine(folder, "result-January.xlsx");
            var februaryOutput = Path.Combine(folder, "result-February.xlsx");
            save.Settings = new SaveStepSettings { Path = basePath };
            Link(pipeline, style, save);

            var export = PipelineWork.AddNode(pipeline, PipelineStepKind.SaveText, 0, 0);
            export.Settings = new SaveTextSettings { Path = Path.Combine(folder, "names.txt"), Column = "Name" };
            Link(pipeline, sort, export);
            var append = PipelineWork.AddNode(pipeline, PipelineStepKind.Append, 0, 0);
            append.Settings = new AppendSettings { SourceHeader = "Source" };
            Link(pipeline, sort, append, "In1");

            var xml = Path.Combine(folder, "batch.xml");
            PipelineWork.SaveFile(pipeline, xml);
            pipeline = PipelineWork.LoadFile(xml);
            var cache = new Dictionary<string, DataTable>();
            var preview = PipelineWork.Run(pipeline, cache, null, false);
            Check(preview.Error == null && !File.Exists(januaryOutput) && !File.Exists(februaryOutput),
                "Batch preview failed or wrote a file: " + preview.Error);
            Check(PipelineWork.ResultItems(cache[style.Id]).Count == 2, "Formatting lost a workbook.");
            Check(cache[append.Id].Rows.Count == 4 && (string)cache[append.Id].Rows[2][0] == "February",
                "Append did not combine the batch and retain the source name.");
            Check(RunnerWork.OutputLines(pipeline, cache).Count == 4, "Run mode did not list each planned output.");

            File.WriteAllText(februaryOutput, "unconfirmed");
            var stopped = PipelineWork.Run(pipeline, cache, null, true, new string[0]);
            Check(stopped.FailedNodeId == save.Id && !File.Exists(januaryOutput) &&
                File.ReadAllText(februaryOutput) == "unconfirmed", "A new replacement skipped confirmation.");
            Check(RunnerWork.ReplacedFiles(pipeline, cache).Contains(februaryOutput), "Run mode missed a replacement.");
            var run = PipelineWork.Run(pipeline, cache, null, true, new[] { februaryOutput });
            Check(run.Error == null && run.Written.Count == 4 && !File.Exists(basePath), "Batch save failed: " + run.Error);
            CheckWorkbook(januaryOutput, "B", 2);
            CheckWorkbook(februaryOutput, "A", 4);
            Check(File.ReadAllText(Path.Combine(folder, "names-February.txt")).StartsWith("Fourth", StringComparison.Ordinal),
                "The second branch did not export the sorted February result.");

            var saveSettings = (SaveStepSettings)PipelineWork.Find(pipeline, save.Id).Settings;
            saveSettings.Mode = SaveStepMode.NewSheet;
            saveSettings.SheetName = "Formatted";
            cache.Remove(save.Id);
            run = PipelineWork.Run(pipeline, cache, save.Id, true);
            Check(run.Error == null && run.Written.Count == 2 &&
                ExcelFile.Load(januaryOutput, "Formatted").Sheets.Length == 2 &&
                ExcelFile.Load(februaryOutput, "Formatted").Sheets.Length == 2, "New sheet mode did not update each workbook.");
            var duplicateSheet = PipelineWork.Run(pipeline, cache, save.Id, false);
            Check(duplicateSheet.FailedNodeId == save.Id && duplicateSheet.Error.Contains("Formatted"),
                "Preview missed a sheet already in the destination.");

            CheckPairedInputs(pipeline, january, february);
            CheckSeparateOutputs(pipeline, sort, style, january, februaryOutput);
            februaryData.Columns[0].Caption = "Other";
            ExcelFile.Save(februaryData, februaryPath);
            var failed = PipelineWork.Run(pipeline, new Dictionary<string, DataTable>(), null, false);
            Check(failed.FailedNodeId == sort.Id && failed.Error.Contains("February") && failed.Error.Contains("Amount"),
                "A batch schema error did not identify the workbook and column.");

            PipelineWork.RemoveNode(pipeline, january.Id);
            Check(PipelineWork.InputLink(pipeline, sort.Id, PipelineWork.PortIn).From == february.Id,
                "Removing the first input broke the remaining link.");
        }
        finally
        {
            foreach (var file in Directory.GetFiles(folder))
                File.Delete(file);
            Directory.Delete(folder);
        }
    }

    static void CheckSeparateOutputs(PipelineDefinition pipeline, PipelineNode sort, PipelineNode style,
        PipelineNode january, string output)
    {
        var isolated = PipelineWork.Clone(pipeline);
        var layout = PipelineWork.ResultLayout(isolated);
        var sortPorts = PipelineWork.OutputPorts(sort, layout);
        var stylePorts = PipelineWork.OutputPorts(style, layout);
        Check(sortPorts.Length == 2 && stylePorts.Length == 2, "A separate workbook did not get its own output.");
        var compare = PipelineWork.AddNode(isolated, PipelineStepKind.Compare, 0, 0);
        Link(isolated, sort, compare, "A");
        Link(isolated, sort, compare, "B");
        Check(PipelineWork.OutputPorts(compare, PipelineWork.ResultLayout(isolated)).Length == 1,
            "Compare exposed multiple output sockets.");

        var route = PipelineWork.AddNode(isolated, PipelineStepKind.Save, 0, 0);
        string path = Path.Combine(Path.GetDirectoryName(output), "february-only.xlsx");
        route.Settings = new SaveStepSettings { Path = path };
        Link(isolated, style, route);
        PipelineWork.InputLink(isolated, route.Id, "In").FromPort = stylePorts[1];
        string xml = Path.Combine(Path.GetDirectoryName(output), "separate-output.xml");
        PipelineWork.SaveFile(isolated, xml);
        isolated = PipelineWork.LoadFile(xml);
        var cache = new Dictionary<string, DataTable>();
        var run = PipelineWork.Run(isolated, cache, route.Id, true);
        Check(run.Error == null && run.Written.Count == 1, "The selected socket did not route just one workbook: " + run.Error);
        CheckWorkbook(path, "A", 4);

        var firstRoute = PipelineWork.AddNode(isolated, PipelineStepKind.Save, 0, 0);
        Link(isolated, style, firstRoute);
        PipelineWork.InputLink(isolated, firstRoute.Id, "In").FromPort = stylePorts[0];
        PipelineWork.RemoveNode(isolated, january.Id);
        PipelineWork.RemoveDisconnectedOutputs(isolated);
        Check(PipelineWork.InputLink(isolated, firstRoute.Id, "In") == null &&
            PipelineWork.InputLink(isolated, route.Id, "In").FromPort == stylePorts[1],
            "Removing the first workbook did not preserve the second workbook's connection.");
        run = PipelineWork.Run(isolated, new Dictionary<string, DataTable>(), route.Id, false);
        Check(run.Error == null, "A surviving output could not be routed: " + run.Error);
    }

    static void CheckPairedInputs(PipelineDefinition pipeline, PipelineNode january, PipelineNode february)
    {
        var fill = PipelineWork.AddNode(pipeline, PipelineStepKind.FillColumns, 0, 0);
        fill.Settings = new FillColumnsSettings
        {
            Keys = { new FillColumnsKey { FormulaA = "1", FormulaB = "1" } },
            Rules = { new FillColumnsRule { Target = "Copied", Value = "$A_Amount" } }
        };
        Link(pipeline, january, fill, "A");
        Link(pipeline, january, fill, "B");
        Link(pipeline, february, fill, "B2");
        var cache = new Dictionary<string, DataTable>();
        var run = PipelineWork.Run(pipeline, cache, fill.Id, false);
        Check(run.Error == null, "Reusing one lookup for both workbooks failed: " + run.Error);
        var items = PipelineWork.ResultItems(cache[fill.Id]);
        Check(items.Count == 2 && items[1].Name == "February" && Convert.ToInt32(items[1].Table.Rows[0][2]) == 8,
            "The single A input was not reused with each B workbook.");

        Link(pipeline, february, fill, "A2");
        cache.Remove(fill.Id);
        run = PipelineWork.Run(pipeline, cache, fill.Id, false);
        Check(run.Error == null && Convert.ToInt32(PipelineWork.ResultItems(cache[fill.Id])[1].Table.Rows[0][2]) == 6,
            "The A and B batches were not paired by input order.");

        Link(pipeline, january, fill, "A3");
        cache.Remove(fill.Id);
        run = PipelineWork.Run(pipeline, cache, fill.Id, false);
        Check(run.FailedNodeId == fill.Id && run.Error.Contains("same size"), "Unequal A and B batches were accepted.");
        PipelineWork.RemoveNode(pipeline, fill.Id);
    }

    static void CheckWorkbook(string path, string amountColumn, int firstValue)
    {
        using (var workbook = new XLWorkbook(path))
        {
            var sheet = workbook.Worksheet(1);
            Check(sheet.Cell("A1").Style.Font.Bold, "A workbook lost its header style.");
            var amount = sheet.Cell(amountColumn + "2");
            Check(amount.Style.NumberFormat.Format == "0.00" && amount.GetValue<int>() == firstValue,
                "Sorting or formatting did not follow the Amount header.");
        }
    }
}
