using System;
using System.Data;
using System.IO;
using System.Xml.Serialization;
using ClosedXML.Excel;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.FormatSheet;
using OfficeWorkAssistant.Features.Highlight;
using OfficeWorkAssistant.Features.Pipeline;

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
            Console.WriteLine("Format sheet smoke tests passed.");
        }
        finally { File.Delete(output); }
    }
}
