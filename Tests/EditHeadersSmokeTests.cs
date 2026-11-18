using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.EditHeaders;
using OfficeWorkAssistant.Features.Pipeline;

internal static class EditHeadersSmokeTests
{
    static int checks;
    static void Check(bool ok, string message)
    {
        if (!ok) throw new Exception(message);
        checks++;
    }

    static void Fails(Action action, string text)
    {
        try { action(); }
        catch (InvalidOperationException ex)
        {
            Check(ex.Message.Contains(text), ex.Message);
            return;
        }
        throw new Exception("Expected error: " + text);
    }

    static void Main()
    {
        string folder = Path.Combine(Path.GetTempPath(), "HeaderTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            var table = new DataTable();
            ExcelFile.AddColumn(table, " qty ");
            ExcelFile.AddColumn(table, "").Caption = "";
            ExcelFile.AddColumn(table, "Price");
            table.Rows.Add(12, "hello", 3.5);
            var s = new EditHeadersSettings { Formula = "if(isblank($Value), \"Field_\" + $Column, upper($Value))" };
            s.Columns.Add(new HeaderEdit { Source = "A", Mode = HeaderEditMode.Text, Value = "Quantity" });
            s.Columns.Add(new HeaderEdit { Source = "C", Mode = HeaderEditMode.Clear });
            s.Columns.Add(new HeaderEdit { Mode = HeaderEditMode.Formula, Value = "\"New_\" + $Column + $Index" });
            var result = EditHeadersWork.Run(table, s);
            Check(ExcelFile.Header(result.Columns[0]) == "Quantity", "Literal override failed.");
            Check(ExcelFile.Header(result.Columns[1]) == "Field_B", "Missing header formula failed.");
            Check(ExcelFile.Header(result.Columns[2]) == "", "Header was not cleared.");
            Check(ExcelFile.Header(result.Columns[3]) == "New_D4", "New column formula failed.");
            Check(result.Rows.Count == 1 && (int)result.Rows[0][0] == 12 && result.Rows[0].IsNull(3), "Data changed.");
            Check(table.Columns.Count == 3 && ExcelFile.Header(table.Columns[2]) == "Price", "Input mutated.");
            var cross = new EditHeadersSettings();
            cross.Columns.Add(new HeaderEdit { Source = "A", Mode = HeaderEditMode.Text, Value = "Changed" });
            cross.Columns.Add(new HeaderEdit { Source = "B", Mode = HeaderEditMode.Formula, Value = "$A + \" copy\"" });
            Check(ExcelFile.Header(EditHeadersWork.Run(table, cross).Columns[1]) == "qty copy", "Formula did not use original headers.");
            Fails(() => EditHeadersWork.Run(table, new EditHeadersSettings { Formula = "upper(" }), "Header formula");
            Fails(() => EditHeadersWork.Run(table, new EditHeadersSettings { Formula = "$Typo" }), "Unknown header variable");
            cross.Columns.Add(new HeaderEdit { Source = "A", Mode = HeaderEditMode.Clear });
            Fails(() => EditHeadersWork.Run(table, cross), "only once");
            Check(EditHeadersWork.Run(table.Clone(), s).Rows.Count == 0, "Empty table failed.");

            // Blank and duplicate names must round-trip to the same target columns.
            table.Columns[2].Caption = "qty";
            var edits = new EditHeadersSettings();
            edits.Columns.Add(new HeaderEdit { Source = "B", Mode = HeaderEditMode.Text, Value = "Missing" });
            edits.Columns.Add(new HeaderEdit { Source = "C", Mode = HeaderEditMode.Text, Value = "Second qty" });
            var saved = PipelineWork.HeadersToNames(edits, table);
            var restored = PipelineWork.HeadersToLetters(PipelineWork.Clone(saved), table, true);
            Check(restored.Columns[0].Source == "B" && restored.Columns[1].Source == "C", "Ambiguous header reference.");
            Check(edits.Columns[0].Source == "B", "Settings mutated.");
            table.Columns.RemoveAt(1);
            Fails(() => PipelineWork.HeadersToLetters(saved, table, true), "missing");

            string source = Path.Combine(folder, "input.xlsx");
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1).Value = 12;
                sheet.Cell(1, 2).Value = new DateTime(2020, 1, 2);
                sheet.Cell(2, 1).Value = 34;
                sheet.Cell(2, 2).Value = true;
                workbook.SaveAs(source);
            }
            var loaded = ExcelFile.Load(source, null, 0, false);
            Check(loaded.Table.Rows.Count == 2 && loaded.TotalRows == 2, "Headerless load lost a row.");
            Check(Convert.ToDouble(loaded.Table.Rows[0][0]) == 12 && loaded.Table.Rows[0][1] is DateTime, "First row types lost.");
            Check(ExcelFile.Header(loaded.Table.Columns[0]) == "", "Headerless load has a name.");
            Check(ExcelFile.Load(source, null, 1, false).Truncated, "Preview limit failed.");
            Check(ExcelFile.Load(source, null).Table.Rows.Count == 1, "Default load behavior changed.");

            var def = new PipelineDefinition();
            var load = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 0);
            load.Settings = new LoadStepSettings { Path = source, FirstRowIsData = true };
            var edit = PipelineWork.AddNode(def, PipelineStepKind.EditHeaders, 100, 0);
            edit.Settings = new EditHeadersSettings { Formula = "\"Field_\" + $Column" };
            var save = PipelineWork.AddNode(def, PipelineStepKind.Save, 200, 0);
            string output = Path.Combine(folder, "output.xlsx");
            save.Settings = new SaveStepSettings { Path = output };
            def.Links.Add(new PipelineLink { From = load.Id, To = edit.Id, Port = "In" });
            def.Links.Add(new PipelineLink { From = edit.Id, To = save.Id, Port = "In" });
            def = PipelineWork.Clone(def);
            var run = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, false);
            Check(run.Error == null && !File.Exists(output), "Preview failed or wrote a file: " + run.Error);
            run = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, true);
            Check(run.Error == null && File.Exists(output), "Pipeline save failed: " + run.Error);
            var reread = ExcelFile.Load(output, null).Table;
            Check(reread.Rows.Count == 2 && ExcelFile.Header(reread.Columns[1]) == "Field_B", "Pipeline round-trip failed.");
            ExcelFile.Save(result, output);
            Check(ExcelFile.Header(ExcelFile.Load(output, null).Table.Columns[2]) == "", "Blank header did not survive save/load.");
            Console.WriteLine("Edit headers: " + checks + " checks passed.");
        }
        finally { Directory.Delete(folder, true); }
    }
}
