using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Features.FileOps;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.LoadImage;
using OfficeWorkAssistant.Features.Pipeline;
using OfficeWorkAssistant.Features.SaveText;

// Run mode: the choices a saved pipeline asks the person running it for, what it says it does,
// and running it with those choices. Everything stays in one temp folder; Documents is never touched.
internal static class RunnerSmokeTests
{
    static int _checks;

    static void Check(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
        _checks++;
    }

    // ---------- small helpers ----------

    static void Link(PipelineDefinition def, PipelineNode from, PipelineNode to)
    {
        def.Links.Add(new PipelineLink { From = from.Id, To = to.Id, Port = PipelineWork.PortIn });
    }

    static PipelineNode Of(PipelineDefinition def, PipelineStepKind kind)
    {
        foreach (var node in def.Nodes)
            if (node.Kind == kind)
                return node;
        throw new Exception("The test pipeline has no " + kind.ToString() + " step.");
    }

    static RunnerInput Choice(List<RunnerInput> inputs, string nodeId)
    {
        foreach (var input in inputs)
            if (input.NodeId == nodeId)
                return input;
        throw new Exception("No choice for step " + nodeId + ".");
    }

    static RunnerStepNote Note(List<RunnerStepNote> notes, PipelineStepKind kind)
    {
        foreach (var note in notes)
            if (note.Kind == kind)
                return note;
        return null;
    }

    static DataTable Table(string marker, int rows)
    {
        var table = new DataTable();
        table.Locale = CultureInfo.InvariantCulture;
        ExcelFile.AddColumn(table, "Name");
        ExcelFile.AddColumn(table, "Amount");
        for (var i = 1; i <= rows; i++)
            table.Rows.Add(marker + " " + i.ToString(CultureInfo.InvariantCulture), i * 10);
        return table;
    }

    static string First(DataTable table, string header)
    {
        return Convert.ToString(table.Rows[0][ExcelFile.FindColumn(table, header)], CultureInfo.InvariantCulture);
    }

    static bool Says(string text, string part)
    {
        return text != null && text.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    static bool Says(List<string> lines, string part)
    {
        foreach (var line in lines)
            if (Says(line, part))
                return true;
        return false;
    }

    static bool Same(byte[] left, byte[] right)
    {
        if (left.Length != right.Length)
            return false;
        for (var i = 0; i < left.Length; i++)
            if (left[i] != right[i])
                return false;
        return true;
    }

    // Load -> Filter & Sort -> Save, List folder -> File action, plus a Load image and a Save text:
    // seven steps and six choices.
    static PipelineDefinition FullDefinition(string source, string output, string textOutput, string scan, string copies, string image)
    {
        var def = new PipelineDefinition();
        var load = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 0);
        load.Settings = new LoadStepSettings { Path = source, Sheet = "" };
        var listed = PipelineWork.AddNode(def, PipelineStepKind.ListFolder, 0, 100);
        listed.Settings = new FolderScanSettings { Folder = scan };
        var picture = PipelineWork.AddNode(def, PipelineStepKind.LoadImage, 0, 200);
        picture.Settings = new LoadImageSettings { Path = image };
        var filter = PipelineWork.AddNode(def, PipelineStepKind.FilterSort, 100, 0);
        filter.Settings = new FilterSortSettings();
        var action = PipelineWork.AddNode(def, PipelineStepKind.FileAction, 100, 100);
        action.Settings = new FileActionSettings { Action = FileActionKind.Copy, PathColumn = "Path", DestFolder = copies };
        var save = PipelineWork.AddNode(def, PipelineStepKind.Save, 200, 0);
        save.Settings = new SaveStepSettings { Path = output, Mode = SaveStepMode.NewFile };
        var text = PipelineWork.AddNode(def, PipelineStepKind.SaveText, 200, 100);
        text.Settings = new SaveTextSettings { Path = textOutput, Format = SaveTextFormat.Text, Column = "Name" };
        Link(def, load, filter);
        Link(def, filter, save);
        Link(def, filter, text);
        Link(def, listed, action);
        return def;
    }

    // Load -> Filter & Sort (keeps everything) -> Save to a new file.
    static PipelineDefinition SimpleDefinition(string source, string output)
    {
        var def = new PipelineDefinition();
        var load = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 0);
        load.Settings = new LoadStepSettings { Path = source, Sheet = "" };
        var filter = PipelineWork.AddNode(def, PipelineStepKind.FilterSort, 100, 0);
        filter.Settings = new FilterSortSettings();
        var save = PipelineWork.AddNode(def, PipelineStepKind.Save, 200, 0);
        save.Settings = new SaveStepSettings { Path = output, Mode = SaveStepMode.NewFile };
        Link(def, load, filter);
        Link(def, filter, save);
        return def;
    }

    // ---------- 1. the choices of a full pipeline ----------

    static void InputChecks(string folder)
    {
        string source = Path.Combine(folder, "input-load-a.xlsx");
        ExcelFile.Save(Table("A", 2), source, "Data", false);
        string output = Path.Combine(folder, "input-out-x.xlsx");
        string textOutput = Path.Combine(folder, "input-result.txt");
        string scan = Path.Combine(folder, "input-scan");
        string copies = Path.Combine(folder, "input-copies");
        Directory.CreateDirectory(scan);
        Directory.CreateDirectory(copies);
        string image = Path.Combine(folder, "input-image.png");
        File.WriteAllBytes(image, new byte[] { 0, 1, 2 });
        var def = FullDefinition(source, output, textOutput, scan, copies, image);
        string file = Path.Combine(folder, "input-runner.xml");
        PipelineWork.SaveFile(def, file);
        byte[] saved = File.ReadAllBytes(file);

        var inputs = RunnerWork.Inputs(def);
        Check(inputs.Count == 6, "A full pipeline should ask six things, not " + inputs.Count + ".");
        var kinds = new List<string>();
        foreach (var input in inputs)
            kinds.Add(input.Kind.ToString());
        Check(string.Join(", ", kinds.ToArray()) == "ExcelFile, Folder, ImageFile, Folder, ExcelOutput, TextOutput",
            "The choices are not the expected kinds in running order: " + string.Join(", ", kinds.ToArray()));
        // Every choice follows the running order of the step it belongs to.
        var running = PipelineWork.Order(def);
        var lastAt = -1;
        foreach (var input in inputs)
        {
            var at = -1;
            for (var i = 0; i < running.Count; i++)
                if (running[i].Id == input.NodeId)
                    at = i;
            Check(at > lastAt, "A choice is out of running order: " + input.Kind.ToString() + " for " + input.StepTitle);
            lastAt = at;
        }

        var loadInput = Choice(inputs, Of(def, PipelineStepKind.Load).Id);
        Check(loadInput.Value == source && loadInput.HasSheet && !loadInput.IsOutput && loadInput.DefaultExtension == "" && !loadInput.IsJson,
            "The Load step did not ask for its Excel file and sheet.");
        var scanInput = Choice(inputs, Of(def, PipelineStepKind.ListFolder).Id);
        Check(scanInput.Kind == RunnerInputKind.Folder && scanInput.Value == scan && !scanInput.IsOutput,
            "The List folder step did not ask for its folder.");
        var imageInput = Choice(inputs, Of(def, PipelineStepKind.LoadImage).Id);
        Check(imageInput.Kind == RunnerInputKind.ImageFile && imageInput.Value == image && !imageInput.IsOutput && !imageInput.HasSheet,
            "The Load image step did not ask for its image.");
        var actionInput = Choice(inputs, Of(def, PipelineStepKind.FileAction).Id);
        Check(actionInput.Kind == RunnerInputKind.Folder && actionInput.Value == copies && !actionInput.IsOutput,
            "The File action step did not ask where the copies go.");
        var saveInput = Choice(inputs, Of(def, PipelineStepKind.Save).Id);
        Check(saveInput.Kind == RunnerInputKind.ExcelOutput && saveInput.IsOutput && !saveInput.IsJson &&
            saveInput.Value == output && saveInput.DefaultExtension == ".xlsx", "The Save step did not ask for an .xlsx output file.");
        var textInput = Choice(inputs, Of(def, PipelineStepKind.SaveText).Id);
        Check(textInput.Kind == RunnerInputKind.TextOutput && textInput.IsOutput && textInput.Value == textOutput &&
            textInput.DefaultExtension == ".txt" && !textInput.IsJson, "The Save text step did not ask for a plain .txt file.");

        string changedSource = Path.Combine(folder, "input-load-b.xlsx");
        string changedOutput = Path.Combine(folder, "input-out-y.xlsx");
        loadInput.Value = changedSource;
        saveInput.Value = changedOutput;
        Check(((LoadStepSettings)Of(def, PipelineStepKind.Load).Settings).Path == changedSource, "Writing a choice did not reach the Load step.");
        Check(((SaveStepSettings)Of(def, PipelineStepKind.Save).Settings).Path == changedOutput, "Writing a choice did not reach the Save step.");
        Check(loadInput.Value == changedSource && saveInput.Value == changedOutput, "A written choice did not read back.");
        Check(Same(saved, File.ReadAllBytes(file)), "Changing a choice wrote to the pipeline file on disk.");
    }

    // ---------- 2. the sheets of the chosen workbook ----------

    static void SheetChecks(string folder)
    {
        string workbook = Path.Combine(folder, "sheets-two.xlsx");
        ExcelFile.Save(Table("One", 1), workbook, "One", false);
        ExcelFile.Save(Table("Two", 2), workbook, "Two", true);
        var def = new PipelineDefinition();
        var load = PipelineWork.AddNode(def, PipelineStepKind.Load, 0, 0);
        ((LoadStepSettings)load.Settings).Path = workbook;
        var input = RunnerWork.Inputs(def)[0];
        Check(input.HasSheet && input.Sheet == "", "A Load step should start on the first sheet.");
        string[] names = input.SheetNames();
        Check(names != null && names.Length == 2 && names[0] == "One" && names[1] == "Two", "The sheet names of the chosen file were not read.");
        Check(input.SheetProblem == null, "A readable workbook was reported as a problem.");
        input.Sheet = "Two";
        Check(input.Sheet == "Two", "The chosen sheet did not round-trip.");
        Check(RunnerWork.InputProblem(input) == null, "A sheet that is in the file was refused: " + RunnerWork.InputProblem(input));
        input.Sheet = "Gone";
        Check(Says(RunnerWork.InputProblem(input), "no sheet named"), "A sheet the file does not have was accepted.");
        Check(Says(RunnerWork.InputProblem(input), "One"), "A missing sheet does not name the sheets the file has.");
        input.Sheet = "Two";
        input.Value = Path.Combine(folder, "sheets-missing.xlsx");
        Check(input.SheetNames() == null, "A file that does not exist returned sheet names.");
    }

    // ---------- 3. what is still missing ----------

    static void ProblemChecks(string folder)
    {
        string source = Path.Combine(folder, "problem-ready.xlsx");
        ExcelFile.Save(Table("R", 1), source, "Data", false);
        string outFolder = Path.Combine(folder, "problem-out");
        string scan = Path.Combine(folder, "problem-scan");
        Directory.CreateDirectory(outFolder);
        Directory.CreateDirectory(scan);
        var def = SimpleDefinition("", "");
        var listed = PipelineWork.AddNode(def, PipelineStepKind.ListFolder, 0, 100);
        listed.Settings = new FolderScanSettings { Folder = "" };
        var save = Of(def, PipelineStepKind.Save);
        var inputs = RunnerWork.Inputs(def);
        var loadInput = Choice(inputs, Of(def, PipelineStepKind.Load).Id);
        var scanInput = Choice(inputs, listed.Id);
        var saveInput = Choice(inputs, save.Id);

        Check(Says(RunnerWork.InputProblem(loadInput), loadInput.Label), "A blank choice does not name what to fill in.");
        loadInput.Value = Path.Combine(folder, "problem-gone.xlsx");
        Check(Says(RunnerWork.InputProblem(loadInput), "does not exist"), "A file that is not there was accepted.");
        scanInput.Value = Path.Combine(folder, "problem-gone-folder");
        Check(Says(RunnerWork.InputProblem(scanInput), "does not exist"), "A folder that is not there was accepted.");
        saveInput.Value = Path.Combine(outFolder, "problem-out.txt");
        Check(Says(RunnerWork.InputProblem(saveInput), ".xlsx"), "An Excel output with the wrong extension was accepted.");
        saveInput.Value = Path.Combine(folder, "problem-gone-folder", "problem-out.xlsx");
        Check(Says(RunnerWork.InputProblem(saveInput), "does not exist"), "An output folder that is not there was accepted.");

        loadInput.Value = "";
        saveInput.Value = Path.Combine(outFolder, "problem-out.txt");
        Check(Says(RunnerWork.CheckInputs(inputs), "Fill in"), "CheckInputs did not stop at the first problem.");
        loadInput.Value = source;
        scanInput.Value = scan;
        saveInput.Value = Path.Combine(outFolder, "problem-final.xlsx");
        Check(RunnerWork.CheckInputs(inputs) == null, "A ready pipeline was not accepted: " + RunnerWork.CheckInputs(inputs));
    }

    // ---------- 4. what the pipeline says it does ----------

    static void DescribeChecks(string folder)
    {
        string source = Path.Combine(folder, "describe-load.xlsx");
        ExcelFile.Save(Table("D", 1), source, "Data", false);
        string scan = Path.Combine(folder, "describe-scan");
        string copies = Path.Combine(folder, "describe-copies");
        Directory.CreateDirectory(scan);
        Directory.CreateDirectory(copies);
        var def = FullDefinition(source, Path.Combine(folder, "describe-out.xlsx"), Path.Combine(folder, "describe.txt"),
            scan, copies, Path.Combine(folder, "describe-image.png"));
        var order = PipelineWork.Order(def);
        var notes = RunnerWork.Describe(def);
        var lined = true;
        var described = true;
        for (var i = 0; i < notes.Count; i++)
        {
            if (notes[i].Kind != order[i].Kind || notes[i].Label != PipelineWork.KindLabel(order[i].Kind))
                lined = false;
            if (string.IsNullOrWhiteSpace(notes[i].Text))
                described = false;
        }
        Check(notes.Count == 7 && notes.Count == order.Count && lined, "Describe should give one line per step in running order.");
        Check(described, "A step has no plain-language description.");
        var saveNote = Note(notes, PipelineStepKind.Save);
        Check(saveNote != null && Says(saveNote.Text, "Excel file"), "The Save line does not mention the Excel file.");
        var actionNote = Note(notes, PipelineStepKind.FileAction);
        Check(actionNote != null && Says(actionNote.Text, "Copies"), "The File action line is not about copying.");
        string headline = RunnerWork.Headline(def);
        Check(Says(headline, "7 steps") && Says(headline, "4 inputs") && Says(headline, "2 outputs"),
            "The headline should count the steps, inputs and outputs: " + headline);

        var quiet = SimpleDefinition(source, Path.Combine(folder, "describe-out.xlsx"));
        PipelineWork.RemoveNode(quiet, Of(quiet, PipelineStepKind.Save).Id);
        Check(!RunnerWork.WritesAnything(quiet), "A pipeline that only reads was said to change files.");
        Check(RunnerWork.WritesAnything(def), "A pipeline with two saves was said to write nothing.");
    }

    // ---------- 5. running with the choices the user made ----------

    static void ChoiceRunChecks(string folder)
    {
        string fileA = Path.Combine(folder, "choice-a.xlsx");
        string fileB = Path.Combine(folder, "choice-b.xlsx");
        ExcelFile.Save(Table("A", 2), fileA, "Data", false);
        ExcelFile.Save(Table("B", 3), fileB, "Data", false);
        string pathX = Path.Combine(folder, "choice-x.xlsx");
        string pathY = Path.Combine(folder, "choice-y.xlsx");
        string pipeline = Path.Combine(folder, "choice.xml");
        PipelineWork.SaveFile(SimpleDefinition(fileA, pathX), pipeline);
        var def = PipelineWork.LoadFile(pipeline);
        var save = Of(def, PipelineStepKind.Save);
        var inputs = RunnerWork.Inputs(def);
        Choice(inputs, Of(def, PipelineStepKind.Load).Id).Value = fileB;
        Choice(inputs, save.Id).Value = pathY;

        var cache = new Dictionary<string, DataTable>();
        var preview = PipelineWork.Run(def, cache, null, false);
        Check(preview.Error == null, "The dry run failed: " + preview.Error);
        Check(!File.Exists(pathX) && !File.Exists(pathY), "The dry run wrote a file.");
        DataTable chosen;
        cache.TryGetValue(save.Id, out chosen);
        Check(chosen != null && chosen.Rows.Count == 3 && First(chosen, "Name") == "B 1", "The dry run did not use the rows of the chosen file.");

        byte[] before = File.ReadAllBytes(fileB);
        cache.Clear(); // the page reads every file again before the real run
        var run = PipelineWork.Run(def, cache, null, true);
        Check(run.Error == null, "The run failed: " + run.Error);
        Check(File.Exists(pathY) && !File.Exists(pathX), "The run did not write the chosen output file only.");
        Check(Same(before, File.ReadAllBytes(fileB)), "The run changed the file it read.");
        DataTable written = ExcelFile.Load(pathY, null).Table;
        Check(written.Rows.Count == 3 && First(written, "Name") == "B 1", "The written file does not hold the chosen file's rows.");

        string disk = File.ReadAllText(pipeline);
        Check(Says(disk, "choice-a.xlsx") && Says(disk, "choice-x.xlsx"), "The saved pipeline lost the file it was built with.");
        Check(!Says(disk, "choice-b.xlsx") && !Says(disk, "choice-y.xlsx"), "The runner wrote its own choices into the pipeline file.");
    }

    // ---------- 6. what a run would write and do ----------

    static void OutputListChecks(string folder)
    {
        string source = Path.Combine(folder, "list-source.xlsx");
        ExcelFile.Save(Table("L", 2), source, "Data", false);
        string target = Path.Combine(folder, "list-target.xlsx");
        var def = SimpleDefinition(source, target);
        var save = Of(def, PipelineStepKind.Save);
        Check(RunnerWork.ReplacedFiles(def).Count == 0, "A file that is not there was reported as replaced.");
        File.WriteAllText(target, "placeholder");
        var replaced = RunnerWork.ReplacedFiles(def);
        Check(replaced.Count == 1 && replaced[0] == target, "An existing output file was not reported as replaced.");
        var cache = new Dictionary<string, DataTable>();
        cache[save.Id] = Table("New", 3);
        var lines = RunnerWork.OutputLines(def, cache);
        Check(lines.Count == 1 && Says(lines, target), "One Save step should give one line naming the file.");
        Check(Says(lines, "3 rows") && Says(lines, "replacing"), "The output line should count the rows and say the file is replaced.");

        var sheetDef = SimpleDefinition(source, target);
        var sheetSave = (SaveStepSettings)Of(sheetDef, PipelineStepKind.Save).Settings;
        sheetSave.Mode = SaveStepMode.NewSheet;
        sheetSave.SheetName = "Extra";
        var sheetCache = new Dictionary<string, DataTable>();
        sheetCache[Of(sheetDef, PipelineStepKind.Save).Id] = Table("New", 3);
        Check(Says(RunnerWork.OutputLines(sheetDef, sheetCache), "new sheet named"), "A new-sheet save was not described.");

        string scan = Path.Combine(folder, "list-scan");
        string copyTo = Path.Combine(folder, "list-copies");
        Directory.CreateDirectory(scan);
        Directory.CreateDirectory(copyTo);
        File.WriteAllText(Path.Combine(scan, "one.txt"), "1");
        File.WriteAllText(Path.Combine(scan, "two.txt"), "2");
        var actionDef = new PipelineDefinition();
        var listed = PipelineWork.AddNode(actionDef, PipelineStepKind.ListFolder, 0, 0);
        listed.Settings = new FolderScanSettings { Folder = scan };
        var action = PipelineWork.AddNode(actionDef, PipelineStepKind.FileAction, 100, 0);
        action.Settings = new FileActionSettings { Action = FileActionKind.Copy, PathColumn = "Path", DestFolder = copyTo };
        Link(actionDef, listed, action);
        var actionCache = new Dictionary<string, DataTable>();
        Check(PipelineWork.Run(actionDef, actionCache, null, false).Error == null, "The file action plan failed.");
        var actionLines = RunnerWork.ActionLines(actionDef, actionCache);
        Check(actionLines.Count == 1 && Says(actionLines, "Copy") && Says(actionLines, copyTo) && Says(actionLines, "2 items"),
            "The action line should copy two items to the destination folder.");
    }

    // ---------- 7. why a run stopped ----------

    static void FailureChecks(string folder)
    {
        string source = Path.Combine(folder, "failure-load.xlsx");
        ExcelFile.Save(Table("F", 2), source, "Data", false);
        var def = SimpleDefinition(Path.Combine(folder, "failure-gone.xlsx"), Path.Combine(folder, "failure-out.xlsx"));
        var result = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, false);
        Check(result.Error != null && result.FailedNodeId == Of(def, PipelineStepKind.Load).Id,
            "A file that is not there did not stop the run at the Load step.");
        string text = RunnerWork.FailureText(def, result);
        Check(Says(text, "Step 1") && Says(text, "Load file"), "The failure does not name the Load step.");
        Check(Says(text, "Nothing was written"), "The failure does not say that nothing was written.");

        // ClosedXML creates a missing output folder when it saves, so a Save step only fails when the
        // folder cannot exist at all: here the folder's place is taken by a real file.
        var blocked = SimpleDefinition(source, Path.Combine(source, "failure-out.xlsx"));
        var saveResult = PipelineWork.Run(blocked, new Dictionary<string, DataTable>(), null, true);
        Check(saveResult.Error != null && saveResult.FailedNodeId == Of(blocked, PipelineStepKind.Save).Id,
            "A Save step that cannot write its file did not stop the run.");
        string saveText = RunnerWork.FailureText(blocked, saveResult);
        Check(Says(saveText, "Step 3") && Says(saveText, "Save file"), "The failure does not name the Save step.");
        Check(Says(saveText, "Nothing was written"), "The failed save does not say that nothing was written.");
    }

    // ---------- 8. what happened, and what to open ----------

    static void FolderChecks(string folder)
    {
        string source = Path.Combine(folder, "folder-load.xlsx");
        ExcelFile.Save(Table("W", 2), source, "Data", false);
        string outFolder = Path.Combine(folder, "folder-out");
        string scan = Path.Combine(folder, "folder-scan");
        string dest = Path.Combine(folder, "folder-dest");
        Directory.CreateDirectory(outFolder);
        Directory.CreateDirectory(scan);
        Directory.CreateDirectory(dest);
        File.WriteAllText(Path.Combine(scan, "paper.txt"), "one");
        var def = SimpleDefinition(source, Path.Combine(outFolder, "folder-out.xlsx"));
        var listed = PipelineWork.AddNode(def, PipelineStepKind.ListFolder, 0, 100);
        listed.Settings = new FolderScanSettings { Folder = scan };
        var action = PipelineWork.AddNode(def, PipelineStepKind.FileAction, 100, 100);
        action.Settings = new FileActionSettings { Action = FileActionKind.Copy, PathColumn = "Path", DestFolder = dest };
        Link(def, listed, action);

        var result = PipelineWork.Run(def, new Dictionary<string, DataTable>(), null, true);
        Check(result.Error == null, "The full run failed: " + result.Error);
        Check(result.Written.Count == 1 && result.Actions.Count == 1, "The run did not report one saved file and one file action.");
        Check(File.Exists(Path.Combine(outFolder, "folder-out.xlsx")) && File.Exists(Path.Combine(dest, "paper.txt")),
            "The run did not write the file and copy the item.");
        var folders = RunnerWork.Folders(def, result);
        Check(folders.Count == 2 && Says(folders, outFolder) && Says(folders, dest),
            "Opening the results should offer the two folders once each: " + folders.Count + ".");
        Check(Says(RunnerWork.DoneLines(result), "Copied") && Says(RunnerWork.DoneHeadline(result), "Finished"),
            "The done lines or the done headline are missing.");
    }

    // ---------- 9. a runner-made choice lives only until someone saves the pipeline ----------

    static void RoundTripChecks(string folder)
    {
        string pipeline = Path.Combine(folder, "round-trip.xml");
        PipelineWork.SaveFile(SimpleDefinition(Path.Combine(folder, "round-built.xlsx"), Path.Combine(folder, "round-out.xlsx")), pipeline);
        var def = PipelineWork.LoadFile(pipeline);
        string chosen = Path.Combine(folder, "round-chosen.xlsx");
        string chosenOut = Path.Combine(folder, "round-chosen-out.xlsx");
        Choice(RunnerWork.Inputs(def), Of(def, PipelineStepKind.Load).Id).Value = chosen;
        Choice(RunnerWork.Inputs(def), Of(def, PipelineStepKind.Save).Id).Value = chosenOut;
        PipelineWork.SaveFile(def, pipeline); // what builder mode does, not the runner
        var again = PipelineWork.LoadFile(pipeline);
        Check(((LoadStepSettings)Of(again, PipelineStepKind.Load).Settings).Path == chosen, "A saved Load choice was lost.");
        Check(((SaveStepSettings)Of(again, PipelineStepKind.Save).Settings).Path == chosenOut, "A saved output choice was lost.");
        Check(Choice(RunnerWork.Inputs(again), Of(again, PipelineStepKind.Load).Id).Value == chosen,
            "The runner does not read the saved Load choice back.");
    }

    static int Main()
    {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeWorkAssistant-Runner-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        bool passed = false;
        try
        {
            InputChecks(folder);
            SheetChecks(folder);
            ProblemChecks(folder);
            DescribeChecks(folder);
            ChoiceRunChecks(folder);
            OutputListChecks(folder);
            FailureChecks(folder);
            FolderChecks(folder);
            RoundTripChecks(folder);
            Console.WriteLine("Passed " + _checks + " run mode checks.");
            passed = true;
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Console.Error.WriteLine("Test files kept in: " + folder);
            return 1;
        }
        finally
        {
            if (passed)
                Directory.Delete(folder, true);
        }
    }
}
