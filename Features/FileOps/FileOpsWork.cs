using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using Microsoft.VisualBasic.FileIO;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FileOps
{
    public enum ValueListMode
    {
        List,
        Set,
        Map
    }

    public enum ValueListSource
    {
        Input,
        Stored
    }

    public enum MapDuplicate
    {
        First,
        Last,
        Error
    }

    public enum FileInclude
    {
        Files,
        Folders,
        Both
    }

    public enum FileMatchMode
    {
        NameEquals,
        StemEquals,
        Contains,
        StartsWith,
        EndsWith,
        Formula
    }

    public enum FileActionKind
    {
        Copy,
        Move,
        Delete
    }

    public enum FileConflict
    {
        Skip,
        Overwrite,
        Rename
    }

    public sealed class ValueListItem
    {
        [XmlAttribute]
        public string Key { get; set; }
        [XmlAttribute]
        public string Value { get; set; }
    }

    // Keys (and values, for a map) built from the input rows, or stored in the pipeline file.
    public sealed class ValueListSettings
    {
        public ValueListMode Mode { get; set; }
        public ValueListSource Source { get; set; }
        public string KeyFormula { get; set; }
        public string ValueFormula { get; set; }
        public bool Trim { get; set; }
        public bool SkipBlanks { get; set; }
        public MapDuplicate OnDuplicate { get; set; }
        public List<ValueListItem> Items { get; set; }

        public ValueListSettings()
        {
            Mode = ValueListMode.Set;
            KeyFormula = "";
            ValueFormula = "";
            Trim = true;
            SkipBlanks = true;
            Items = new List<ValueListItem>();
        }
    }

    // Where to look and what to list. Used alone (List folder) and inside Find files.
    public sealed class FolderScanSettings
    {
        public string Folder { get; set; }
        public bool Recursive { get; set; }
        public FileInclude Include { get; set; }
        // Optional true/false formulas over the file variables: one for files, one for folders.
        public string FileFilter { get; set; }
        public string FolderFilter { get; set; }
        // One character that splits the stem into $1 $2 ...
        public string Separator { get; set; }

        public FolderScanSettings()
        {
            Folder = "";
            Recursive = true;
            FileFilter = "";
            FolderFilter = "";
            Separator = "_";
        }
    }

    public sealed class FindFilesSettings
    {
        public FolderScanSettings Scan { get; set; }
        public FileMatchMode Match { get; set; }
        // Formula mode: true when the item belongs to the key ($Key, $Value and the file variables).
        // Files use FileFormula, folders use FolderFormula.
        public string FileFormula { get; set; }
        public string FolderFormula { get; set; }
        public bool KeepUnmatched { get; set; }
        public bool SkipInsideMatchedFolders { get; set; }

        public FindFilesSettings()
        {
            Scan = new FolderScanSettings();
            Match = FileMatchMode.StemEquals;
            FileFormula = "";
            FolderFormula = "";
            SkipInsideMatchedFolders = true;
        }
    }

    public sealed class FileActionSettings
    {
        public FileActionKind Action { get; set; }
        // Header of the input column that holds the paths.
        public string PathColumn { get; set; }
        public string DestFolder { get; set; }
        // Optional formulas, one pair for files and one for folders. Subfolder may give nested
        // folders with \. A blank name keeps the item's name.
        public string FileSubfolderFormula { get; set; }
        public string FileNameFormula { get; set; }
        public string FolderSubfolderFormula { get; set; }
        public string FolderNameFormula { get; set; }
        // Recreate the input's Relative folder under the destination.
        public bool KeepStructure { get; set; }
        public FileConflict Conflict { get; set; }
        // Delete only: skip the Recycle Bin.
        public bool Permanent { get; set; }
        public string Separator { get; set; }

        public FileActionSettings()
        {
            PathColumn = "Path";
            DestFolder = "";
            FileSubfolderFormula = "";
            FileNameFormula = "";
            FolderSubfolderFormula = "";
            FolderNameFormula = "";
            Separator = "_";
        }
    }

    // Pure logic for the pipeline's file steps. No WPF. Tables in, tables out; nothing on disk
    // changes except in Apply.
    public static class FileOpsWork
    {
        public const string StatusReady = "Ready";
        public const string StatusOverwrite = "Overwrite";

        static readonly char[] Illegal = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

        sealed class FileItem
        {
            public string Path;
            public bool IsFolder;
        }

        static FileOpsWork()
        {
            var item = new FileItem { Path = @"C:\src\docs\A100_report.pdf" };
            var lookup = FileLookup(item, @"C:\src", '_', null);
            if (!Matches("$1 == \"A100\" && $Ext == \".pdf\" && $Relative == \"docs\" && !$IsFolder", lookup))
                throw new InvalidOperationException("FileOpsWork file variables check failed.");
            if (NextFree(@"C:\x\a.pdf", false, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { @"C:\x\a (2).pdf" }) != @"C:\x\a (3).pdf")
                throw new InvalidOperationException("FileOpsWork rename check failed.");
            if (SafeSubPath(@"..\2024\ a:b ") != @"2024\ab")
                throw new InvalidOperationException("FileOpsWork subfolder check failed.");
        }

        // ---------- List / set / map ----------

        public static string Check(ValueListSettings s)
        {
            if (s.Source == ValueListSource.Input)
            {
                var error = CheckFormula(s.KeyFormula, true, "Key");
                if (error != null)
                    return error;
                if (s.Mode == ValueListMode.Map)
                    return CheckFormula(s.ValueFormula, true, "Value");
            }
            return null;
        }

        // A table with Key (and Value for a map).
        public static DataTable BuildList(DataTable input, ValueListSettings s)
        {
            var error = Check(s);
            if (error != null)
                throw new InvalidOperationException(error);

            var pairs = new List<ValueListItem>();
            if (s.Source == ValueListSource.Stored)
            {
                foreach (var item in s.Items)
                {
                    if (item != null)
                        pairs.Add(new ValueListItem { Key = item.Key ?? "", Value = item.Value ?? "" });
                }
            }
            else
            {
                if (input == null)
                    throw new InvalidOperationException("Connect an input step, or switch to stored values.");
                var key = ExpressionEngine.Compile(s.KeyFormula);
                var value = s.Mode == ValueListMode.Map ? ExpressionEngine.Compile(s.ValueFormula) : null;
                for (var r = 0; r < input.Rows.Count; r++)
                {
                    var lookup = RowLookup(input, input.Rows[r]);
                    try
                    {
                        pairs.Add(new ValueListItem
                        {
                            Key = ExpressionEngine.ToText(key(lookup)),
                            Value = value == null ? "" : ExpressionEngine.ToText(value(lookup))
                        });
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Row " + (r + 2).ToString(CultureInfo.InvariantCulture) + ": " + ex.Message, ex);
                    }
                }
            }

            var table = new DataTable();
            var keyCol = ExcelFile.AddColumn(table, "Key");
            var valueCol = s.Mode == ValueListMode.Map ? ExcelFile.AddColumn(table, "Value") : null;
            var rows = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in pairs)
            {
                var k = s.Trim ? pair.Key.Trim() : pair.Key;
                var v = s.Trim ? pair.Value.Trim() : pair.Value;
                if (s.SkipBlanks && k.Trim().Length == 0)
                    continue;
                if (s.Mode != ValueListMode.List)
                {
                    DataRow existing;
                    if (rows.TryGetValue(k, out existing))
                    {
                        if (valueCol == null || s.OnDuplicate == MapDuplicate.First)
                            continue;
                        if (s.OnDuplicate == MapDuplicate.Error && !string.Equals(ExpressionEngine.ToText(existing[valueCol]), v, StringComparison.Ordinal))
                            throw new InvalidOperationException("Key '" + k + "' has two values: '" + ExpressionEngine.ToText(existing[valueCol]) + "' and '" + v + "'.");
                        existing[valueCol] = v;
                        continue;
                    }
                }
                var row = table.NewRow();
                row[keyCol] = k;
                if (valueCol != null)
                    row[valueCol] = v;
                table.Rows.Add(row);
                if (s.Mode != ValueListMode.List)
                    rows[k] = row;
            }
            return table;
        }

        // The list the input gives today, as values to store in the step.
        public static List<ValueListItem> Capture(DataTable input, ValueListSettings s)
        {
            var live = CloneSettings(s);
            live.Source = ValueListSource.Input;
            var table = BuildList(input, live);
            var items = new List<ValueListItem>(table.Rows.Count);
            foreach (DataRow row in table.Rows)
            {
                items.Add(new ValueListItem
                {
                    Key = ExpressionEngine.ToText(row[0]),
                    Value = table.Columns.Count > 1 ? ExpressionEngine.ToText(row[1]) : ""
                });
            }
            return items;
        }

        static ValueListSettings CloneSettings(ValueListSettings s)
        {
            return new ValueListSettings
            {
                Mode = s.Mode,
                Source = s.Source,
                KeyFormula = s.KeyFormula,
                ValueFormula = s.ValueFormula,
                Trim = s.Trim,
                SkipBlanks = s.SkipBlanks,
                OnDuplicate = s.OnDuplicate,
                Items = new List<ValueListItem>(s.Items)
            };
        }

        // ---------- List folder / Find files ----------

        public static string Check(FolderScanSettings s)
        {
            if (s == null || string.IsNullOrWhiteSpace(s.Folder))
                return "Choose a folder to search.";
            return CheckFormula(s.FileFilter, false, "File filter") ?? CheckFormula(s.FolderFilter, false, "Folder filter");
        }

        public static string Check(FindFilesSettings s)
        {
            var error = Check(s.Scan);
            if (error != null)
                return error;
            if (s.Match != FileMatchMode.Formula)
                return null;
            // Required only for the kinds of item the step looks for.
            return CheckFormula(s.FileFormula, s.Scan.Include != FileInclude.Folders, "File match") ??
                CheckFormula(s.FolderFormula, s.Scan.Include != FileInclude.Files, "Folder match");
        }

        public static DataTable ListFolder(FolderScanSettings s)
        {
            var error = Check(s);
            if (error != null)
                throw new InvalidOperationException(error);
            var table = NewFileTable(false, false, false);
            foreach (var item in Scan(s))
                AddFileRow(table, item, s.Folder, null, null, true);
            return table;
        }

        // One row per (key, matching item). keys: a Key column, or else its first column.
        public static DataTable FindFiles(DataTable keys, FindFilesSettings s)
        {
            var error = Check(s);
            if (error != null)
                throw new InvalidOperationException(error);
            if (keys == null || keys.Columns.Count == 0)
                throw new InvalidOperationException("The input has no columns to take keys from.");

            var keyCol = ExcelFile.FindColumn(keys, "Key") ?? keys.Columns[0];
            var valueCol = ExcelFile.FindColumn(keys, "Value");
            if (valueCol == keyCol)
                valueCol = null;

            var keyList = new List<KeyValuePair<string, string>>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in keys.Rows)
            {
                var k = ExpressionEngine.ToText(row[keyCol]).Trim();
                if (k.Length == 0 || !seen.Add(k))
                    continue;
                keyList.Add(new KeyValuePair<string, string>(k, valueCol == null ? "" : ExpressionEngine.ToText(row[valueCol])));
            }

            var items = Scan(s.Scan);
            var sep = SeparatorOf(s.Scan.Separator);
            var matches = new List<KeyValuePair<int, FileItem>>();

            if (s.Match == FileMatchMode.NameEquals || s.Match == FileMatchMode.StemEquals)
            {
                var byName = new Dictionary<string, List<FileItem>>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in items)
                {
                    var name = MatchName(item, s.Match);
                    List<FileItem> list;
                    if (!byName.TryGetValue(name, out list))
                        byName[name] = list = new List<FileItem>();
                    list.Add(item);
                }
                for (var k = 0; k < keyList.Count; k++)
                {
                    List<FileItem> list;
                    if (byName.TryGetValue(keyList[k].Key, out list))
                    {
                        foreach (var item in list)
                            matches.Add(new KeyValuePair<int, FileItem>(k, item));
                    }
                }
            }
            else
            {
                var formula = s.Match == FileMatchMode.Formula;
                var fileFormula = formula ? CompileOptional(s.FileFormula) : null;
                var folderFormula = formula ? CompileOptional(s.FolderFormula) : null;
                for (var k = 0; k < keyList.Count; k++)
                {
                    var key = keyList[k];
                    foreach (var item in items)
                    {
                        bool hit;
                        if (formula)
                        {
                            // A blank formula never matches its kind of item.
                            var f = item.IsFolder ? folderFormula : fileFormula;
                            hit = f != null && Matches(f, FileLookup(item, s.Scan.Folder, sep, KeyLookup(key.Key, key.Value)));
                        }
                        else
                        {
                            var name = MatchName(item, s.Match);
                            if (s.Match == FileMatchMode.Contains)
                                hit = name.IndexOf(key.Key, StringComparison.OrdinalIgnoreCase) >= 0;
                            else if (s.Match == FileMatchMode.StartsWith)
                                hit = name.StartsWith(key.Key, StringComparison.OrdinalIgnoreCase);
                            else
                                hit = name.EndsWith(key.Key, StringComparison.OrdinalIgnoreCase);
                        }
                        if (hit)
                            matches.Add(new KeyValuePair<int, FileItem>(k, item));
                    }
                }
            }

            var found = new bool[keyList.Count];
            foreach (var m in matches)
                found[m.Key] = true;

            // A matched folder is acted on as a whole, so drop matches inside it.
            if (s.SkipInsideMatchedFolders)
            {
                var folders = new List<string>();
                foreach (var m in matches)
                {
                    if (m.Value.IsFolder)
                        folders.Add(m.Value.Path.TrimEnd('\\') + "\\");
                }
                if (folders.Count > 0)
                {
                    matches.RemoveAll(m =>
                    {
                        foreach (var f in folders)
                        {
                            if (m.Value.Path.StartsWith(f, StringComparison.OrdinalIgnoreCase))
                                return true;
                        }
                        return false;
                    });
                }
            }

            var table = NewFileTable(true, valueCol != null, s.KeepUnmatched);
            var byKey = new List<FileItem>[keyList.Count];
            foreach (var m in matches)
            {
                if (byKey[m.Key] == null)
                    byKey[m.Key] = new List<FileItem>();
                byKey[m.Key].Add(m.Value);
            }
            for (var k = 0; k < keyList.Count; k++)
            {
                if (byKey[k] != null)
                {
                    foreach (var item in byKey[k])
                        AddFileRow(table, item, s.Scan.Folder, keyList[k].Key, keyList[k].Value, true);
                }
                else if (s.KeepUnmatched && !found[k])
                    AddFileRow(table, null, s.Scan.Folder, keyList[k].Key, keyList[k].Value, false);
            }
            return table;
        }

        static string MatchName(FileItem item, FileMatchMode mode)
        {
            var name = Path.GetFileName(item.Path);
            if (mode == FileMatchMode.StemEquals && !item.IsFolder)
                return Path.GetFileNameWithoutExtension(name);
            return name;
        }

        static DataTable NewFileTable(bool withKey, bool withValue, bool withFound)
        {
            var table = new DataTable();
            if (withKey)
                ExcelFile.AddColumn(table, "Key");
            if (withValue)
                ExcelFile.AddColumn(table, "Value");
            if (withFound)
                ExcelFile.AddColumn(table, "Found");
            foreach (var h in new[] { "Type", "Name", "Path", "Folder", "Relative", "Size", "Modified" })
                ExcelFile.AddColumn(table, h);
            return table;
        }

        static void AddFileRow(DataTable table, FileItem item, string root, string key, string value, bool found)
        {
            var row = table.NewRow();
            var keyCol = ExcelFile.FindColumn(table, "Key");
            var valueCol = ExcelFile.FindColumn(table, "Value");
            var foundCol = ExcelFile.FindColumn(table, "Found");
            if (keyCol != null)
                row[keyCol] = key ?? "";
            if (valueCol != null)
                row[valueCol] = value ?? "";
            if (foundCol != null)
                row[foundCol] = found;
            var c = ExcelFile.FindColumn(table, "Type").Ordinal;
            if (item != null)
            {
                row[c] = item.IsFolder ? "Folder" : "File";
                row[c + 1] = Path.GetFileName(item.Path);
                row[c + 2] = item.Path;
                row[c + 3] = Path.GetDirectoryName(item.Path) ?? "";
                row[c + 4] = RelativeFolder(item.Path, root);
                if (!item.IsFolder)
                {
                    try { row[c + 5] = (double)new FileInfo(item.Path).Length; }
                    catch { }
                }
                try { row[c + 6] = item.IsFolder ? Directory.GetLastWriteTime(item.Path) : File.GetLastWriteTime(item.Path); }
                catch { }
            }
            table.Rows.Add(row);
        }

        // Every file and/or folder under the scan folder that passes its filter.
        static List<FileItem> Scan(FolderScanSettings s)
        {
            var root = s.Folder.Trim();
            if (!Directory.Exists(root))
                throw new InvalidOperationException("Folder not found: " + root);
            var all = new List<FileItem>();
            Walk(root, s.Recursive, s.Include, all);
            var fileFilter = CompileOptional(s.FileFilter);
            var folderFilter = CompileOptional(s.FolderFilter);
            if (fileFilter == null && folderFilter == null)
                return all;
            var sep = SeparatorOf(s.Separator);
            return all.FindAll(i =>
            {
                var filter = i.IsFolder ? folderFilter : fileFilter;
                return filter == null || Matches(filter, FileLookup(i, root, sep, null));
            });
        }

        static Func<Func<string, object>, object> CompileOptional(string formula)
        {
            return string.IsNullOrWhiteSpace(formula) ? null : ExpressionEngine.Compile(formula);
        }

        // Manual walk: a folder we may not open (or a junction such as "Application Data" on
        // Windows 7) is skipped instead of failing the whole search.
        static void Walk(string dir, bool recursive, FileInclude include, List<FileItem> into)
        {
            if (include != FileInclude.Folders)
            {
                string[] files;
                try { files = Directory.GetFiles(dir); }
                catch (UnauthorizedAccessException) { files = new string[0]; }
                catch (IOException) { files = new string[0]; }
                foreach (var f in files)
                    into.Add(new FileItem { Path = f });
            }
            string[] subs;
            try { subs = Directory.GetDirectories(dir); }
            catch (UnauthorizedAccessException) { return; }
            catch (IOException) { return; }
            foreach (var sub in subs)
            {
                if (IsLink(sub))
                    continue;
                if (include != FileInclude.Files)
                    into.Add(new FileItem { Path = sub, IsFolder = true });
                if (recursive)
                    Walk(sub, true, include, into);
            }
        }

        static bool IsLink(string dir)
        {
            try { return (File.GetAttributes(dir) & FileAttributes.ReparsePoint) != 0; }
            catch { return true; }
        }

        // ---------- File action ----------

        public static string Check(FileActionSettings s)
        {
            if (string.IsNullOrWhiteSpace(s.PathColumn))
                return "Choose the column that holds the paths.";
            if (s.Action != FileActionKind.Delete)
            {
                if (string.IsNullOrWhiteSpace(s.DestFolder))
                    return "Choose the destination folder.";
                return CheckFormula(s.FileSubfolderFormula, false, "File subfolder") ??
                    CheckFormula(s.FileNameFormula, false, "File new name") ??
                    CheckFormula(s.FolderSubfolderFormula, false, "Folder subfolder") ??
                    CheckFormula(s.FolderNameFormula, false, "Folder new name");
            }
            return null;
        }

        // What would happen, row by row. Touches nothing. The input rows plus Destination and Status.
        public static DataTable PlanActions(DataTable input, FileActionSettings s)
        {
            var error = Check(s);
            if (error != null)
                throw new InvalidOperationException(error);
            var pathCol = ExcelFile.FindColumn(input, s.PathColumn);
            if (pathCol == null)
                throw new InvalidOperationException("Column '" + s.PathColumn + "' is not in this step's input.");

            var table = input.Copy();
            var destCol = ExcelFile.FindColumn(table, "Destination") ?? ExcelFile.AddColumn(table, "Destination");
            var statusCol = ExcelFile.FindColumn(table, "Status") ?? ExcelFile.AddColumn(table, "Status");
            var relCol = ExcelFile.FindColumn(table, "Relative");
            var sep = SeparatorOf(s.Separator);
            var fileSub = CompileOptional(s.FileSubfolderFormula);
            var fileName = CompileOptional(s.FileNameFormula);
            var folderSub = CompileOptional(s.FolderSubfolderFormula);
            var folderName = CompileOptional(s.FolderNameFormula);
            string dest = null;
            if (s.Action != FileActionKind.Delete)
                dest = Path.GetFullPath(s.DestFolder.Trim());

            var sources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in table.Rows)
            {
                row[destCol] = "";
                var src = ExpressionEngine.ToText(row[pathCol.Ordinal]).Trim();
                if (src.Length == 0)
                {
                    row[statusCol] = "No path";
                    continue;
                }
                try
                {
                    src = Path.GetFullPath(src);
                }
                catch (Exception ex)
                {
                    row[statusCol] = "Bad path: " + ex.Message;
                    continue;
                }
                bool isFolder = Directory.Exists(src);
                if (!isFolder && !File.Exists(src))
                {
                    row[statusCol] = "Missing";
                    continue;
                }
                if (!sources.Add(src))
                {
                    row[statusCol] = "Duplicate";
                    continue;
                }
                if (s.Action == FileActionKind.Delete)
                {
                    row[statusCol] = StatusReady;
                    continue;
                }

                var item = new FileItem { Path = src, IsFolder = isFolder };
                var subFormula = isFolder ? folderSub : fileSub;
                var nameFormula = isFolder ? folderName : fileName;
                var lookup = FileLookup(item, null, sep, RowLookup(table, row));
                var name = Path.GetFileName(src);
                if (nameFormula != null)
                {
                    var stem = Sanitize(Text(nameFormula, lookup));
                    name = stem.Length == 0 ? "" : stem + (isFolder ? "" : Path.GetExtension(src));
                }
                var sub = subFormula == null ? "" : SafeSubPath(Text(subFormula, lookup));
                if (name.Length == 0 || (subFormula != null && sub.Length == 0))
                {
                    row[statusCol] = "Bad name";
                    continue;
                }
                var rel = s.KeepStructure && relCol != null ? SafeSubPath(ExpressionEngine.ToText(row[relCol])) : "";
                var target = Path.Combine(Path.Combine(Path.Combine(dest, rel), sub), name);
                row[destCol] = target;

                if (string.Equals(target, src, StringComparison.OrdinalIgnoreCase))
                {
                    row[statusCol] = "Same path";
                    continue;
                }
                if (isFolder && target.StartsWith(src.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase))
                {
                    row[statusCol] = "Inside itself";
                    continue;
                }

                var taken = targets.Contains(target);
                var onDisk = File.Exists(target) || Directory.Exists(target);
                if (!taken && !onDisk)
                {
                    targets.Add(target);
                    row[statusCol] = StatusReady;
                }
                else if (s.Conflict == FileConflict.Rename)
                {
                    target = NextFree(target, isFolder, targets);
                    targets.Add(target);
                    row[destCol] = target;
                    row[statusCol] = StatusReady;
                }
                else if (taken)
                    row[statusCol] = "Duplicate target";
                else if (s.Conflict == FileConflict.Overwrite && isFolder == Directory.Exists(target))
                {
                    targets.Add(target);
                    row[statusCol] = StatusOverwrite;
                }
                else
                    row[statusCol] = "Exists - skip";
            }
            return table;
        }

        // Plans against the disk as it is now, then does it. Each row's Status says what happened.
        public static DataTable Apply(DataTable input, FileActionSettings s)
        {
            var table = PlanActions(input, s);
            var destCol = ExcelFile.FindColumn(table, "Destination");
            var statusCol = ExcelFile.FindColumn(table, "Status");
            var pathCol = ExcelFile.FindColumn(table, s.PathColumn);
            foreach (DataRow row in table.Rows)
            {
                var status = ExpressionEngine.ToText(row[statusCol]);
                var overwrite = status == StatusOverwrite;
                if (status != StatusReady && !overwrite)
                    continue;
                var src = Path.GetFullPath(ExpressionEngine.ToText(row[pathCol]).Trim());
                var target = ExpressionEngine.ToText(row[destCol]);
                try
                {
                    bool isFolder = Directory.Exists(src);
                    if (!isFolder && !File.Exists(src))
                    {
                        row[statusCol] = "Missing";
                        continue;
                    }
                    switch (s.Action)
                    {
                        case FileActionKind.Copy:
                            Copy(src, target, isFolder, overwrite);
                            row[statusCol] = overwrite ? "Copied (replaced)" : "Copied";
                            break;
                        case FileActionKind.Move:
                            Move(src, target, isFolder, overwrite);
                            row[statusCol] = overwrite ? "Moved (replaced)" : "Moved";
                            break;
                        default:
                            Delete(src, isFolder, s.Permanent);
                            row[statusCol] = s.Permanent ? "Deleted" : "Recycled";
                            break;
                    }
                }
                catch (Exception ex)
                {
                    row[statusCol] = "Failed: " + ex.Message;
                }
            }
            return table;
        }

        // Rows a plan would act on.
        public static int CountReady(DataTable plan)
        {
            var statusCol = plan == null ? null : ExcelFile.FindColumn(plan, "Status");
            if (statusCol == null)
                return 0;
            var n = 0;
            foreach (DataRow row in plan.Rows)
            {
                var status = ExpressionEngine.ToText(row[statusCol]);
                if (status == StatusReady || status == StatusOverwrite)
                    n++;
            }
            return n;
        }

        // "Copied 12, Exists - skip 3, Failed 1": counts by status, in first-seen order.
        public static string Outcome(DataTable table)
        {
            var statusCol = ExcelFile.FindColumn(table, "Status");
            if (statusCol == null || table.Rows.Count == 0)
                return "nothing to do";
            var counts = new Dictionary<string, int>();
            var order = new List<string>();
            foreach (DataRow row in table.Rows)
            {
                var status = ExpressionEngine.ToText(row[statusCol]);
                if (status.StartsWith("Failed", StringComparison.Ordinal))
                    status = "Failed";
                if (!counts.ContainsKey(status))
                {
                    counts[status] = 0;
                    order.Add(status);
                }
                counts[status]++;
            }
            var parts = new List<string>();
            foreach (var status in order)
                parts.Add(status + " " + counts[status].ToString(CultureInfo.InvariantCulture));
            return string.Join(", ", parts.ToArray());
        }

        public static string Describe(FileActionSettings s, int count)
        {
            var n = count.ToString(CultureInfo.InvariantCulture) + " item" + (count == 1 ? "" : "s");
            switch (s.Action)
            {
                case FileActionKind.Copy: return "Copy " + n + " to " + s.DestFolder;
                case FileActionKind.Move: return "Move " + n + " to " + s.DestFolder;
                default: return "Delete " + n + (s.Permanent ? " PERMANENTLY" : " (to the Recycle Bin)");
            }
        }

        static void Copy(string src, string target, bool isFolder, bool overwrite)
        {
            if (isFolder)
            {
                CopyFolder(src, target, overwrite);
                return;
            }
            var dir = Path.GetDirectoryName(target);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            File.Copy(src, target, overwrite);
        }

        static void CopyFolder(string src, string target, bool overwrite)
        {
            Directory.CreateDirectory(target);
            foreach (var f in Directory.GetFiles(src))
                File.Copy(f, Path.Combine(target, Path.GetFileName(f)), overwrite);
            foreach (var d in Directory.GetDirectories(src))
            {
                if (!IsLink(d))
                    CopyFolder(d, Path.Combine(target, Path.GetFileName(d)), overwrite);
            }
        }

        static void Move(string src, string target, bool isFolder, bool overwrite)
        {
            var dir = Path.GetDirectoryName(target);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            if (!isFolder)
            {
                if (overwrite && File.Exists(target))
                    File.Delete(target);
                File.Move(src, target);
                return;
            }
            // Directory.Move cannot merge into an existing folder or cross drives.
            var sameDrive = string.Equals(Path.GetPathRoot(src), Path.GetPathRoot(target), StringComparison.OrdinalIgnoreCase);
            if (!overwrite && sameDrive)
            {
                Directory.Move(src, target);
                return;
            }
            CopyFolder(src, target, overwrite);
            Directory.Delete(src, true);
        }

        static void Delete(string path, bool isFolder, bool permanent)
        {
            if (permanent)
            {
                if (isFolder)
                    Directory.Delete(path, true);
                else
                    File.Delete(path);
                return;
            }
            if (isFolder)
                FileSystem.DeleteDirectory(path, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin, UICancelOption.ThrowException);
            else
                FileSystem.DeleteFile(path, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin, UICancelOption.ThrowException);
        }

        // name (2).ext, name (3).ext, ... free on disk and not taken by this run.
        static string NextFree(string target, bool isFolder, HashSet<string> taken)
        {
            var dir = Path.GetDirectoryName(target) ?? "";
            var name = Path.GetFileName(target);
            var ext = isFolder ? "" : Path.GetExtension(name);
            var stem = isFolder ? name : Path.GetFileNameWithoutExtension(name);
            for (var i = 2; ; i++)
            {
                var candidate = Path.Combine(dir, stem + " (" + i.ToString(CultureInfo.InvariantCulture) + ")" + ext);
                if (!taken.Contains(candidate) && !File.Exists(candidate) && !Directory.Exists(candidate))
                    return candidate;
            }
        }

        // ---------- formulas ----------

        public static string CheckFormula(string formula, bool required, string what)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return required ? what + " formula: enter a formula." : null;
            var error = ExpressionEngine.Validate(formula);
            return error == null ? null : what + " formula: " + error;
        }

        public static List<VariableHelp> FileVariableHelp(bool withKey, bool withColumns)
        {
            var list = new List<VariableHelp>();
            if (withKey)
            {
                list.Add(new VariableHelp { Name = "$Key", Description = "The key being matched (from the input's Key column, or its first column).", Example = "A100" });
                list.Add(new VariableHelp { Name = "$Value", Description = "The value next to the key, when the input is a map. Blank otherwise." });
            }
            if (withColumns)
                list.Add(new VariableHelp { Name = "$Key, $Relative, ...", Description = "Any column of the input row, by header name.", Example = "$Key" });
            list.Add(new VariableHelp { Name = "$Name", Description = "File or folder name, with extension.", Example = "A100_report.pdf" });
            list.Add(new VariableHelp { Name = "$Stem", Description = "Name without the extension.", Example = "A100_report" });
            list.Add(new VariableHelp { Name = "$Ext", Description = "Extension with the dot. Blank for a folder.", Example = ".pdf" });
            list.Add(new VariableHelp { Name = "$Path", Description = "Full path." });
            list.Add(new VariableHelp { Name = "$Folder", Description = "Folder that holds the item." });
            list.Add(new VariableHelp { Name = "$FolderName", Description = "Name of the folder that holds the item.", Example = "docs" });
            if (!withColumns)
                list.Add(new VariableHelp { Name = "$Relative", Description = "Folder path relative to the searched folder.", Example = "2024\\docs" });
            list.Add(new VariableHelp { Name = "$IsFolder", Description = "TRUE for a folder, FALSE for a file." });
            list.Add(new VariableHelp { Name = "$Size", Description = "File size in bytes. Blank for a folder." });
            list.Add(new VariableHelp { Name = "$Modified", Description = "Last write time." });
            list.Add(new VariableHelp { Name = "$1", Description = "First part of the stem split by the split character. $2 is the second.", Example = "A100" });
            return list;
        }

        static bool Matches(string formula, Func<string, object> lookup)
        {
            return Matches(ExpressionEngine.Compile(formula), lookup);
        }

        // A formula that fails for one item (say, a date function on a blank) does not match it.
        static bool Matches(Func<Func<string, object>, object> formula, Func<string, object> lookup)
        {
            try
            {
                return ExpressionEngine.ToBool(formula(lookup));
            }
            catch
            {
                return false;
            }
        }

        static string Text(Func<Func<string, object>, object> formula, Func<string, object> lookup)
        {
            try
            {
                return ExpressionEngine.ToText(formula(lookup));
            }
            catch
            {
                return "";
            }
        }

        static Func<string, object> KeyLookup(string key, string value)
        {
            return delegate(string name)
            {
                var n = (name ?? "").Trim();
                if (n.Equals("Key", StringComparison.OrdinalIgnoreCase))
                    return key;
                if (n.Equals("Value", StringComparison.OrdinalIgnoreCase))
                    return value;
                return null;
            };
        }

        // Columns by Excel letter ($A) or header ($Code), then built-ins.
        static Func<string, object> RowLookup(DataTable table, DataRow row)
        {
            return delegate(string name)
            {
                if (string.IsNullOrEmpty(name))
                    return null;
                int index;
                if (TryLetter(name, out index) && index < table.Columns.Count)
                    return row[index];
                var col = ExcelFile.FindColumn(table, name);
                if (col != null)
                    return row[col];
                object v;
                if (ExpressionEngine.TryGetVariable(name, out v))
                    return v;
                return null;
            };
        }

        static bool TryLetter(string s, out int index)
        {
            index = -1;
            var t = s.Trim();
            if (t.Length == 0 || t.Length > 3)
                return false;
            var n = 0;
            foreach (var ch in t)
            {
                var c = char.ToUpperInvariant(ch);
                if (c < 'A' || c > 'Z')
                    return false;
                n = n * 26 + (c - 'A' + 1);
            }
            index = n - 1;
            return true;
        }

        // File variables; extra (key or row columns) is asked first. root null: no $Relative of its own.
        static Func<string, object> FileLookup(FileItem item, string root, char separator, Func<string, object> extra)
        {
            return delegate(string name)
            {
                var key = name != null ? name.Trim() : "";
                if (extra != null && !IsFileVariable(key, root != null))
                {
                    var x = extra(key);
                    if (x != null)
                        return x;
                }
                object builtin;
                if (ExpressionEngine.TryGetVariable(key, out builtin))
                    return builtin;

                var fileName = Path.GetFileName(item.Path);
                var stem = item.IsFolder ? fileName : Path.GetFileNameWithoutExtension(fileName);
                int n;
                if (int.TryParse(key.TrimStart('$'), NumberStyles.Integer, CultureInfo.InvariantCulture, out n))
                {
                    var parts = stem.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
                    return n >= 1 && n <= parts.Length ? parts[n - 1] : "";
                }
                switch (key.ToUpperInvariant())
                {
                    case "NAME":
                    case "FILENAME":
                        return fileName;
                    case "STEM":
                    case "FILESTEM":
                        return stem;
                    case "EXT":
                        return item.IsFolder ? "" : Path.GetExtension(fileName);
                    case "PATH":
                        return item.Path;
                    case "FOLDER":
                        return Path.GetDirectoryName(item.Path) ?? "";
                    case "FOLDERNAME":
                    case "PARENTFOLDERNAME":
                        return FolderNameOf(item.Path);
                    case "ISFOLDER":
                        return item.IsFolder;
                    case "SIZE":
                        if (item.IsFolder)
                            return null;
                        try { return (double)new FileInfo(item.Path).Length; }
                        catch { return null; }
                    case "MODIFIED":
                        try { return item.IsFolder ? Directory.GetLastWriteTime(item.Path) : File.GetLastWriteTime(item.Path); }
                        catch { return null; }
                }
                if (root != null && key.Equals("Relative", StringComparison.OrdinalIgnoreCase))
                    return RelativeFolder(item.Path, root);
                return null;
            };
        }

        static bool IsFileVariable(string key, bool hasRoot)
        {
            switch (key.ToUpperInvariant())
            {
                case "NAME":
                case "FILENAME":
                case "STEM":
                case "FILESTEM":
                case "EXT":
                case "PATH":
                case "FOLDER":
                case "FOLDERNAME":
                case "PARENTFOLDERNAME":
                case "ISFOLDER":
                case "SIZE":
                case "MODIFIED":
                    return true;
                case "RELATIVE":
                    return hasRoot;
            }
            int n;
            return int.TryParse(key.TrimStart('$'), NumberStyles.Integer, CultureInfo.InvariantCulture, out n);
        }

        static char SeparatorOf(string s)
        {
            return string.IsNullOrEmpty(s) ? '_' : s[0];
        }

        static string FolderNameOf(string path)
        {
            var dir = Path.GetDirectoryName(path);
            if (string.IsNullOrEmpty(dir))
                return "";
            var name = Path.GetFileName(dir);
            return string.IsNullOrEmpty(name) ? dir : name;
        }

        static string RelativeFolder(string path, string root)
        {
            var dir = Path.GetDirectoryName(path);
            if (string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(root))
                return "";
            try
            {
                var r = Path.GetFullPath(root).TrimEnd('\\', '/');
                var d = Path.GetFullPath(dir);
                if (d.Equals(r, StringComparison.OrdinalIgnoreCase))
                    return "";
                if (d.StartsWith(r + "\\", StringComparison.OrdinalIgnoreCase))
                    return d.Substring(r.Length + 1);
                return d;
            }
            catch
            {
                return dir;
            }
        }

        // One or more folder names joined by \, each cleaned; never climbs out with "..".
        static string SafeSubPath(string text)
        {
            var parts = new List<string>();
            foreach (var raw in (text ?? "").Split('\\', '/'))
            {
                var part = Sanitize(raw).Trim('.', ' ');
                if (part.Length > 0)
                    parts.Add(part);
            }
            return string.Join("\\", parts.ToArray());
        }

        static string Sanitize(string name)
        {
            var sb = new StringBuilder(name.Length);
            foreach (var c in name)
            {
                if (c > 31 && Array.IndexOf(Illegal, c) < 0)
                    sb.Append(c);
            }
            return sb.ToString().Trim();
        }
    }
}
