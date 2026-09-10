using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using WorkAssistant.Expressions;

namespace WorkAssistant.Features.CopyFiles
{
    public sealed class CopyFilesPlan
    {
        public string Source { get; set; }
        public string WrapFolder { get; set; }
        public string NewName { get; set; }
        public string DestPath { get; set; }
        public string Status { get; set; }
    }

    public static class CopyFilesWork
    {
        static readonly char[] Illegal = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

        static CopyFilesWork()
        {
            var lookup = FileLookup(@"C:\src\docs\report_2024_final.pdf", @"C:\src", '_');
            var name = EvalName("$3 & \"-\" & $1", lookup);
            if (name != "final-report")
                throw new InvalidOperationException("CopyFilesWork name check failed: " + name);
            if (!MatchesFilter("CONTAINS({Name}, \"report\") && {Ext} == \".pdf\"", lookup))
                throw new InvalidOperationException("CopyFilesWork file filter check failed.");
            if (MatchesFilter("STARTSWITH({Name}, \"inv\")", lookup))
                throw new InvalidOperationException("CopyFilesWork file filter reject check failed.");
            var folder = FolderLookup(@"C:\src\a-b\report.pdf", @"C:\src", '-');
            if (!MatchesFilter("CONTAINS({Relative}, \"a-b\") && $1 == \"a\" && $2 == \"b\"", folder))
                throw new InvalidOperationException("CopyFilesWork folder filter check failed.");
            var wrap = EvalName("$2 & \"-\" & $1", folder);
            if (wrap != "b-a")
                throw new InvalidOperationException("CopyFilesWork folder name check failed: " + wrap);
        }

        public static string ValidateExpression(string expression, bool required)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return required ? "enter a formula." : null;
            return ExpressionEngine.Validate(expression);
        }

        public static List<CopyFilesPlan> PlanMoves(
            string sourceDir, string destDir, char fileSep, char folderSep, string filePattern, bool recursive, bool wrapFolder,
            string folderPattern = null, string fileFilter = null, string folderFilter = null)
        {
            var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(sourceDir, "*", option);
            var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var plans = new List<CopyFilesPlan>(files.Length);

            foreach (var source in files)
            {
                var fileLookup = FileLookup(source, sourceDir, fileSep);
                var folderLookup = FolderLookup(source, sourceDir, folderSep);
                if (!MatchesFilter(fileFilter, fileLookup))
                    continue;
                if (!MatchesFilter(folderFilter, folderLookup))
                    continue;

                var ext = Path.GetExtension(source);
                var newName = EvalName(filePattern, fileLookup);
                var wrapName = wrapFolder ? EvalName(folderPattern, folderLookup) : "";
                var destPath = "";
                var status = "Ready";

                if (string.IsNullOrEmpty(newName) || (wrapFolder && string.IsNullOrEmpty(wrapName)))
                {
                    status = "Bad name";
                }
                else
                {
                    destPath = wrapFolder
                        ? Path.Combine(destDir, wrapName, newName + ext)
                        : Path.Combine(destDir, newName + ext);

                    if (string.Equals(Path.GetFullPath(source), Path.GetFullPath(destPath), StringComparison.OrdinalIgnoreCase))
                        status = "Same path";
                    else if (File.Exists(destPath) || !reserved.Add(destPath))
                        status = "Exists";
                }

                plans.Add(new CopyFilesPlan
                {
                    Source = source,
                    WrapFolder = wrapName,
                    NewName = newName,
                    DestPath = destPath,
                    Status = status
                });
            }

            return plans;
        }

        public static void ApplyMoves(IList<CopyFilesPlan> plans)
        {
            foreach (var plan in plans)
            {
                if (plan.Status != "Ready")
                    continue;
                try
                {
                    var dir = Path.GetDirectoryName(plan.DestPath);
                    if (!string.IsNullOrEmpty(dir))
                        Directory.CreateDirectory(dir);
                    File.Copy(plan.Source, plan.DestPath);
                    plan.Status = "Copied";
                }
                catch (Exception ex)
                {
                    plan.Status = "Failed: " + ex.Message;
                }
            }
        }

        static string EvalName(string expression, Func<string, object> lookup)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return "";
            try
            {
                return Sanitize(ExpressionEngine.ToText(ExpressionEngine.Evaluate(expression, lookup)));
            }
            catch
            {
                return "";
            }
        }

        static bool MatchesFilter(string expression, Func<string, object> lookup)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return true;
            try
            {
                return ExpressionEngine.ToBool(ExpressionEngine.Evaluate(expression, lookup));
            }
            catch
            {
                return false;
            }
        }

        static Func<string, object> FileLookup(string source, string sourceDir, char separator)
        {
            return delegate(string name)
            {
                object builtin;
                if (ExpressionEngine.TryGetVariable(name, out builtin))
                    return builtin;

                var fileName = Path.GetFileName(source);
                var stem = Path.GetFileNameWithoutExtension(fileName);
                var part = LookupPart(stem, separator, name);
                if (part != null)
                    return part;

                var key = name != null ? name.Trim() : "";
                if (key.Equals("Name", StringComparison.OrdinalIgnoreCase))
                    return fileName;
                if (key.Equals("Stem", StringComparison.OrdinalIgnoreCase))
                    return stem;
                if (key.Equals("Ext", StringComparison.OrdinalIgnoreCase))
                    return Path.GetExtension(fileName);
                if (key.Equals("Path", StringComparison.OrdinalIgnoreCase))
                    return source;
                if (key.Equals("Folder", StringComparison.OrdinalIgnoreCase))
                    return Path.GetDirectoryName(source);
                if (key.Equals("FolderName", StringComparison.OrdinalIgnoreCase))
                    return FolderNameOf(source);
                if (key.Equals("Relative", StringComparison.OrdinalIgnoreCase))
                    return RelativeFolder(source, sourceDir);
                if (key.Equals("Size", StringComparison.OrdinalIgnoreCase))
                {
                    try { return (double)new FileInfo(source).Length; }
                    catch { return 0d; }
                }
                if (key.Equals("Modified", StringComparison.OrdinalIgnoreCase))
                {
                    try { return File.GetLastWriteTime(source); }
                    catch { return null; }
                }
                return null;
            };
        }

        static Func<string, object> FolderLookup(string source, string sourceDir, char separator)
        {
            return delegate(string name)
            {
                object builtin;
                if (ExpressionEngine.TryGetVariable(name, out builtin))
                    return builtin;

                var dir = Path.GetDirectoryName(source);
                var folderName = FolderNameOf(source);
                var part = LookupPart(folderName, separator, name);
                if (part != null)
                    return part;

                var key = name != null ? name.Trim() : "";
                if (key.Equals("Name", StringComparison.OrdinalIgnoreCase) ||
                    key.Equals("FolderName", StringComparison.OrdinalIgnoreCase))
                    return folderName;
                if (key.Equals("Path", StringComparison.OrdinalIgnoreCase) ||
                    key.Equals("Folder", StringComparison.OrdinalIgnoreCase))
                    return dir;
                if (key.Equals("Relative", StringComparison.OrdinalIgnoreCase))
                    return RelativeFolder(source, sourceDir);
                return null;
            };
        }

        static object LookupPart(string stem, char separator, string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            var t = name.Trim();
            if (t.Length > 0 && t[0] == '$')
                t = t.Substring(1);
            int n;
            if (!int.TryParse(t, out n))
                return null;
            var parts = (stem ?? "").Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
            if (n < 1 || n > parts.Length)
                return "";
            return parts[n - 1];
        }

        static string FolderNameOf(string sourcePath)
        {
            var dir = Path.GetDirectoryName(sourcePath);
            if (string.IsNullOrEmpty(dir))
                return "";
            var name = Path.GetFileName(dir);
            return string.IsNullOrEmpty(name) ? dir : name;
        }

        static string RelativeFolder(string sourcePath, string sourceDir)
        {
            var dir = Path.GetDirectoryName(sourcePath);
            if (string.IsNullOrEmpty(dir))
                return "";
            try
            {
                var root = Path.GetFullPath(sourceDir);
                var fullDir = Path.GetFullPath(dir);
                root = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (fullDir.Equals(root, StringComparison.OrdinalIgnoreCase))
                    return "";
                if (fullDir.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                    fullDir.StartsWith(root + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    return fullDir.Substring(root.Length + 1);
                return fullDir;
            }
            catch
            {
                return dir;
            }
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
