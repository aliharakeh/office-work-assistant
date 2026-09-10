using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WorkAssistant.Features.CopyFiles
{
    public sealed class CopyFilesPlan
    {
        public string Source { get; set; }
        public string NewName { get; set; }
        public string DestPath { get; set; }
        public string Status { get; set; }
    }

    public static class CopyFilesWork
    {
        static readonly Regex Token = new Regex(@"\$(\d+)", RegexOptions.Compiled);
        static readonly char[] Illegal = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

        static CopyFilesWork()
        {
            var name = BuildName("report_2024_final.pdf", '_', "$3-$1");
            if (name != "final-report")
                throw new InvalidOperationException("CopyFilesWork name check failed: " + name);
        }

        public static string BuildName(string fileName, char separator, string pattern)
        {
            var stem = Path.GetFileNameWithoutExtension(fileName);
            var parts = stem.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
            var raw = Token.Replace(pattern ?? "", m =>
            {
                var n = int.Parse(m.Groups[1].Value);
                if (n < 1 || n > parts.Length)
                    return "";
                return parts[n - 1];
            });
            return Sanitize(raw);
        }

        public static List<CopyFilesPlan> PlanMoves(
            string sourceDir, string destDir, char separator, string pattern, bool recursive, bool wrapFolder,
            string extensionFilter = null, string nameContains = null, bool folderOnly = false,
            string folderContains = null)
        {
            var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(sourceDir, "*", option);
            var allowedExts = ParseExtensions(extensionFilter);
            var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var plans = new List<CopyFilesPlan>(files.Length);

            foreach (var source in files)
            {
                if (!MatchesFilters(Path.GetFileName(source), allowedExts, nameContains))
                    continue;
                if (!MatchesFolder(source, sourceDir, folderContains))
                    continue;

                var ext = Path.GetExtension(source);
                var newName = BuildName(Path.GetFileName(source), separator, pattern);
                var destPath = "";
                var status = "Ready";

                if (string.IsNullOrEmpty(newName))
                {
                    status = "Bad name";
                }
                else if (folderOnly && wrapFolder)
                {
                    destPath = Path.Combine(destDir, newName, Path.GetFileName(source));

                    if (string.Equals(Path.GetFullPath(source), Path.GetFullPath(destPath), StringComparison.OrdinalIgnoreCase))
                        status = "Same path";
                    else if (File.Exists(destPath) || !reserved.Add(destPath))
                        status = "Exists";
                }
                else
                {
                    destPath = wrapFolder
                        ? Path.Combine(destDir, newName, newName + ext)
                        : Path.Combine(destDir, newName + ext);

                    if (string.Equals(Path.GetFullPath(source), Path.GetFullPath(destPath), StringComparison.OrdinalIgnoreCase))
                        status = "Same path";
                    else if (File.Exists(destPath) || !reserved.Add(destPath))
                        status = "Exists";
                }

                plans.Add(new CopyFilesPlan
                {
                    Source = source,
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

        static HashSet<string> ParseExtensions(string filter)
        {
            if (string.IsNullOrWhiteSpace(filter))
                return null;
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var tokens = filter.Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var raw in tokens)
            {
                var t = raw.Trim().TrimStart('*').Trim();
                if (t.Length == 0)
                    continue;
                if (!t.StartsWith(".", StringComparison.Ordinal))
                    t = "." + t;
                set.Add(t);
            }
            return set.Count == 0 ? null : set;
        }

        static bool MatchesFilters(string fileName, HashSet<string> allowedExts, string nameContains)
        {
            if (allowedExts != null && !allowedExts.Contains(Path.GetExtension(fileName)))
                return false;
            if (!string.IsNullOrWhiteSpace(nameContains) &&
                fileName.IndexOf(nameContains.Trim(), StringComparison.OrdinalIgnoreCase) < 0)
                return false;
            return true;
        }

        static bool MatchesFolder(string sourcePath, string sourceDir, string folderContains)
        {
            if (string.IsNullOrWhiteSpace(folderContains))
                return true;
            var needle = folderContains.Trim();
            if (needle.Length == 0)
                return true;

            var dir = Path.GetDirectoryName(sourcePath);
            if (string.IsNullOrEmpty(dir))
                return false;

            string relative;
            try
            {
                var root = Path.GetFullPath(sourceDir);
                var fullDir = Path.GetFullPath(dir);
                root = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (fullDir.Equals(root, StringComparison.OrdinalIgnoreCase))
                {
                    relative = "";
                }
                else if (fullDir.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                         fullDir.StartsWith(root + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                {
                    relative = fullDir.Substring(root.Length + 1);
                }
                else
                {
                    relative = fullDir;
                }
            }
            catch
            {
                relative = dir;
            }

            if (relative.Length == 0)
            {
                var folderName = Path.GetFileName(dir);
                if (string.IsNullOrEmpty(folderName))
                    folderName = dir;
                return folderName.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return relative.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
