using System;
using System.Collections.Generic;
using System.IO;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.MergeDuplicates
{
    public sealed class MergeDuplicatesEntry
    {
        public string FullPath { get; set; }
        public string Name { get; set; }
        public string[] Parts { get; set; }
    }

    public sealed class MergeDuplicatesGroup
    {
        public List<MergeDuplicatesEntry> Members { get; set; }

        public MergeDuplicatesGroup()
        {
            Members = new List<MergeDuplicatesEntry>();
        }
    }

    public sealed class MergeDuplicatesPlan
    {
        public string SourcePath { get; set; }
        public string TargetPath { get; set; }
        public string Status { get; set; }
        public int Moved { get; set; }
        public int Renamed { get; set; }
        public int Skipped { get; set; }
    }

    public static class MergeDuplicatesWork
    {
        static MergeDuplicatesWork()
        {
            var a = Entry("foo_bar", '_');
            var b = Entry("foo_zzz", '_');
            if (!Matches(a, b, "$A1 == $B1"))
                throw new InvalidOperationException("MergeDuplicatesWork match check failed.");
            if (Matches(a, b, "$A2 == $B2"))
                throw new InvalidOperationException("MergeDuplicatesWork reject check failed.");
            if (!Matches(a, b, "$A1 == $B1 && CONTAINS($A, \"foo\")"))
                throw new InvalidOperationException("MergeDuplicatesWork formula check failed.");
        }

        static MergeDuplicatesEntry Entry(string name, char separator)
        {
            return new MergeDuplicatesEntry { Name = name, Parts = SplitParts(name, separator) };
        }

        public static string[] SplitParts(string folderName, char separator)
        {
            if (string.IsNullOrEmpty(folderName))
                return new string[0];
            var raw = folderName.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
            var parts = new List<string>(raw.Length);
            foreach (var p in raw)
            {
                var t = p.Trim();
                if (t.Length > 0)
                    parts.Add(t);
            }
            return parts.ToArray();
        }

        public static List<MergeDuplicatesEntry> ListFolders(string parentDir, char separator)
        {
            var dirs = Directory.GetDirectories(parentDir);
            Array.Sort(dirs, StringComparer.OrdinalIgnoreCase);
            var list = new List<MergeDuplicatesEntry>(dirs.Length);
            foreach (var dir in dirs)
            {
                var name = Path.GetFileName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                list.Add(new MergeDuplicatesEntry
                {
                    FullPath = dir,
                    Name = name,
                    Parts = SplitParts(name, separator)
                });
            }
            return list;
        }

        public static void CheckCondition(string condition)
        {
            if (string.IsNullOrWhiteSpace(condition))
                throw new InvalidOperationException("Enter a formula such as $A1 == $B1.");
            var err = ExpressionEngine.Validate(condition);
            if (err != null)
                throw new InvalidOperationException(err);
        }

        public static bool Matches(MergeDuplicatesEntry a, MergeDuplicatesEntry b, string condition)
        {
            return ExpressionEngine.ToBool(ExpressionEngine.Evaluate(condition, Lookup(a, b)));
        }

        public static List<MergeDuplicatesGroup> FindGroups(string parentDir, char separator, string condition)
        {
            CheckCondition(condition);
            var entries = ListFolders(parentDir, separator);
            var n = entries.Count;
            var parent = new int[n];
            for (var i = 0; i < n; i++)
                parent[i] = i;

            for (var i = 0; i < n; i++)
            {
                for (var j = i + 1; j < n; j++)
                {
                    bool hit;
                    try
                    {
                        hit = Matches(entries[i], entries[j], condition);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Bad condition: " + ex.Message);
                    }
                    if (hit)
                        Union(parent, i, j);
                }
            }

            var buckets = new Dictionary<int, MergeDuplicatesGroup>();
            for (var i = 0; i < n; i++)
            {
                var root = Find(parent, i);
                MergeDuplicatesGroup g;
                if (!buckets.TryGetValue(root, out g))
                {
                    g = new MergeDuplicatesGroup();
                    buckets[root] = g;
                }
                g.Members.Add(entries[i]);
            }

            var groups = new List<MergeDuplicatesGroup>();
            foreach (var g in buckets.Values)
            {
                if (g.Members.Count < 2)
                    continue;
                g.Members.Sort(delegate (MergeDuplicatesEntry a, MergeDuplicatesEntry b)
                {
                    return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
                });
                groups.Add(g);
            }
            groups.Sort(delegate (MergeDuplicatesGroup a, MergeDuplicatesGroup b)
            {
                return string.Compare(a.Members[0].Name, b.Members[0].Name, StringComparison.OrdinalIgnoreCase);
            });
            return groups;
        }

        public static List<MergeDuplicatesPlan> BuildMergePlans(List<MergeDuplicatesGroup> groups, IList<string> targets)
        {
            return BuildMergePlans(groups, targets, null);
        }

        public static List<MergeDuplicatesPlan> BuildMergePlans(List<MergeDuplicatesGroup> groups, IList<string> targets, IList<IList<string>> selected)
        {
            var plans = new List<MergeDuplicatesPlan>();
            for (var gi = 0; gi < groups.Count; gi++)
            {
                var g = groups[gi];
                string target = targets != null && gi < targets.Count ? targets[gi] : null;
                if (string.IsNullOrEmpty(target))
                    target = g.Members[0].Name;
                HashSet<string> wanted = null;
                if (selected != null && gi < selected.Count && selected[gi] != null)
                {
                    wanted = new HashSet<string>(selected[gi], StringComparer.OrdinalIgnoreCase);
                }
                string targetPath = null;
                foreach (var m in g.Members)
                {
                    if (string.Equals(m.Name, target, StringComparison.OrdinalIgnoreCase))
                    {
                        targetPath = m.FullPath;
                        break;
                    }
                }
                if (targetPath == null)
                    throw new InvalidOperationException("Group " + (gi + 1) + ": target \"" + target + "\" is not in the group.");
                foreach (var m in g.Members)
                {
                    if (string.Equals(m.FullPath, targetPath, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (wanted != null && !wanted.Contains(m.Name))
                        continue;
                    plans.Add(new MergeDuplicatesPlan
                    {
                        SourcePath = m.FullPath,
                        TargetPath = targetPath,
                        Status = "Ready"
                    });
                }
            }
            return plans;
        }

        public static void ExecuteMerge(IList<MergeDuplicatesPlan> plans)
        {
            foreach (var plan in plans)
            {
                if (plan.Status != "Ready")
                    continue;
                try
                {
                    if (!Directory.Exists(plan.SourcePath))
                    {
                        plan.Status = "Skipped: missing";
                        continue;
                    }
                    Directory.CreateDirectory(plan.TargetPath);
                    int moved = 0;
                    int renamed = 0;
                    int skipped = 0;
                    MoveContents(plan.SourcePath, plan.TargetPath, ref moved, ref renamed, ref skipped);
                    DeleteEmptyDirs(plan.SourcePath);
                    try
                    {
                        if (Directory.Exists(plan.SourcePath) &&
                            Directory.GetFileSystemEntries(plan.SourcePath).Length == 0)
                            Directory.Delete(plan.SourcePath);
                    }
                    catch
                    {
                    }
                    plan.Moved = moved;
                    plan.Renamed = renamed;
                    plan.Skipped = skipped;
                    if (moved == 0 && renamed == 0 && skipped > 0)
                        plan.Status = "Skipped";
                    else if (skipped > 0)
                        plan.Status = "Merged (" + moved + " moved, " + renamed + " renamed, " + skipped + " skipped)";
                    else
                        plan.Status = "Merged (" + moved + " moved, " + renamed + " renamed)";
                }
                catch (Exception ex)
                {
                    plan.Status = "Failed: " + ex.Message;
                }
            }
        }

        static void MoveContents(string source, string target, ref int moved, ref int renamed, ref int skipped)
        {
            foreach (var file in Directory.GetFiles(source))
            {
                var name = Path.GetFileName(file);
                var dest = Path.Combine(target, name);
                if (string.Equals(Path.GetFullPath(file), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase))
                {
                    skipped++;
                    continue;
                }
                if (File.Exists(dest))
                {
                    var alt = NextFreeFile(target, name);
                    if (alt == null)
                    {
                        skipped++;
                        continue;
                    }
                    try
                    {
                        File.Move(file, alt);
                        renamed++;
                    }
                    catch
                    {
                        skipped++;
                    }
                    continue;
                }
                try
                {
                    File.Move(file, dest);
                    moved++;
                }
                catch
                {
                    skipped++;
                }
            }
            foreach (var dir in Directory.GetDirectories(source))
            {
                var name = Path.GetFileName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                var dest = Path.Combine(target, name);
                if (string.Equals(Path.GetFullPath(dir), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase))
                {
                    skipped++;
                    continue;
                }
                if (Directory.Exists(dest))
                {
                    MoveContents(dir, dest, ref moved, ref renamed, ref skipped);
                    continue;
                }
                try
                {
                    Directory.Move(dir, dest);
                    moved++;
                }
                catch
                {
                    try
                    {
                        Directory.CreateDirectory(dest);
                        MoveContents(dir, dest, ref moved, ref renamed, ref skipped);
                    }
                    catch
                    {
                        skipped++;
                    }
                }
            }
        }

        static string NextFreeFile(string targetDir, string fileName)
        {
            var stem = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            for (var i = 2; i < 1000; i++)
            {
                var candidate = Path.Combine(targetDir, stem + " (" + i + ")" + ext);
                if (!File.Exists(candidate) && !Directory.Exists(candidate))
                    return candidate;
            }
            return null;
        }

        static void DeleteEmptyDirs(string root)
        {
            string[] dirs;
            try
            {
                dirs = Directory.GetDirectories(root);
            }
            catch
            {
                return;
            }
            foreach (var dir in dirs)
            {
                DeleteEmptyDirs(dir);
                try
                {
                    if (Directory.GetFileSystemEntries(dir).Length == 0)
                        Directory.Delete(dir);
                }
                catch
                {
                }
            }
        }

        static int Find(int[] parent, int x)
        {
            while (parent[x] != x)
            {
                parent[x] = parent[parent[x]];
                x = parent[x];
            }
            return x;
        }

        static void Union(int[] parent, int a, int b)
        {
            var ra = Find(parent, a);
            var rb = Find(parent, b);
            if (ra != rb)
                parent[rb] = ra;
        }

        static Func<string, object> Lookup(MergeDuplicatesEntry a, MergeDuplicatesEntry b)
        {
            return delegate(string name)
            {
                object builtin;
                if (ExpressionEngine.TryGetVariable(name, out builtin))
                    return builtin;

                var key = name != null ? name.Trim() : "";
                if (key.Length == 0)
                    return null;
                if (key[0] == '$')
                    key = key.Substring(1);

                if (key.Equals("A", StringComparison.OrdinalIgnoreCase) ||
                    key.Equals("NameA", StringComparison.OrdinalIgnoreCase))
                    return a != null ? a.Name : "";
                if (key.Equals("B", StringComparison.OrdinalIgnoreCase) ||
                    key.Equals("NameB", StringComparison.OrdinalIgnoreCase))
                    return b != null ? b.Name : "";

                return LookupPart(key, a, b);
            };
        }

        static object LookupPart(string key, MergeDuplicatesEntry a, MergeDuplicatesEntry b)
        {
            if (string.IsNullOrEmpty(key))
                return null;
            char side = char.ToUpperInvariant(key[0]);
            string digits = key;
            MergeDuplicatesEntry entry = a;
            if (side == 'A' || side == 'B')
            {
                digits = key.Substring(1);
                entry = side == 'A' ? a : b;
            }
            int n;
            if (!int.TryParse(digits, out n) || n < 1)
                return null;
            if (entry == null || entry.Parts == null || n > entry.Parts.Length)
                return "";
            return entry.Parts[n - 1];
        }
    }
}
