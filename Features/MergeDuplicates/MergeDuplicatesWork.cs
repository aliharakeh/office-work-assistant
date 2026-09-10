using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WorkAssistant.Features.MergeDuplicates
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
                throw new InvalidOperationException("Enter a condition such as $A1 == $B1.");
            var parser = new ConditionParser(condition);
            parser.Parse();
        }

        public static bool Matches(string[] partsA, string[] partsB, string condition)
        {
            var parser = new ConditionParser(condition);
            var root = parser.Parse();
            return root.Eval(partsA, partsB);
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
                        hit = Matches(entries[i].Parts, entries[j].Parts, condition);
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

        // ponytail: tiny recursive-descent parser, no expression lib to stay on net48 BCL.
        sealed class ConditionParser
        {
            readonly List<Token> _tokens;
            int _pos;

            public ConditionParser(string text)
            {
                _tokens = Tokenize(text);
            }

            public Node Parse()
            {
                var node = ParseOr();
                if (_pos != _tokens.Count)
                    throw new InvalidOperationException("Unexpected \"" + _tokens[_pos].Text + "\". Use == != && || ( ).");
                return node;
            }

            Node ParseOr()
            {
                var left = ParseAnd();
                while (Match("||"))
                    left = new OrNode(left, ParseAnd());
                return left;
            }

            Node ParseAnd()
            {
                var left = ParseUnary();
                while (Match("&&"))
                    left = new AndNode(left, ParseUnary());
                return left;
            }

            Node ParseUnary()
            {
                if (Match("!"))
                    return new NotNode(ParseUnary());
                return ParsePrimary();
            }

            Node ParsePrimary()
            {
                if (Match("("))
                {
                    var inner = ParseOr();
                    Expect(")");
                    return inner;
                }
                var left = ParseOperand();
                string op;
                if (Match("=="))
                    op = "==";
                else if (Match("!="))
                    op = "!=";
                else
                    throw new InvalidOperationException("Expected == or != after \"" + left.Describe() + "\".");
                var right = ParseOperand();
                return new CompareNode(left, op, right);
            }

            Operand ParseOperand()
            {
                if (_pos >= _tokens.Count)
                    throw new InvalidOperationException("Incomplete condition, expected a value.");
                var t = _tokens[_pos++];
                if (t.Kind == TokenKind.Part)
                    return new Operand { Side = t.Side, Index = t.Index };
                if (t.Kind == TokenKind.Literal)
                    return new Operand { Literal = t.Text };
                throw new InvalidOperationException("Unexpected \"" + t.Text + "\".");
            }

            bool Match(string op)
            {
                if (_pos < _tokens.Count && _tokens[_pos].Kind == TokenKind.Op && _tokens[_pos].Text == op)
                {
                    _pos++;
                    return true;
                }
                return false;
            }

            void Expect(string op)
            {
                if (!Match(op))
                    throw new InvalidOperationException("Expected \"" + op + "\".");
            }

            enum TokenKind { Op, Part, Literal }

            sealed class Token
            {
                public TokenKind Kind;
                public string Text;
                public char Side;
                public int Index;
            }

            static List<Token> Tokenize(string text)
            {
                var tokens = new List<Token>();
                var i = 0;
                while (i < text.Length)
                {
                    var c = text[i];
                    if (char.IsWhiteSpace(c))
                    {
                        i++;
                        continue;
                    }
                    if (c == '(' || c == ')')
                    {
                        tokens.Add(new Token { Kind = TokenKind.Op, Text = c.ToString() });
                        i++;
                        continue;
                    }
                    if (c == '!' || c == '=')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '=')
                        {
                            tokens.Add(new Token { Kind = TokenKind.Op, Text = text.Substring(i, 2) });
                            i += 2;
                            continue;
                        }
                        if (c == '!')
                        {
                            tokens.Add(new Token { Kind = TokenKind.Op, Text = "!" });
                            i++;
                            continue;
                        }
                        throw new InvalidOperationException("Single = is not valid, use ==.");
                    }
                    if (c == '&' || c == '|')
                    {
                        if (i + 1 < text.Length && text[i + 1] == c)
                        {
                            tokens.Add(new Token { Kind = TokenKind.Op, Text = text.Substring(i, 2) });
                            i += 2;
                            continue;
                        }
                        throw new InvalidOperationException("Use && and ||, not single " + c + ".");
                    }
                    if (c == '"' || c == '\'')
                    {
                        var sb = new StringBuilder();
                        i++;
                        bool closed = false;
                        while (i < text.Length)
                        {
                            if (text[i] == c)
                            {
                                closed = true;
                                i++;
                                break;
                            }
                            sb.Append(text[i]);
                            i++;
                        }
                        if (!closed)
                            throw new InvalidOperationException("Unclosed quote in condition.");
                        tokens.Add(new Token { Kind = TokenKind.Literal, Text = sb.ToString() });
                        continue;
                    }
                    char partSide;
                    int partIndex;
                    int partLen;
                    if ((c == '$' || c == 'A' || c == 'a' || c == 'B' || c == 'b') && TryReadPart(text, i, out partSide, out partIndex, out partLen))
                    {
                        tokens.Add(new Token { Kind = TokenKind.Part, Text = "$" + partSide + partIndex, Side = partSide, Index = partIndex });
                        i += partLen;
                        continue;
                    }
                    var start = i;
                    while (i < text.Length && !char.IsWhiteSpace(text[i]) && "()!=\"'&|=".IndexOf(text[i]) < 0)
                        i++;
                    if (i == start)
                        throw new InvalidOperationException("Unexpected character '" + text[i] + "'.");
                    tokens.Add(new Token { Kind = TokenKind.Literal, Text = text.Substring(start, i - start) });
                }
                if (tokens.Count == 0)
                    throw new InvalidOperationException("Enter a condition such as $A1 == $B1.");
                return tokens;
            }

            static bool TryReadPart(string text, int at, out char side, out int index, out int len)
            {
                side = '\0';
                index = 0;
                len = 0;
                var i = at;
                if (text[i] == '$')
                    i++;
                else
                    return false;
                if (i >= text.Length || (text[i] != 'A' && text[i] != 'a' && text[i] != 'B' && text[i] != 'b'))
                    return false;
                side = char.ToUpperInvariant(text[i]);
                i++;
                var d = i;
                while (i < text.Length && char.IsDigit(text[i]))
                    i++;
                if (i == d)
                    return false;
                if (!int.TryParse(text.Substring(d, i - d), out index) || index < 1)
                    return false;
                len = i - at;
                return true;
            }
        }

        abstract class Node
        {
            public abstract bool Eval(string[] a, string[] b);
        }

        sealed class OrNode : Node
        {
            readonly Node _l;
            readonly Node _r;
            public OrNode(Node l, Node r) { _l = l; _r = r; }
            public override bool Eval(string[] a, string[] b) { return _l.Eval(a, b) || _r.Eval(a, b); }
        }

        sealed class AndNode : Node
        {
            readonly Node _l;
            readonly Node _r;
            public AndNode(Node l, Node r) { _l = l; _r = r; }
            public override bool Eval(string[] a, string[] b) { return _l.Eval(a, b) && _r.Eval(a, b); }
        }

        sealed class NotNode : Node
        {
            readonly Node _inner;
            public NotNode(Node inner) { _inner = inner; }
            public override bool Eval(string[] a, string[] b) { return !_inner.Eval(a, b); }
        }

        sealed class CompareNode : Node
        {
            readonly Operand _l;
            readonly string _op;
            readonly Operand _r;
            public CompareNode(Operand l, string op, Operand r) { _l = l; _op = op; _r = r; }
            public override bool Eval(string[] a, string[] b)
            {
                var lv = _l.Resolve(a, b, true);
                var rv = _r.Resolve(a, b, false);
                var eq = string.Equals(lv.Trim(), rv.Trim(), StringComparison.OrdinalIgnoreCase);
                return _op == "==" ? eq : !eq;
            }
        }

        sealed class Operand
        {
            public char Side;
            public int Index;
            public string Literal;
            public bool IsPart { get { return Side == 'A' || Side == 'B'; } }

            public string Resolve(string[] a, string[] b, bool isLeft)
            {
                if (!IsPart)
                    return Literal ?? "";
                var parts = Side == 'A' ? a : b;
                if (parts == null || Index < 1 || Index > parts.Length)
                    return "";
                return parts[Index - 1] ?? "";
            }

            public string Describe()
            {
                if (IsPart)
                    return "$" + Side + Index;
                return Literal ?? "";
            }
        }
    }
}
