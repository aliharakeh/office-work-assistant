using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace OfficeWorkAssistant.Features.Pipeline
{
    public sealed class SavedPipeline
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public DateTime Modified { get; set; }

        public string ModifiedText
        {
            get { return Modified.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture); }
        }
    }

    // Saved pipelines live as <name>.xml in Documents\Office Work Assistant\Pipelines.
    // The file name is the pipeline's name. No WPF.
    public static class PipelineStore
    {
        static readonly char[] Illegal = System.IO.Path.GetInvalidFileNameChars();

        public static string Folder
        {
            get
            {
                return System.IO.Path.Combine(System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Office Work Assistant"), "Pipelines");
            }
        }

        // Newest first.
        public static List<SavedPipeline> List()
        {
            var list = new List<SavedPipeline>();
            if (!Directory.Exists(Folder))
                return list;
            foreach (var path in Directory.GetFiles(Folder, "*.xml"))
            {
                list.Add(new SavedPipeline
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(path),
                    Path = path,
                    Modified = File.GetLastWriteTime(path)
                });
            }
            list.Sort((a, b) => b.Modified.CompareTo(a.Modified));
            return list;
        }

        public static string CheckName(string name)
        {
            var n = (name ?? "").Trim();
            if (n.Length == 0)
                return "Give the pipeline a name.";
            if (n.IndexOfAny(Illegal) >= 0)
                return "A name cannot contain \\ / : * ? \" < > |";
            if (n.EndsWith(".", StringComparison.Ordinal))
                return "A name cannot end with a dot.";
            return null;
        }

        public static string PathFor(string name)
        {
            return System.IO.Path.Combine(Folder, name.Trim() + ".xml");
        }

        // Writes def under name. oldPath is the file it was opened from (or null): when the
        // name changed, that file is removed so a rename does not leave a copy behind.
        public static string Save(PipelineDefinition def, string name, string oldPath)
        {
            var error = CheckName(name);
            if (error != null)
                throw new InvalidOperationException(error);
            Directory.CreateDirectory(Folder);
            var path = PathFor(name);
            def.Name = name.Trim();
            // Write beside, then swap, so a failed save never damages the old file.
            var temp = path + ".tmp";
            PipelineWork.SaveFile(def, temp);
            if (File.Exists(path))
                File.Delete(path);
            File.Move(temp, path);
            if (!string.IsNullOrEmpty(oldPath) && !string.Equals(System.IO.Path.GetFullPath(oldPath), System.IO.Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase) &&
                File.Exists(oldPath))
                File.Delete(oldPath);
            return path;
        }

        public static PipelineDefinition Load(string path)
        {
            var def = PipelineWork.LoadFile(path);
            def.Name = System.IO.Path.GetFileNameWithoutExtension(path);
            return def;
        }

        // "name (copy)", "name (copy 2)", ... not yet taken.
        public static string Duplicate(string path)
        {
            var stem = System.IO.Path.GetFileNameWithoutExtension(path) + " (copy";
            var name = stem + ")";
            for (var i = 2; File.Exists(PathFor(name)); i++)
                name = stem + " " + i.ToString(CultureInfo.InvariantCulture) + ")";
            File.Copy(path, PathFor(name));
            return PathFor(name);
        }

        // To the Recycle Bin, so a wrong click can be undone.
        public static void Delete(string path)
        {
            FileSystem.DeleteFile(path, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin, UICancelOption.ThrowException);
        }

        // Copies a pipeline file from elsewhere into the folder; returns its new path.
        public static string Import(string source)
        {
            PipelineWork.LoadFile(source); // throws when it is not a pipeline
            Directory.CreateDirectory(Folder);
            var stem = System.IO.Path.GetFileNameWithoutExtension(source);
            var name = stem;
            for (var i = 2; File.Exists(PathFor(name)); i++)
                name = stem + " (" + i.ToString(CultureInfo.InvariantCulture) + ")";
            File.Copy(source, PathFor(name));
            return PathFor(name);
        }
    }
}
