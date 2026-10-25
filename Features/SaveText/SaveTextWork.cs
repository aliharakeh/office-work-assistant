using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.SaveText
{
    public enum SaveTextFormat
    {
        Text,
        TextMap,
        JsonList,
        JsonMap
    }

    public sealed class SaveTextSettings
    {
        public string Path { get; set; }
        public SaveTextFormat Format { get; set; }
        public string Column { get; set; }
        public string ValueColumn { get; set; }

        public SaveTextSettings()
        {
            Path = "";
            Column = "";
            ValueColumn = "";
        }
    }

    public static class SaveTextWork
    {
        public static bool IsMap(SaveTextFormat format)
        {
            return format == SaveTextFormat.TextMap || format == SaveTextFormat.JsonMap;
        }

        public static bool IsJson(SaveTextFormat format)
        {
            return format == SaveTextFormat.JsonList || format == SaveTextFormat.JsonMap;
        }

        public static string FormatLabel(SaveTextFormat format)
        {
            switch (format)
            {
                case SaveTextFormat.Text: return "Text column";
                case SaveTextFormat.TextMap: return "Text map (key + tab + value)";
                case SaveTextFormat.JsonList: return "JSON array (list / set)";
                case SaveTextFormat.JsonMap: return "JSON object (map)";
                default: throw new InvalidOperationException("Choose a supported text / JSON format.");
            }
        }

        public static string OutputPath(SaveTextSettings settings)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.Path))
                throw new InvalidOperationException("Choose where to save the text / JSON file.");
            FormatLabel(settings.Format);
            string path = settings.Path.Trim();
            if (!Path.IsPathRooted(path) || Path.GetPathRoot(path).Length < 3 ||
                path.StartsWith(@"\\.\", StringComparison.Ordinal) || path.StartsWith(@"\\?\", StringComparison.Ordinal))
                throw new InvalidOperationException("Use a full output file path.");
            if (Directory.Exists(path) || string.IsNullOrEmpty(Path.GetFileName(path)))
                throw new InvalidOperationException("Choose a file, not a folder.");
            if (Path.GetFileName(path).IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || path.EndsWith(".", StringComparison.Ordinal))
                throw new InvalidOperationException("The output filename contains invalid characters or ends with a dot.");
            if (string.IsNullOrEmpty(Path.GetExtension(path)))
                path += IsJson(settings.Format) ? ".json" : ".txt";
            path = Path.GetFullPath(path);
            if (!Directory.Exists(Path.GetDirectoryName(path)))
                throw new InvalidOperationException("The output folder does not exist: " + Path.GetDirectoryName(path));
            return path;
        }

        public static void Validate(DataTable input, SaveTextSettings settings)
        {
            if (input == null)
                throw new InvalidOperationException("Connect a table, OCR result or list / set / map first.");
            if (settings == null)
                throw new InvalidOperationException("Set up Save text / JSON first.");
            FormatLabel(settings.Format);
            ExcelColumns.Index(input, settings.Column);
            if (IsMap(settings.Format))
            {
                int value = ExcelColumns.Index(input, settings.ValueColumn);
                if (value == ExcelColumns.Index(input, settings.Column))
                    throw new InvalidOperationException("Choose different key and value columns for the map.");
            }
            OutputPath(settings);
        }

        // Pure preview: produce the exact UTF-8 content, without creating or replacing any files.
        public static DataTable Run(DataTable input, SaveTextSettings settings)
        {
            Validate(input, settings);
            int column = ExcelColumns.Index(input, settings.Column);
            int valueColumn = IsMap(settings.Format) ? ExcelColumns.Index(input, settings.ValueColumn) : -1;
            string content;
            if (!IsJson(settings.Format))
            {
                var text = new StringBuilder();
                for (var r = 0; r < input.Rows.Count; r++)
                {
                    if (r > 0)
                        text.Append(Environment.NewLine);
                    text.Append(Text(input.Rows[r][column]));
                    if (valueColumn >= 0)
                        text.Append('\t').Append(Text(input.Rows[r][valueColumn]));
                }
                content = text.ToString();
            }
            else
            {
                var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                if (settings.Format == SaveTextFormat.JsonList)
                {
                    var values = new List<object>();
                    foreach (DataRow row in input.Rows)
                        values.Add(JsonValue(row[column]));
                    content = json.Serialize(values);
                }
                else
                {
                    var values = new Dictionary<string, object>(StringComparer.Ordinal);
                    for (var r = 0; r < input.Rows.Count; r++)
                    {
                        object keyValue = input.Rows[r][column];
                        if (keyValue == DBNull.Value || keyValue == null)
                            throw new InvalidOperationException("JSON map row " + (r + 1) + " has a blank/null key.");
                        string key = Text(keyValue);
                        if (values.ContainsKey(key))
                            throw new InvalidOperationException("JSON map has duplicate key '" + key + "'. Use a map step to resolve duplicates first.");
                        values.Add(key, JsonValue(input.Rows[r][valueColumn]));
                    }
                    content = json.Serialize(values);
                }
            }

            // Validate encoding during preview, before any pipeline disk actions can run.
            new UTF8Encoding(false, true).GetByteCount(content);
            var plan = new DataTable();
            plan.Locale = CultureInfo.InvariantCulture;
            ExcelFile.AddColumn(plan, "Path");
            ExcelFile.AddColumn(plan, "Format");
            ExcelFile.AddColumn(plan, "Rows");
            ExcelFile.AddColumn(plan, "Content");
            ExcelFile.AddColumn(plan, "Status");
            string path = OutputPath(settings);
            plan.Rows.Add(path, FormatLabel(settings.Format), input.Rows.Count, content, File.Exists(path) ? "Overwrite" : "Ready");
            return plan;
        }

        static string Text(object value)
        {
            if (value == null || value == DBNull.Value)
                return "";
            if (value is DateTime)
                return ((DateTime)value).ToString("o", CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        static object JsonValue(object value)
        {
            if (value == null || value == DBNull.Value)
                return null;
            // Keep simple table values typed; dates and other display values are strings.
            if (value is bool || value is byte || value is sbyte || value is short || value is ushort ||
                value is int || value is uint || value is long || value is ulong || value is decimal)
                return value;
            if (value is double || value is float)
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(number) || double.IsInfinity(number))
                    throw new InvalidOperationException("JSON cannot contain NaN or Infinity.");
                return value;
            }
            return Text(value);
        }

        // Called by the pipeline's final save phase only. A failed write never truncates the old file.
        public static DataTable Write(DataTable plan)
        {
            if (plan == null || plan.Rows.Count != 1)
                throw new InvalidOperationException("Run the text / JSON preview before saving.");
            string path = Convert.ToString(plan.Rows[0][ExcelFile.FindColumn(plan, "Path")], CultureInfo.InvariantCulture);
            string content = Convert.ToString(plan.Rows[0][ExcelFile.FindColumn(plan, "Content")], CultureInfo.InvariantCulture);
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temp, content, new UTF8Encoding(false, true));
                if (File.Exists(path))
                    File.Replace(temp, path, null);
                else
                    File.Move(temp, path);
            }
            finally
            {
                if (File.Exists(temp))
                    File.Delete(temp);
            }
            var output = plan.Copy();
            output.Rows[0][ExcelFile.FindColumn(output, "Status")] = "Saved";
            return output;
        }
    }
}
