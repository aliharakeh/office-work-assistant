using System;
using System.Data;
using System.Globalization;
using System.IO;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.LoadImage
{
    public sealed class LoadImageSettings
    {
        public string Path { get; set; }

        public LoadImageSettings()
        {
            Path = "";
        }
    }

    // A source table for image-processing steps. The image itself stays on disk.
    public static class LoadImageWork
    {
        public static DataTable Run(LoadImageSettings settings)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.Path))
                throw new InvalidOperationException("Choose an image file.");
            string path = settings.Path.Trim();
            if (!Path.IsPathRooted(path) || Path.GetPathRoot(path).Length < 3)
                throw new InvalidOperationException("Use a full image path.");
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension != ".png" && extension != ".jpg" && extension != ".jpeg" &&
                extension != ".bmp" && extension != ".tif" && extension != ".tiff")
                throw new InvalidOperationException("Choose a PNG, JPEG, BMP or TIFF image.");
            if (!File.Exists(path))
                throw new InvalidOperationException("Image not found or not accessible: " + path);

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            ExcelFile.AddColumn(output, "Name");
            ExcelFile.AddColumn(output, "Path");
            output.Rows.Add(Path.GetFileName(path), Path.GetFullPath(path));
            return output;
        }
    }
}
