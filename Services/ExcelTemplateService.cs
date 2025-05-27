using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocTransform.Models;
using System.Collections.ObjectModel;
using System.Linq;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace DocTransform.Services;

public class ExcelTemplateService
{
    private readonly ImageProcessingService _imageProcessingService;

    public ExcelTemplateService(ImageProcessingService imageProcessingService)
    {
        _imageProcessingService = imageProcessingService;
    }

    public static async Task<bool> IsValidTemplateAsync(string templatePath)
    {
        return await Task.Run(() =>
            File.Exists(templatePath) && Path.GetExtension(templatePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase));
    }

    public static async Task<List<string>> ExtractPlaceholdersAsync(string templatePath)
    {
        return await Task.Run(() =>
        {
            var placeholders = new HashSet<string>();
            using var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(fs);
            for (int s = 0; s < workbook.NumberOfSheets; s++)
            {
                var sheet = workbook.GetSheetAt(s);
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.CellType == CellType.String)
                        {
                            var matches = Regex.Matches(cell.StringCellValue, "\\{(.*?)\\}");
                            foreach (Match match in matches)
                                placeholders.Add(match.Groups[1].Value);
                        }
                    }
                }
            }
            return placeholders.ToList();
        });
    }

    public static async Task<(bool Success, string Message)> ProcessTemplateAsync(
        string templatePath, string outputPath, Dictionary<string, string> data, object? _ = null)
    {
        return await ProcessTemplateWithImagesAsync(templatePath, outputPath, data, new Dictionary<string, string>());
    }

    public static async Task<(bool Success, string Message)> ProcessTemplateWithImagesAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        ObservableCollection<ImageSourceDirectory> imageDirectories,
        ImageFillMode _1,
        int _2,
        object? _3 = null)
    {
        var imagePaths = imageDirectories
            .SelectMany(d => d.ImageFiles.Select(img => new { Key = Path.GetFileNameWithoutExtension(img), Path = img }))
            .ToDictionary(k => k.Key, v => v.Path);

        return await ProcessTemplateWithImagesAsync(templatePath, outputPath, data, imagePaths);
    }

    public static async Task<(bool Success, string Message)> ProcessTemplateWithImagesAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        Dictionary<string, string> imagePaths,
        IProgress<int>? progress = null)
    {
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath)) return (false, "模板文件不存在");
        if (string.IsNullOrEmpty(outputPath)) return (false, "输出路径无效");

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        return await Task.Run(() =>
        {
            try
            {
                File.Copy(templatePath, outputPath, true);

                using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
                var workbook = new XSSFWorkbook(fs);

                int totalSheets = workbook.NumberOfSheets;
                int processedSheets = 0;

                for (int s = 0; s < totalSheets; s++)
                {
                    var sheet = workbook.GetSheetAt(s);

                    for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                    {
                        var row = sheet.GetRow(r);
                        if (row == null) continue;

                        foreach (var cell in row.Cells)
                        {
                            if (cell.CellType == CellType.String)
                            {
                                string value = cell.StringCellValue;
                                string newValue = value;

                                var originalStyle = cell.CellStyle;

                                foreach (var item in data)
                                {
                                    var placeholder = $"{{{item.Key}}}";
                                    newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                }

                                var match = Regex.Match(newValue, "^{(.*?)\\.img}$");
                                if (match.Success)
                                {
                                    var key = match.Groups[1].Value;
                                    if (imagePaths.TryGetValue(key, out var imagePath) && imagePath is not null && File.Exists(imagePath))
                                    {
                                        InsertImage(workbook, sheet, imagePath, r, cell.ColumnIndex, ImageFillMode.Fit, merge: true);
                                        newValue = string.Empty;
                                    }
                                }

                                if (newValue != value)
                                {
                                    cell.SetCellValue(newValue);
                                    cell.CellStyle = originalStyle;
                                }
                            }
                        }
                    }

                    processedSheets++;
                    progress?.Report(processedSheets * 100 / totalSheets);
                }

                using var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFs);
                return (true, "模板处理完成");
            }
            catch (Exception ex)
            {
                return (false, $"处理模板出错: {ex.Message}");
            }
        });
    }

    private static void InsertImage(
        XSSFWorkbook workbook,
        ISheet sheet,
        string imagePath,
        int rowIndex,
        int colIndex,
        ImageFillMode mode,
        bool merge = false)
    {
        var bytes = File.ReadAllBytes(imagePath);
        int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

        if (merge)
        {
            int row2 = rowIndex + 1;
            int col2 = colIndex + 1;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, row2, colIndex, col2));
        }

        var drawing = sheet.CreateDrawingPatriarch();
        var anchor = new XSSFClientAnchor
        {
            Col1 = colIndex,
            Row1 = rowIndex,
            Col2 = colIndex + 1,
            Row2 = rowIndex + 1,
            Dx1 = 100,
            Dy1 = 50,
            AnchorType = AnchorType.MoveAndResize
        };

        var picture = drawing.CreatePicture(anchor, pictureIdx);

        switch (mode)
        {
            case ImageFillMode.Stretch:
                picture.Resize(1.0);
                break;
            case ImageFillMode.Fit:
                picture.Resize();
                break;
            case ImageFillMode.Fill:
                picture.Resize(1.2);
                break;
        }
    }
}
