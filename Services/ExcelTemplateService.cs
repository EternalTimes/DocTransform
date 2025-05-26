using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using DocTransform.Models;

namespace DocTransform.Services;

/// <summary>
///     Excel模板处理服务（NPOI重构）
/// </summary>

public enum ImageFillMode
{
    Fit,
    Fill,
    Stretch
}

public class ExcelTemplateService
{
    private readonly ImageProcessingService _imageProcessingService;

    public ExcelTemplateService(ImageProcessingService imageProcessingService)
    {
        _imageProcessingService = imageProcessingService;
    }

    public async Task<bool> IsValidTemplateAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return false;

        return await Task.Run(() =>
        {
            try
            {
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fs);
                return workbook.NumberOfSheets > 0;
            }
            catch
            {
                return false;
            }
        });
    }

    public async Task<List<string>> ExtractPlaceholdersAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return new List<string>();

        return await Task.Run(() =>
        {
            var placeholders = new HashSet<string>();
            var regex = new Regex(@"{([^{}]+)}");

            try
            {
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fs);

                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    var sheet = workbook.GetSheetAt(s);
                    for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                    {
                        var row = sheet.GetRow(r);
                        if (row == null) continue;

                        foreach (var cell in row.Cells)
                        {
                            var text = cell?.ToString();
                            if (!string.IsNullOrEmpty(text))
                            {
                                var matches = regex.Matches(text);
                                foreach (Match match in matches)
                                    placeholders.Add(match.Value);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"提取占位符失败: {ex.Message}");
            }

            return new List<string>(placeholders);
        });
    }

        public async Task<(bool Success, string Message)> ProcessTemplateAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        IProgress<int> progress = null)
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
                IWorkbook workbook = new XSSFWorkbook(fs);

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

                                foreach (var item in data)
                                {
                                    var placeholder = $"{{{item.Key}}}";
                                    if (newValue.Contains(placeholder))
                                    {
                                        newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                    }
                                }

                                if (newValue != value)
                                    cell.SetCellValue(newValue);
                            }
                        }
                    }

                    processedSheets++;
                    progress?.Report(processedSheets * 100 / totalSheets);
                }

                using var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFs);
                return (true, "Excel模板处理成功");
            }
            catch (Exception ex)
            {
                return (false, $"处理Excel模板时出错: {ex.Message}");
            }
        });
    }


    public async Task<(bool Success, string Message)> ProcessTemplateWithImagesAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        Dictionary<string, string> imagePaths,
        IProgress<int> progress = null)
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
                IWorkbook workbook = new XSSFWorkbook(fs);

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

                                // 普通文本占位符替换
                                foreach (var item in data)
                                {
                                    var placeholder = $"{{{item.Key}}}";
                                    if (newValue.Contains(placeholder))
                                        newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                }

                                // 图像占位符处理
                                var match = Regex.Match(newValue, @"^{(.*?)\.img}$");
                                if (match.Success)
                                {
                                    var key = match.Groups[1].Value;
                                    if (imagePaths.TryGetValue(key, out string imagePath) && File.Exists(imagePath))
                                    {
                                        InsertImage(workbook, sheet, imagePath, r, cell.ColumnIndex);
                                        newValue = string.Empty;
                                    }
                                }

                                if (newValue != value)
                                    cell.SetCellValue(newValue);
                            }
                        }
                    }

                    processedSheets++;
                    progress?.Report(processedSheets * 100 / totalSheets);
                }

                using var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFs);
                return (true, "Excel模板图像处理成功");
            }
            catch (Exception ex)
            {
                return (false, $"处理模板图像时出错: {ex.Message}");
            }
        });
    }

    public void InsertImage(
        IWorkbook workbook,
        ISheet sheet,
        string imagePath,
        int rowIndex,
        int colIndex,
        ImageFillMode mode)
    {
        var bytes = File.ReadAllBytes(imagePath);
        int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

        var drawing = sheet.CreateDrawingPatriarch();
        var anchor = new XSSFClientAnchor
        {
            Col1 = colIndex,
            Row1 = rowIndex,
            Col2 = colIndex + 1,
            Row2 = rowIndex + 1,
            AnchorType = AnchorType.MoveAndResize
        };

        var picture = drawing.CreatePicture(anchor, pictureIdx);

        switch (mode)
        {
            case ImageFillMode.Stretch:
                picture.Resize(1.0); // 拉伸：完全填满
                break;
            case ImageFillMode.Fit:
                picture.Resize(); // 适应：等比例缩放
                break;
            case ImageFillMode.Fill:
                // Fill：拉伸后居中，可扩展进一步控制
                picture.Resize(1.2); // 稍微放大填充
                break;
        }
    }

}
