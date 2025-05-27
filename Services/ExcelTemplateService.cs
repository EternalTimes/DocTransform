using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocTransform.Models;
using NPOI.XSSF.UserModel.Helpers;

namespace DocTransform.Services;

public class ExcelTemplateService
{
    private readonly ImageProcessingService _imageProcessingService;

    public ExcelTemplateService(ImageProcessingService imageProcessingService)
    {
        _imageProcessingService = imageProcessingService;
    }

    public async Task<(bool Success, string Message)> ProcessTemplateWithImagesAsync(
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

                                var originalStyle = cell.CellStyle;

                                foreach (var item in data)
                                {
                                    var placeholder = $"{{{item.Key}}}";
                                    newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                }

                                var match = Regex.Match(newValue, @"^{(.*?)\.img}$");
                                if (match.Success)
                                {
                                    var key = match.Groups[1].Value;
                                    if (imagePaths.TryGetValue(key, out string? imagePath) && File.Exists(imagePath))
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

    private void InsertImage(
        IWorkbook workbook,
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
            // 合并 2x2 区域作为插图目标区域
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
            Dx1 = 100, // 精细控制图像偏移量（单位EMU）
            Dy1 = 50,
            AnchorType = AnchorType.MoveAndResize
        };

        var picture = drawing.CreatePicture(anchor, pictureIdx);

        switch (mode)
        {
            case ImageFillMode.Stretch:
                picture.Resize(1.0); // 完全拉伸
                break;
            case ImageFillMode.Fit:
                picture.Resize(); // 保持比例适应单元格
                break;
            case ImageFillMode.Fill:
                picture.Resize(1.2); // 略微放大以填充
                break;
        }
    }
}
