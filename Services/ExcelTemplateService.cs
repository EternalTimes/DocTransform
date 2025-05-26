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

    // 其他方法（如模板替换）建议按需重写，并兼容 NPOI 插图逻辑



}
