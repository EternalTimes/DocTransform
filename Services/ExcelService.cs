using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocTransform.Models;

namespace DocTransform.Services;

/// <summary>
///     基于 NPOI 的 Excel 文件处理服务
/// </summary>
public class ExcelService
{
    /// <summary>
    ///     异步读取 Excel 文件的第一个工作表
    /// </summary>
    public async Task<ExcelData> ReadExcelFileAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel文件不存在", filePath);

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(stream);
            var sheet = workbook.GetSheetAt(0);

            if (sheet == null)
                throw new InvalidOperationException("Excel文件不包含工作表");

            var excelData = ParseSheet(sheet);
            excelData.SourceFileName = Path.GetFileName(filePath);
            return excelData;
        });
    }

    /// <summary>
    ///     异步读取 Excel 文件中所有非空工作表
    /// </summary>
    public async Task<List<ExcelData>> ReadAllSheetsAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel文件不存在", filePath);

            var result = new List<ExcelData>();

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(stream);

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                if (sheet == null || sheet.LastRowNum < 1) continue;

                var excelData = ParseSheet(sheet);
                excelData.SourceFileName = $"{Path.GetFileName(filePath)} - {sheet.SheetName}";

                if (excelData.Headers.Count > 0 && excelData.Rows.Count > 0)
                    result.Add(excelData);
            }

            return result;
        });
    }

    /// <summary>
    ///     从工作表解析 ExcelData
    /// </summary>
    private ExcelData ParseSheet(ISheet sheet)
    {
        var data = new ExcelData();
        var headerRow = sheet.GetRow(0);

        if (headerRow == null) return data;

        for (int i = 0; i < headerRow.LastCellNum; i++)
        {
            var cell = headerRow.GetCell(i);
            var header = cell?.ToString()?.Trim();
            if (!string.IsNullOrEmpty(header))
                data.Headers.Add(header);
        }

        for (int r = 1; r <= sheet.LastRowNum; r++)
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;

            var rowData = new Dictionary<string, string>();
            bool hasData = false;

            for (int c = 0; c < data.Headers.Count; c++)
            {
                var cellValue = row.GetCell(c)?.ToString()?.Trim() ?? string.Empty;
                rowData[data.Headers[c]] = cellValue;
                if (!string.IsNullOrEmpty(cellValue)) hasData = true;
            }

            if (hasData)
                data.Rows.Add(rowData);
        }

        return data;
    }
}