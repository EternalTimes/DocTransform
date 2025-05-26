using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocTransform.Services;

public class WordGenerationOptions
{
    public int ImageWidthEmu { get; set; } = 250 * 9525;
    public int ImageHeightEmu { get; set; } = 150 * 9525;
    public bool CenterImage { get; set; } = true;
    public bool LogReplacements { get; set; } = true;
    public bool InsertPageNumber { get; set; } = true;
    public bool InsertTableOfContents { get; set; } = true;
    public Dictionary<string, RunStyle> FieldStyles { get; set; } = new();
}

public class RunStyle
{
    public string? FontName { get; set; } = "微软雅黑";
    public int FontSize { get; set; } = 11;
    public bool Bold { get; set; } = false;
    public bool Italic { get; set; } = false;
    public string? Color { get; set; } = "000000";
}

public class LogMessage
{
    public string Field { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}

public class WordService
{
    public List<LogMessage> Logs { get; } = new();

    public void GenerateWordFromTemplate(
        string templatePath,
        string outputPath,
        Dictionary<string, string> placeholders,
        Dictionary<string, string>? imagePaths = null,
        List<Dictionary<string, string>>? tableData = null,
        WordGenerationOptions? options = null)
    {
        options ??= new WordGenerationOptions();
        Logs.Clear();

        using var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
        var doc = new XWPFDocument(fs);

        ReplaceInAll(doc.Paragraphs, placeholders, imagePaths, options, "Body");

        foreach (var table in doc.Tables)
        {
            if (tableData != null && table.GetRow(1)?.GetTableCells()?.Any(c => GetCellText(c).Contains("{Row.")) == true)
                InsertTableData(table, tableData, options);
            else
            {
                foreach (var row in table.Rows)
                    foreach (var cell in row.GetTableCells())
                        ReplaceInAll(cell.Paragraphs, placeholders, imagePaths, options, "Table");
            }
        }

        foreach (var header in doc.HeaderList)
            ReplaceInAll(header.Paragraphs, placeholders, imagePaths, options, "Header");

        foreach (var footer in doc.FooterList)
        {
            ReplaceInAll(footer.Paragraphs, placeholders, imagePaths, options, "Footer");
            if (options.InsertPageNumber)
                AddPageNumberField(footer);
        }

        if (options.InsertTableOfContents)
            InsertTableOfContents(doc);

        using var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
        doc.Write(outFs);
    }

    private void ReplaceInAll(IEnumerable<XWPFParagraph> paragraphs, Dictionary<string, string> placeholders,
                              Dictionary<string, string>? imagePaths, WordGenerationOptions options, string section)
    {
        foreach (var para in paragraphs)
        {
            for (int i = 0; i < para.Runs.Count; i++)
            {
                var run = para.Runs[i];
                var text = run.ToString();

                foreach (var kv in placeholders)
                {
                    var placeholder = $"{{{kv.Key}}}";
                    if (text.Contains(placeholder))
                    {
                        try
                        {
                            text = text.Replace(placeholder, kv.Value ?? "");
                            run.SetText(text, 0);
                            ApplyStyle(run, kv.Key, options);
                            if (options.LogReplacements)
                                Logs.Add(new LogMessage { Field = kv.Key, Location = section, Status = "Replaced" });
                        }
                        catch
                        {
                            Logs.Add(new LogMessage { Field = kv.Key, Location = section, Status = "Error" });
                        }
                    }
                }

                if (imagePaths != null)
                {
                    foreach (var kv in imagePaths)
                    {
                        var imgKey = $"{{{kv.Key}.img}}";
                        if (text.Contains(imgKey) && File.Exists(kv.Value))
                        {
                            para.RemoveRun(i);
                            AddImage(para, kv.Value, options);
                            if (options.LogReplacements)
                                Logs.Add(new LogMessage { Field = kv.Key + ".img", Location = section, Status = "Inserted" });
                            break;
                        }
                    }
                }
            }
        }
    }

    private void InsertTableData(XWPFTable table, List<Dictionary<string, string>> data, WordGenerationOptions options)
    {
        var templateRow = table.GetRow(1);
        for (int i = 0; i < data.Count; i++)
        {
            var rowData = data[i];
            var newRow = table.InsertNewTableRow(i + 2);
            foreach (var cell in templateRow.GetTableCells())
            {
                var newCell = newRow.CreateCell();
                foreach (var para in cell.Paragraphs)
                {
                    var newPara = newCell.AddParagraph();
                    var newRun = newPara.CreateRun();
                    var text = para.Text;
                    foreach (var kv in rowData)
                        text = text.Replace($"{{Row.{kv.Key}}}", kv.Value ?? "");
                    newRun.SetText(text);
                    ApplyStyle(newRun, "", options);
                }
            }
        }
        table.RemoveRow(1);
    }

    private void AddImage(XWPFParagraph para, string imagePath, WordGenerationOptions options)
    {
        var run = para.CreateRun();
        using var imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        var imageType = GetPictureType(imagePath);
        run.AddPicture(imgStream, imageType, Path.GetFileName(imagePath), options.ImageWidthEmu, options.ImageHeightEmu);
        if (options.CenterImage)
            para.Alignment = ParagraphAlignment.CENTER;
    }

    private void AddPageNumberField(XWPFHeaderFooter footer)
    {
        foreach (var para in footer.Paragraphs)
        {
            var run = para.CreateRun();
            run.SetText("页码: ");
            run.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.begin;
            run = para.CreateRun();
            run.GetCTR().AddNewInstrText().Value = "PAGE";
            run = para.CreateRun();
            run.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.end;
        }
    }

    private void InsertTableOfContents(XWPFDocument doc)
    {
        var para = doc.CreateParagraph();
        para.Style = "TOCHeading";
        para.Alignment = ParagraphAlignment.CENTER;
        var run = para.CreateRun();
        run.SetText("目录");

        para = doc.CreateParagraph();
        run = para.CreateRun();
        run.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.begin;
        run = para.CreateRun();
        run.GetCTR().AddNewInstrText().Value = @"TOC \o ""1-3"" \h \z \u";
        run = para.CreateRun();
        run.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.end;
    }

    private int GetPictureType(string path)
    {
        var ext = Path.GetExtension(path).ToLower();
        return ext switch
        {
            ".emf" => (int)PictureType.EMF,
            ".wmf" => (int)PictureType.WMF,
            ".pict" => (int)PictureType.PICT,
            ".jpeg" or ".jpg" => (int)PictureType.JPEG,
            ".png" => (int)PictureType.PNG,
            ".dib" => (int)PictureType.DIB,
            ".gif" => (int)PictureType.GIF,
            ".tiff" or ".tif" => (int)PictureType.TIFF,
            ".eps" => (int)PictureType.EPS,
            ".bmp" => (int)PictureType.BMP,
            ".wpg" => (int)PictureType.WPG,
            _ => (int)PictureType.PNG
        };
    }

    private void ApplyStyle(XWPFRun run, string fieldKey, WordGenerationOptions options)
    {
        var style = options.FieldStyles.TryGetValue(fieldKey, out var s) ? s : new RunStyle();
        run.FontSize = style.FontSize;
        run.FontFamily = style.FontName;
        run.SetColor(style.Color ?? "000000");
        run.IsBold = style.Bold;
        run.IsItalic = style.Italic;
    }

    private string GetCellText(XWPFTableCell cell) =>
        string.Join("", cell.Paragraphs.Select(p => p.ParagraphText));
}
