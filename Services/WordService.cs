using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.Util;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocTransform.Services;

public class WordService
{
    /// <summary>
    /// 使用 NPOI 替换 Word 模板中的文本与图像占位符，支持段落、表格、页眉页脚。
    /// </summary>
    public void ReplacePlaceholders(string templatePath, string outputPath, Dictionary<string, string> placeholders, Dictionary<string, string>? imagePaths = null)
    {
        using var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
        var doc = new XWPFDocument(fs);

        ReplaceInParagraphs(doc.Paragraphs, placeholders, imagePaths);

        foreach (var table in doc.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    ReplaceInParagraphs(cell.Paragraphs, placeholders, imagePaths);
                }
            }
        }

        // 页眉页脚替换
        foreach (var header in doc.HeaderList)
            ReplaceInParagraphs(header.Paragraphs, placeholders, imagePaths);

        foreach (var footer in doc.FooterList)
            ReplaceInParagraphs(footer.Paragraphs, placeholders, imagePaths);

        using var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
        doc.Write(outFs);
    }

    private void ReplaceInParagraphs(IEnumerable<XWPFParagraph> paragraphs, Dictionary<string, string> placeholders, Dictionary<string, string>? imagePaths)
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
                        text = text.Replace(placeholder, kv.Value ?? "");
                        run.SetText(text, 0);
                        ApplyDefaultStyle(run);
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
                            AddImage(para, kv.Value);
                            break;
                        }
                    }
                }
            }
        }
    }

    private void AddImage(XWPFParagraph para, string imagePath)
    {
        var run = para.CreateRun();
        using var imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        var imageType = GetPictureType(imagePath);
        run.AddPicture(imgStream, imageType, Path.GetFileName(imagePath), Units.ToEMU(250), Units.ToEMU(150));
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

    private void ApplyDefaultStyle(XWPFRun run)
    {
        run.FontSize = 11;
        run.FontFamily = "微软雅黑";
        run.IsBold = false;
        run.SetColor("000000"); // 黑色
    }
}
