using ImageMagick;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DocTransform.Services;

public class ImageProcessingService
{
    private static readonly string[] SupportedExtensions =
    [
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff",
        ".svg", ".ico", ".heic", ".heif", ".psd", ".ai"
    ];

    /// <summary>
    /// 获取目录下支持的图像文件路径列表
    /// </summary>
    public static async Task<List<string>> GetImageFilesAsync(string directory)
    {
        return await Task.Run(() =>
        {
            if (!Directory.Exists(directory))
                return new List<string>();

            return Directory.EnumerateFiles(directory, "*.*", SearchOption.TopDirectoryOnly)
                            .Where(file => SupportedExtensions.Contains(Path.GetExtension(file).ToLower()))
                            .ToList();
        });
    }

    /// <summary>
    /// 获取图像的尺寸（宽，高）
    /// </summary>
    public static async Task<(int Width, int Height)?> GetImageSizeAsync(string path)
    {
        return await Task.Run<(int Width, int Height)?>(() =>
        {
            try
            {
                using var image = new MagickImage(path);
                return ((int)image.Width, (int)image.Height);
            }
            catch
            {
                return null;
            }
        });
    }

    /// <summary>
    /// 批量加载图像信息（路径，宽，高）
    /// </summary>
    public static async Task<List<(string Path, int Width, int Height)>> LoadImageInfoAsync(IEnumerable<string> paths)
    {
        var result = new List<(string Path, int Width, int Height)>();

        foreach (var path in paths)
        {
            var size = await GetImageSizeAsync(path);
            if (size.HasValue)
                result.Add((path, size.Value.Width, size.Value.Height));
        }

        return result;
    }

    /// <summary>
    /// 将图像转换为 PNG 格式并保存到目标路径
    /// </summary>
    public static async Task<bool> ConvertToPngAsync(string inputPath, string outputPath)
    {
        return await Task.Run(() =>
        {
            try
            {
                using var image = new MagickImage(inputPath);
                image.Format = MagickFormat.Png;
                image.Write(outputPath);
                return true;
            }
            catch
            {
                return false;
            }
        });
    }

    public static async Task<List<string>> ScanDirectoryForImagesAsync(string directoryPath)
    {
        return await GetImageFilesAsync(directoryPath);
    }
}
