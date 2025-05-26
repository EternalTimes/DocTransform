using Microsoft.UI.Xaml.Data;
using System;
using System.IO; // 解决方案 1: 引入 System.IO 命名空间以使用 Path 类

namespace DocTransform.Converters;

// 解决方案 2: 将类标记为 partial 以兼容 C#/WinRT 工具链
public partial class FilenameOnlyConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        // 现在 Path.GetFileName 可以被正确识别
        if (value is string path && !string.IsNullOrEmpty(path))
        {
            return Path.GetFileName(path);
        }
        return string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}