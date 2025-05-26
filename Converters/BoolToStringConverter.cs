using Microsoft.UI.Xaml.Data;
using System;

namespace DocTransform.Converters;

// 添加 partial 关键字以解决 CSWinRT1028
public partial class BoolToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool boolValue && parameter is string options)
        {
            var parts = options.Split(';');
            if (parts.Length >= 2)
            {
                return boolValue ? parts[0] : parts[1];
            }
        }

        return value?.ToString() ?? string.Empty;
    }

    // 必须实现 ConvertBack 方法以满足 IValueConverter 接口
    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}