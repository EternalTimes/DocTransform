using Microsoft.UI.Xaml.Data;
using System;

namespace DocTransform.Converters;

public partial class StringNotEmptyToBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is string strValue) return !string.IsNullOrEmpty(strValue);

        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}