using Microsoft.UI.Xaml.Data;
using System;

namespace DocTransform.Converters;

// Add the partial keyword here
public partial class InverseBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        return !(bool)value;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        return !(bool)value;
    }
}