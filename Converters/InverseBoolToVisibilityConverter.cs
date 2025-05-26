using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using System;

namespace DocTransform.Converters;

// Add the partial keyword to resolve the CSWinRT1028 message
public partial class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool boolValue)
        {
            return boolValue ? Visibility.Collapsed : Visibility.Visible;
        }

        return Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        if (value is Visibility visibility)
        {
            return visibility != Visibility.Visible;
        }

        return false;
    }
}