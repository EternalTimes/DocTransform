using Microsoft.UI.Xaml.Data;

using System;

namespace DocTransform.Converters;

public class BoolToIconConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        var boolValue = (bool)value;
        var icons = ((string)parameter).Split(';');
        return boolValue ? icons[0] : icons[1];
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}