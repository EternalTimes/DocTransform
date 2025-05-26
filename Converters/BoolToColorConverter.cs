using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using System;
using Windows.UI;
using System.Reflection;

namespace DocTransform.Converters
{

    public partial class BoolToColorConverter : IValueConverter
    {
        private readonly Brush _trueBrushFallback = new SolidColorBrush(Color.FromArgb(255, 0xFF, 0x8C, 0x00));
        private readonly Brush _falseBrushFallback = new SolidColorBrush(Color.FromArgb(255, 0x00, 0x64, 0x00));

        public object Convert(object value, Type targetType, object parameter, string language)
        {
            bool isTrueCondition = false;
            if (value is bool bVal)
            {
                isTrueCondition = bVal;
            }

            string[]? colorParams = ((string)parameter)?.Split(';');

            if (colorParams == null || colorParams.Length != 2)
            {
                return isTrueCondition ? _trueBrushFallback : _falseBrushFallback;
            }

            string targetColorStr = isTrueCondition ? colorParams[0].Trim() : colorParams[1].Trim();

            if (Application.Current.Resources.TryGetValue(targetColorStr, out var brushResource))
            {
                if (brushResource is Brush brush)
                {
                    return brush;
                }
                if (brushResource is Color colorFromResource)
                {
                    return new SolidColorBrush(colorFromResource);
                }
            }

            try
            {

                if (!string.IsNullOrEmpty(targetColorStr) && targetColorStr[0] == '#')
                {
                    byte a = 255;
                    int startIndex = 1;

                    if (targetColorStr.Length == 9)
                    {
                        a = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                        startIndex += 2;
                    }
                    else if (targetColorStr.Length != 7 && targetColorStr.Length != 4)
                    {
                        return isTrueCondition ? _trueBrushFallback : _falseBrushFallback;
                    }

                    byte r_val, g_val, b_val;

                    if (targetColorStr.Length == 4)
                    {
                        char rChar = targetColorStr[startIndex++];
                        char gChar = targetColorStr[startIndex++];
                        char bChar = targetColorStr[startIndex];

                        r_val = System.Convert.ToByte(new string(rChar, 2), 16);
                        g_val = System.Convert.ToByte(new string(gChar, 2), 16);
                        b_val = System.Convert.ToByte(new string(bChar, 2), 16);
                    }
                    else
                    {
                        r_val = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                        startIndex += 2;
                        g_val = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                        startIndex += 2;
                        b_val = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                    }
                    return new SolidColorBrush(Color.FromArgb(a, r_val, g_val, b_val));
                }

                var colorProperty = typeof(Color).GetProperty(targetColorStr, BindingFlags.Static | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (colorProperty != null && colorProperty.GetValue(null) is Color namedColor)
                {
                    return new SolidColorBrush(namedColor);
                }
            }
 
            catch
            {
                // 解析失败，将使用下面的备用颜色
            }

            return isTrueCondition ? _trueBrushFallback : _falseBrushFallback;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
}