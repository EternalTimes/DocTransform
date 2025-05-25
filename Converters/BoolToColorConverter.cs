using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using System;
using Windows.UI; // Essential for Windows.UI.Colors and Color struct
using System.Reflection; // For Type.GetProperty for named colors

namespace DocTransform.Converters
{
    public class BoolToColorConverter : IValueConverter
    {
        // 定义备用颜色，如果参数解析失败或资源找不到时使用
        // DarkOrange (Hex: #FF8C00)
        private readonly Brush _trueBrushFallback = new SolidColorBrush(Color.FromArgb(255, 0xFF, 0x8C, 0x00));
        // DarkGreen (Hex: #006400)
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
                // 参数格式不正确，返回基于布尔值的备用颜色
                return isTrueCondition ? _trueBrushFallback : _falseBrushFallback;
            }

            string targetColorStr = isTrueCondition ? colorParams[0].Trim() : colorParams[1].Trim();

            // 1. 尝试从应用资源中查找画刷 (如果参数是资源键)
            if (Application.Current.Resources.TryGetValue(targetColorStr, out var brushResource))
            {
                if (brushResource is Brush brush)
                {
                    return brush;
                }
                // 如果资源本身是 Color 定义 (虽然不常见于 ThemeResource 的 Brush Key, 但为了兼容自定义资源)
                if (brushResource is Color colorFromResource)
                {
                    return new SolidColorBrush(colorFromResource);
                }
            }

            // 2. 如果不是资源键，尝试将其解析为颜色字符串 (十六进制或命名颜色)
            try
            {
                // 尝试解析十六进制颜色 (例如 #AARRGGBB, #RRGGBB, #RGB)
                if (targetColorStr.StartsWith("#"))
                {
                    byte a = 255;
                    int startIndex = 1;

                    if (targetColorStr.Length == 9) // #AARRGGBB
                    {
                        a = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                        startIndex += 2;
                    }
                    // 只接受 #RRGGBB (7 chars) 或 #RGB (4 chars) 或 #AARRGGBB (9 chars)
                    else if (targetColorStr.Length != 7 && targetColorStr.Length != 4)
                    {
                        // System.Diagnostics.Debug.WriteLine($"BoolToColorConverter: Hex color '{targetColorStr}' format incorrect. Using fallback.");
                        return isTrueCondition ? _trueBrushFallback : _falseBrushFallback;
                    }

                    byte r_val, g_val, b_val;

                    if (targetColorStr.Length == 4) // #RGB -> #RRGGBB
                    {
                        char rChar = targetColorStr[startIndex++];
                        char gChar = targetColorStr[startIndex++];
                        char bChar = targetColorStr[startIndex];
                        r_val = System.Convert.ToByte(new string(new char[] { rChar, rChar }), 16);
                        g_val = System.Convert.ToByte(new string(new char[] { gChar, gChar }), 16);
                        b_val = System.Convert.ToByte(new string(new char[] { bChar, bChar }), 16);
                    }
                    else // #RRGGBB (from #AARRGGBB or #RRGGBB)
                    {
                        r_val = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                        startIndex += 2;
                        g_val = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                        startIndex += 2;
                        b_val = System.Convert.ToByte(targetColorStr.Substring(startIndex, 2), 16);
                    }
                    return new SolidColorBrush(Color.FromArgb(a, r_val, g_val, b_val));
                }

                // 尝试解析 Windows.UI.Colors 中的命名颜色 (例如 "Orange", "Green")
                var colorProperty = typeof(Windows.UI.Color).GetProperty(targetColorStr, BindingFlags.Static | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (colorProperty != null && colorProperty.GetValue(null) is Color namedColor)
                {
                    return new SolidColorBrush(namedColor);
                }
            }
            catch (Exception ex)
            {
                // System.Diagnostics.Debug.WriteLine($"BoolToColorConverter: Error parsing color string '{targetColorStr}': {ex.Message}. Using fallback.");
                // 解析失败，将使用下面的备用颜色
            }

            // System.Diagnostics.Debug.WriteLine($"BoolToColorConverter: Could not resolve or parse '{targetColorStr}' as resource, hex, or named color. Using fallback.");
            return isTrueCondition ? _trueBrushFallback : _falseBrushFallback;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
}