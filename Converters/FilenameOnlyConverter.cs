using Microsoft.UI.Xaml.Data;
using System;
using System.IO; // ������� 1: ���� System.IO �����ռ���ʹ�� Path ��

namespace DocTransform.Converters;

// ������� 2: ������Ϊ partial �Լ��� C#/WinRT ������
public partial class FilenameOnlyConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        // ���� Path.GetFileName ���Ա���ȷʶ��
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