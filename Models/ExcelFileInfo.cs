using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;

namespace DocTransform.Models;

public partial class ExcelFileInfo : ObservableObject
{
    private string _fileName = string.Empty;
    public string FileName
    {
        get => _fileName;
        set => SetProperty(ref _fileName, value);
    }

    private string _filePath = string.Empty;
    public string FilePath
    {
        get => _filePath;
        set => SetProperty(ref _filePath, value);
    }

    private int _rowCount;
    public int RowCount
    {
        get => _rowCount;
        set => SetProperty(ref _rowCount, value);
    }

    public ExcelFileInfo(string path)
    {
        FilePath = path;
        FileName = Path.GetFileName(path);
    }
}