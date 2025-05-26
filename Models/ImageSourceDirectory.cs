using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace DocTransform.Models; // 确保命名空间与您项目一致

public partial class ImageSourceDirectory : ObservableObject
{
    [ObservableProperty]
    private string _directoryPath = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PlaceholderName))]
    private string _directoryName = string.Empty;

    [ObservableProperty]
    private string _matchingColumn = string.Empty;

    [ObservableProperty]
    private ObservableCollection<string> _imageFiles = new();

    /// <summary>
    ///     占位符名称，格式为 {目录名}。
    ///     修正：引用私有字段 _directoryName
    /// </summary>
    public string PlaceholderName => $"{{{_directoryName}}}";

    /// <summary>
    ///     图片文件数量。
    ///     修正：引用私有字段 _imageFiles
    /// </summary>
    public int ImageCount => _imageFiles.Count;

    public ImageSourceDirectory()
    {
        // 修正：在私有字段 _imageFiles 上监听事件
        _imageFiles.CollectionChanged += (s, e) => OnPropertyChanged(nameof(ImageCount));
    }
}