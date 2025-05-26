using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace DocTransform.Models;

public class ImageSourceDirectory : ObservableObject
{
    private string _directoryPath = string.Empty;
    public string DirectoryPath
    {
        get => _directoryPath;
        set => SetProperty(ref _directoryPath, value);
    }

    private string _directoryName = string.Empty;
    public string DirectoryName
    {
        get => _directoryName;
        set
        {
            if (SetProperty(ref _directoryName, value))
            {
                OnPropertyChanged(nameof(PlaceholderName));
            }
        }
    }

    private string _matchingColumn = string.Empty;
    public string MatchingColumn
    {
        get => _matchingColumn;
        set => SetProperty(ref _matchingColumn, value);
    }

    private ObservableCollection<string> _imageFiles = new();
    public ObservableCollection<string> ImageFiles
    {
        get => _imageFiles;
        set
        {
            if (_imageFiles != null)
                _imageFiles.CollectionChanged -= ImageFiles_CollectionChanged;

            if (SetProperty(ref _imageFiles, value))
            {
                _imageFiles.CollectionChanged += ImageFiles_CollectionChanged;
                OnPropertyChanged(nameof(ImageCount));
            }
        }
    }

    public string PlaceholderName => $"{{{DirectoryName}}}";
    public int ImageCount => ImageFiles.Count;

    public ImageSourceDirectory()
    {
        _imageFiles.CollectionChanged += ImageFiles_CollectionChanged;
    }

    private void ImageFiles_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(ImageCount));
    }
}
