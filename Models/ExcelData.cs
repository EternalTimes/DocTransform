using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace DocTransform.Models;

// 1. 继承自 ObservableObject，以获得属性变更通知的能力
public partial class ExcelData : ObservableObject
{
    // 2. 将 List<T> 替换为 ObservableCollection<T>，以便UI可以监听到集合中项目的增删
    public ObservableCollection<string> Headers { get; set; } = new();
    public ObservableCollection<Dictionary<string, string>> Rows { get; set; } = new();
    public ObservableCollection<string> SelectedColumns { get; set; } = new();

    // 3. 使用 [ObservableProperty] 特性来自动生成可通知UI的属性
    //    这会自动创建一个名为 SourceFileName 的属性，和一个名为 _sourceFileName 的私有字段
    [ObservableProperty]
    private string _sourceFileName = string.Empty;

    // 4. RowCount 是一个只读属性，它的值依赖于Rows集合
    public int RowCount => Rows.Count;

    public ExcelData()
    {
        // 5. 监听Rows集合的变化事件。当Rows集合内容改变时（增/删），
        //    手动触发 RowCount 属性的变更通知，从而更新UI上显示的数量
        Rows.CollectionChanged += OnRowsCollectionChanged;
    }

    private void OnRowsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // 通知UI，RowCount属性的值已经发生了变化
        OnPropertyChanged(nameof(RowCount));
    }
}