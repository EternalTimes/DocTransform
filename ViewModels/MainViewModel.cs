using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocTransform.Constants;
using DocTransform.Models;
using DocTransform.Services;
using Microsoft.UI.Dispatching;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
// using DocumentFormat.OpenXml.Packaging; // 如果 CheckPlaceholders 方法不再使用 OpenXML SDK，可以移除

namespace DocTransform.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ExcelService _excelService;
    private readonly WordService _wordService;
    private readonly IdCardService _idCardService;
    private readonly ImageProcessingService _imageProcessingService;
    private readonly ExcelTemplateService _excelTemplateService;
    private readonly DispatcherQueue _dispatcherQueue;

    private ExcelData _excelData = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ExcelFilePathDisplayText))]
    [NotifyPropertyChangedFor(nameof(ExcelFilePathIsPlaceholder))]
    [NotifyPropertyChangedFor(nameof(SingleModeExcelDisplayText))]
    [NotifyPropertyChangedFor(nameof(SingleModeExcelIsPlaceholder))]
    private string _excelFilePath = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CurrentTableModeText))]
    [NotifyPropertyChangedFor(nameof(ExcelButtonText))]
    [NotifyPropertyChangedFor(nameof(ExcelFilePathDisplayText))]
    [NotifyPropertyChangedFor(nameof(ExcelFilePathIsPlaceholder))]
    [NotifyPropertyChangedFor(nameof(MergedDataSummary))]
    private bool _isMultiTableMode;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsOutputFolderAvailable))]
    private string _outputDirectory = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(MergedDataSummary))]
    private string _selectedKeyColumn = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ExcelFilePathDisplayText))]
    [NotifyPropertyChangedFor(nameof(ExcelFilePathIsPlaceholder))]
    [NotifyPropertyChangedFor(nameof(MultiModeExcelSummaryText))]
    [NotifyPropertyChangedFor(nameof(MultiModeExcelIsPlaceholder))]
    [NotifyPropertyChangedFor(nameof(MergedDataSummary))]
    private MultiTableData _multiTableData;

    [ObservableProperty] private ObservableCollection<string> _availableColumns = new();
    [ObservableProperty] private ObservableCollection<string> _availableIdCardColumns = new();
    [ObservableProperty] private ObservableCollection<string> _availableKeyColumns = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SingleModeExcelDisplayText))]
    [NotifyPropertyChangedFor(nameof(SingleModeExcelIsPlaceholder))]
    private ExcelData _currentExcelData = new();

    [ObservableProperty] private ObservableCollection<string> _detectedExcelPlaceholders = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IdCardExtractionRelatedPropertiesMightChange))]
    private bool _enableIdCardExtraction;

    [ObservableProperty] private string _excelTemplatePath = string.Empty;
    [ObservableProperty] private List<string> _idCardPlaceholders = PlaceholderConstants.AllPlaceholders;
    [ObservableProperty] private ObservableCollection<ImageSourceDirectory> _imageDirectories = new();
    [ObservableProperty] private ImageFillMode _imageFillMode = ImageFillMode.Fit;
    [ObservableProperty] private List<ImageFillModeItem> _imageFillModeItems = ImageFillModeItem.GetAll();
    [ObservableProperty] private int _imageFillPercentage = 90;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsProcessingOverlayVisible))]
    private bool _isProcessing;

    public bool IsProcessingOverlayVisible => IsProcessing;

    [ObservableProperty] private string _outputFileNameTemplate = "{序号}_{姓名}_{时间}";
    [ObservableProperty] private int _processedItems;
    [ObservableProperty] private string _processResultText = string.Empty;
    [ObservableProperty] private bool _processSuccess;
    [ObservableProperty] private int _progressValue;
    [ObservableProperty] private ObservableCollection<string> _selectedColumns = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IdCardExtractionRelatedPropertiesMightChange))]
    private string _selectedIdCardColumn = string.Empty;

    [ObservableProperty] private ImageFillModeItem _selectedImageFillModeItem;
    [ObservableProperty] private string _selectedImageMatchingColumn = string.Empty;
    [ObservableProperty] private string _statusMessage = "准备就绪";
    [ObservableProperty] private int _totalItems;
    [ObservableProperty] private bool _useExcelTemplate;
    [ObservableProperty] private bool _useImageReplacement;
    [ObservableProperty] private string _wordTemplatePath = string.Empty;

    public string CurrentTableModeText => IsMultiTableMode ? "多表格" : "单表格";
    public string ExcelButtonText => IsMultiTableMode ? "添加Excel" : "浏览";
    public string ExcelFilePathDisplayText => IsMultiTableMode ? MultiModeExcelSummaryText : SingleModeExcelDisplayText;
    public bool ExcelFilePathIsPlaceholder => IsMultiTableMode ? MultiModeExcelIsPlaceholder : SingleModeExcelIsPlaceholder;
    public string SingleModeExcelDisplayText => string.IsNullOrEmpty(ExcelFilePath) ? "未选择Excel文件" : Path.GetFileName(ExcelFilePath);
    public bool SingleModeExcelIsPlaceholder => string.IsNullOrEmpty(ExcelFilePath);
    public string MultiModeExcelSummaryText => (_multiTableData?.Tables?.Count ?? 0) == 0 ? "未添加Excel表格" : $"已添加 {MultiTableData.Tables.Count} 个表格";
    public bool MultiModeExcelIsPlaceholder => (_multiTableData?.Tables?.Count ?? 0) == 0;
    public bool IsOutputFolderAvailable => !string.IsNullOrEmpty(OutputDirectory) && Directory.Exists(OutputDirectory);
    public string MergedDataSummary =>
        (IsMultiTableMode && !string.IsNullOrEmpty(SelectedKeyColumn) && (_multiTableData?.MergedRows?.Count ?? 0) > 0)
            ? $"已使用 “{SelectedKeyColumn}” 列合并数据，共有 {MultiTableData.MergedRows.Count} 条记录"
            : (IsMultiTableMode && !string.IsNullOrEmpty(SelectedKeyColumn) ? "合并后无匹配数据或未选择有效的键列" : string.Empty);
    public bool IdCardExtractionRelatedPropertiesMightChange => true;

    public MainViewModel(DispatcherQueue dispatcherQueue)
    {
        _dispatcherQueue = dispatcherQueue;
        _excelService = new ExcelService();
        _wordService = new WordService();
        _idCardService = new IdCardService();
        _imageProcessingService = new ImageProcessingService();
        _excelTemplateService = new ExcelTemplateService(_imageProcessingService);

        OutputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        _selectedImageFillModeItem = ImageFillModeItems.First(item => item.Value == ImageFillMode.Fit);

        _multiTableData = new MultiTableData();
        SubscribeToMultiTableDataChanges(_multiTableData);
    }

    partial void OnMultiTableDataChanged(MultiTableData? oldValue, MultiTableData newValue)
    {
        if (oldValue != null) UnsubscribeFromMultiTableDataChanges(oldValue);
        if (newValue != null) SubscribeToMultiTableDataChanges(newValue);
        _dispatcherQueue.TryEnqueue(NotifyMultiTableDependentProperties);
    }

    private void SubscribeToMultiTableDataChanges(MultiTableData data)
    {
        if (data != null) data.Tables.CollectionChanged += OnMultiTableDataCollectionsChanged;
    }

    private void UnsubscribeFromMultiTableDataChanges(MultiTableData data)
    {
        if (data != null) data.Tables.CollectionChanged -= OnMultiTableDataCollectionsChanged;
    }

    private void OnMultiTableDataCollectionsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        _dispatcherQueue.TryEnqueue(NotifyMultiTableDependentProperties);
    }

    private void NotifyMultiTableDependentProperties()
    {
        OnPropertyChanged(nameof(ExcelFilePathDisplayText));
        OnPropertyChanged(nameof(ExcelFilePathIsPlaceholder));
        OnPropertyChanged(nameof(MultiModeExcelSummaryText));
        OnPropertyChanged(nameof(MultiModeExcelIsPlaceholder));
        OnPropertyChanged(nameof(MergedDataSummary));
        UpdateKeyColumns(); UpdateAvailableColumns(); UpdateIdCardColumns();
    }

    [RelayCommand] private void BrowseExcelFile() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择Excel文件...");
    [RelayCommand] private void BrowseWordTemplate() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择Word模板...");
    [RelayCommand] private void BrowseOutputDirectory() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择输出文件夹...");
    [RelayCommand] private void BrowseExcelTemplate() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择Excel模板...");
    [RelayCommand] private void AddImageDirectoryCmd() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择图片目录...");

    async partial void OnExcelFilePathChanged(string? oldValue, string? newValue)
    {
        if (!string.IsNullOrEmpty(newValue) && File.Exists(newValue))
        {
            if (IsMultiTableMode) await AddExcelFileToMultiTable(newValue);
            else await LoadSingleExcelFile(newValue);
        }
        else if (string.IsNullOrEmpty(newValue) && !IsMultiTableMode)
        {
            _dispatcherQueue.TryEnqueue(() => {
                CurrentExcelData = new ExcelData(); _excelData = CurrentExcelData;
                UpdateAvailableColumns(); UpdateIdCardColumns();
            });
        }
    }

    async partial void OnWordTemplatePathChanged(string? oldValue, string? newValue)
    {
        if (!string.IsNullOrEmpty(newValue) && File.Exists(newValue))
        {
            var isValid = await _wordService.IsValidTemplateAsync(newValue);
            _dispatcherQueue.TryEnqueue(() => {
                if (isValid) { CheckPlaceholdersCommand?.Execute(null); }
                else { StatusMessage = "选择的Word模板无效"; /* WordTemplatePath = string.Empty; */ }
            });
        }
        else { _dispatcherQueue.TryEnqueue(() => { /* Handle path cleared if needed */ }); }
    }

    async partial void OnExcelTemplatePathChanged(string? oldValue, string? newValue)
    {
        if (!string.IsNullOrEmpty(newValue) && File.Exists(newValue))
        {
            var isValid = await _excelTemplateService.IsValidTemplateAsync(newValue);
            _dispatcherQueue.TryEnqueue(() => {
                if (!isValid) { StatusMessage = "选择的Excel模板无效"; /* ExcelTemplatePath = string.Empty; */ }
                else { CheckExcelPlaceholdersCommand?.Execute(null); }
            });
        }
        else { _dispatcherQueue.TryEnqueue(() => { /* Handle path cleared */ }); }
    }

    public async Task ProcessAddedImageDirectory(string directoryPath)
    {
        if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath)) return;
        try
        {
            _dispatcherQueue.TryEnqueue(() => IsProcessing = true);
            var directoryName = Path.GetFileName(directoryPath);
            if (string.IsNullOrEmpty(directoryName)) directoryName = new DirectoryInfo(directoryPath).Name;
            // ... (目录名去重逻辑) ...
            _dispatcherQueue.TryEnqueue(() => StatusMessage = "正在扫描图片目录...");
            var imageFiles = await _imageProcessingService.ScanDirectoryForImagesAsync(directoryPath);

            _dispatcherQueue.TryEnqueue(() => {
                if (imageFiles.Count == 0) { StatusMessage = $"目录 {directoryName} 中未找到支持的图片文件"; IsProcessing = false; return; }
                var imageDirectory = new ImageSourceDirectory
                {
                    DirectoryPath = directoryPath,
                    DirectoryName = directoryName,
                    MatchingColumn = SelectedImageMatchingColumn,
                    ImageFiles = new ObservableCollection<string>(imageFiles)
                };
                ImageDirectories.Add(imageDirectory); UseImageReplacement = true;
                StatusMessage = $"已添加图片目录: {directoryName}，包含 {imageFiles.Count} 个图片";
            });
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"添加图片目录失败: {ex.Message}"); }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }

    public async Task LoadSingleExcelFile(string filePath)
    {
        try
        {
            _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = "正在加载Excel文件..."; });
            ExcelData loadedData = await _excelService.ReadExcelFileAsync(filePath);
            _dispatcherQueue.TryEnqueue(() => {
                CurrentExcelData = loadedData; _excelData = loadedData;
                UpdateAvailableColumns(); UpdateIdCardColumns();
                StatusMessage = $"Excel文件加载完成，共有 {loadedData.Rows.Count} 行数据";
            });
        }
        catch (Exception ex)
        {
            _dispatcherQueue.TryEnqueue(() => { StatusMessage = $"加载Excel文件失败: {ex.Message}"; ExcelFilePath = string.Empty; });
        }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }

    public async Task AddExcelFileToMultiTable(string filePath)
    {
        try
        {
            _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = "正在加载Excel文件..."; });
            var allSheets = await _excelService.ReadAllSheetsAsync(filePath);
            _dispatcherQueue.TryEnqueue(() => {
                if (allSheets.Count == 0) { StatusMessage = "Excel文件中没有有效数据"; IsProcessing = false; return; }
                foreach (var sheet in allSheets) _multiTableData.Tables.Add(sheet);
                NotifyMultiTableDependentProperties();
                StatusMessage = $"添加了 {allSheets.Count} 个工作表，共有 {_multiTableData.TotalRowCount} 行数据";
            });
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"加载Excel文件失败: {ex.Message}"); }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }

    private void UpdateAvailableColumns()
    {
        _dispatcherQueue.TryEnqueue(() => {
            AvailableColumns.Clear(); SelectedColumns.Clear();
            var headersToAdd = new HashSet<string>();
            if (IsMultiTableMode) foreach (var header in _multiTableData.AllHeaders) headersToAdd.Add(header);
            else if (CurrentExcelData != null) foreach (var header in CurrentExcelData.Headers) headersToAdd.Add(header);
            foreach (var h in headersToAdd.OrderBy(x => x)) AvailableColumns.Add(h);
        });
    }

    private void UpdateKeyColumns()
    {
        _dispatcherQueue.TryEnqueue(() => {
            AvailableKeyColumns.Clear(); string tempSelectedKeyColumn = string.Empty;
            if (IsMultiTableMode && _multiTableData != null) foreach (var header in _multiTableData.CommonHeaders.OrderBy(x => x))
                {
                    AvailableKeyColumns.Add(header);
                    if (string.IsNullOrEmpty(tempSelectedKeyColumn) &&
                        (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) || header.Contains("ID", StringComparison.OrdinalIgnoreCase) ||
                         header.Contains("编号", StringComparison.OrdinalIgnoreCase) || header.Contains("姓名", StringComparison.OrdinalIgnoreCase) ||
                         header.Contains("名字", StringComparison.OrdinalIgnoreCase)))
                        tempSelectedKeyColumn = header;
                }
            SelectedKeyColumn = tempSelectedKeyColumn;
        });
    }

    private void UpdateIdCardColumns()
    {
        _dispatcherQueue.TryEnqueue(() => {
            AvailableIdCardColumns.Clear(); string tempSelectedIdCardCol = string.Empty;
            List<string> headers = IsMultiTableMode ? (_multiTableData?.AllHeaders.ToList() ?? new List<string>()) : (CurrentExcelData?.Headers ?? new List<string>());
            foreach (var header in headers.OrderBy(x => x)) if (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) || header.Contains("证件", StringComparison.OrdinalIgnoreCase) || header.Contains("ID", StringComparison.OrdinalIgnoreCase))
                {
                    AvailableIdCardColumns.Add(header);
                    if (string.IsNullOrEmpty(tempSelectedIdCardCol) && header.Contains("身份证", StringComparison.OrdinalIgnoreCase)) tempSelectedIdCardCol = header;
                }
            if (string.IsNullOrEmpty(tempSelectedIdCardCol) && AvailableIdCardColumns.Count > 0) tempSelectedIdCardCol = AvailableIdCardColumns[0];
            SelectedIdCardColumn = tempSelectedIdCardCol;
            if (EnableIdCardExtraction && string.IsNullOrEmpty(SelectedIdCardColumn) && headers.Any()) StatusMessage = "警告：已启用身份证信息提取，但未自动匹配到身份证列";
        });
    }

    [RelayCommand]
    private void ToggleTableMode()
    {
        IsMultiTableMode = !IsMultiTableMode;
        _dispatcherQueue.TryEnqueue(() => {
            if (IsMultiTableMode)
            {
                _multiTableData.Clear();
                if (CurrentExcelData != null && CurrentExcelData.Rows.Any())
                {
                    var newExcelData = new ExcelData
                    {
                        Headers = new List<string>(CurrentExcelData.Headers),
                        Rows = new List<Dictionary<string, string>>(CurrentExcelData.Rows.Select(r => new Dictionary<string, string>(r))),
                        SourceFileName = CurrentExcelData.SourceFileName
                    };
                    _multiTableData.Tables.Add(newExcelData);
                }
                NotifyMultiTableDependentProperties(); // 确保更新所有依赖项
                StatusMessage = "已切换到多表格模式";
            }
            else
            {
                var firstTable = _multiTableData.Tables.FirstOrDefault();
                CurrentExcelData = firstTable != null ? new ExcelData
                {
                    Headers = new List<string>(firstTable.Headers),
                    Rows = new List<Dictionary<string, string>>(firstTable.Rows.Select(r => new Dictionary<string, string>(r))),
                    SourceFileName = firstTable.SourceFileName
                } : new ExcelData();
                _excelData = CurrentExcelData;
                NotifyMultiTableDependentProperties(); // 确保更新所有依赖项
                StatusMessage = "已切换到单表格模式";
            }
        });
    }

    [RelayCommand]
    private void RemoveTable(ExcelData? table)
    {
        if (table != null && _multiTableData.Tables.Contains(table))
        {
            _multiTableData.Tables.Remove(table);
            _dispatcherQueue.TryEnqueue(() => {
                NotifyMultiTableDependentProperties();
                StatusMessage = $"已移除表格: {table.SourceFileName}";
            });
        }
    }

    async partial void OnSelectedKeyColumnChanged(string? oldValue, string? newValue)
    {
        if (IsMultiTableMode && !string.IsNullOrEmpty(newValue))
        {
            await Task.Run(() => _multiTableData.MergeData(newValue));
            _dispatcherQueue.TryEnqueue(() => {
                StatusMessage = $"已使用 “{newValue}” 列合并数据，共有 {_multiTableData.MergedRows.Count} 条记录";
                OnPropertyChanged(nameof(MergedDataSummary));
            });
        }
        else if (IsMultiTableMode && string.IsNullOrEmpty(newValue))
        {
            if (_multiTableData.MergedRows.Any()) _multiTableData.MergedRows.Clear();
            _dispatcherQueue.TryEnqueue(() => {
                OnPropertyChanged(nameof(MergedDataSummary));
                StatusMessage = "请选择匹配键列以合并数据";
            });
        }
    }

    [RelayCommand]
    public async Task HandleDroppedFilesAsync(string[]? files)
    {
        if (files == null || files.Length == 0) return;
        try
        {
            var file = files[0];
            if (Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) { ExcelFilePath = file; }
            else if (Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase)) { WordTemplatePath = file; }
            else { _dispatcherQueue.TryEnqueue(() => StatusMessage = "不支持的文件类型"); }
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"处理拖放文件时出错: {ex.Message}"); }
    }

    [RelayCommand]
    private void CopyToClipboard(string? text)
    {
        if (string.IsNullOrEmpty(text)) return;
        try
        {
            DataPackage dataPackage = new DataPackage(); dataPackage.SetText(text);
            Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(dataPackage);
            _dispatcherQueue.TryEnqueue(() => StatusMessage = $"已复制「{text}」到剪贴板");
            var timer = _dispatcherQueue.CreateTimer(); timer.Interval = TimeSpan.FromSeconds(3); timer.IsRepeating = false;
            timer.Tick += (s, e) => _dispatcherQueue.TryEnqueue(() => { if (StatusMessage.StartsWith($"已复制「{text}」")) StatusMessage = "准备就绪"; });
            timer.Start();
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"复制到剪贴板失败: {ex.Message}"); }
    }

    [RelayCommand]
    private async Task GenerateDocuments()
    {
        if (string.IsNullOrEmpty(OutputDirectory) || !Directory.Exists(OutputDirectory)) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请选择有效的输出目录"); return; }
        var hasWordTemplate = !string.IsNullOrEmpty(WordTemplatePath) && File.Exists(WordTemplatePath);
        var hasExcelTemplate = UseExcelTemplate && !string.IsNullOrEmpty(ExcelTemplatePath) && File.Exists(ExcelTemplatePath);
        if (!hasWordTemplate && !hasExcelTemplate) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请至少选择一个Word模板或Excel模板"); return; }

        List<Dictionary<string, string>> dataRows;
        if (IsMultiTableMode)
        {
            if (!_multiTableData.Tables.Any()) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请先添加Excel文件"); return; }
            if (string.IsNullOrEmpty(SelectedKeyColumn)) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请选择用于匹配记录的列"); return; }
            dataRows = _multiTableData.MergedRows;
        }
        else
        {
            dataRows = _excelData.Rows;
        }
        if (dataRows == null || !dataRows.Any()) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "没有可处理的数据"); return; }

        _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; ProgressValue = 0; TotalItems = dataRows.Count; ProcessedItems = 0; ProcessResultText = string.Empty; });
        var localSuccessCount = 0; var localFailCount = 0;

        try
        {
            await Task.Run(async () => { // Offload the loop to a background thread
                for (var i = 0; i < dataRows.Count; i++)
                {
                    if (!_isProcessing) break; // Allow cancellation
                    var rowData = new Dictionary<string, string>(dataRows[i]);
                    rowData["序号"] = (i + 1).ToString();
                    rowData["时间"] = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                    rowData["日期"] = DateTime.Now.ToString("yyyy-MM-dd");
                    if (EnableIdCardExtraction && !string.IsNullOrEmpty(SelectedIdCardColumn) && rowData.TryGetValue(SelectedIdCardColumn, out var idCard) && !string.IsNullOrEmpty(idCard))
                    {
                        try
                        {
                            rowData["身份证性别"] = _idCardService.ExtractGender(idCard);
                            // ... (extract other ID card fields) ...
                        }
                        catch { /* Handle ID extraction error silently or log */ }
                    }
                    var fileName = GenerateOutputFileName(rowData, i);

                    if (hasWordTemplate)
                    {
                        var wordOutputPath = Path.Combine(OutputDirectory, $"{fileName}.docx");
                        var wordResult = await _wordService.ProcessTemplateAsync(WordTemplatePath, wordOutputPath, rowData, null);
                        if (wordResult.Success) localSuccessCount++; else localFailCount++;
                    }
                    if (hasExcelTemplate)
                    {
                        var excelOutputPath = Path.Combine(OutputDirectory, $"{fileName}.xlsx");
                        var excelResult = UseImageReplacement && ImageDirectories.Any()
                           ? await _excelTemplateService.ProcessTemplateWithImagesAsync(ExcelTemplatePath, excelOutputPath, rowData, ImageDirectories, ImageFillMode, ImageFillPercentage, null)
                           : await _excelTemplateService.ProcessTemplateAsync(ExcelTemplatePath, excelOutputPath, rowData, null);
                        if (excelResult.Success) localSuccessCount++; else localFailCount++;
                    }
                    // These are [ObservableProperty], Community Toolkit handles dispatching if set from background
                    ProcessedItems = i + 1;
                    ProgressValue = (int)((double)ProcessedItems / TotalItems * 100);
                }
            });
            _dispatcherQueue.TryEnqueue(() => {
                ProcessResultText = $"处理完成：成功 {localSuccessCount} 个，失败 {localFailCount} 个"; ProcessSuccess = localFailCount == 0;
                StatusMessage = $"文档生成完成，输出到 {OutputDirectory}"; ProgressValue = 100;
            });
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => { StatusMessage = $"处理过程中出错: {ex.Message}"; ProcessSuccess = false; ProcessResultText = $"处理失败: {ex.Message}"; }); }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }

    private string GenerateOutputFileName(Dictionary<string, string> rowData, int index)
    {
        var fileName = OutputFileNameTemplate;
        foreach (var item in rowData) fileName = fileName.Replace($"{{{item.Key}}}", item.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        fileName = fileName.Replace("{序号}", rowData.ContainsKey("序号") ? rowData["序号"] : (index + 1).ToString(), StringComparison.OrdinalIgnoreCase);
        fileName = fileName.Replace("{时间}", rowData.ContainsKey("时间") ? rowData["时间"] : DateTime.Now.ToString("yyyyMMdd-HHmmss"), StringComparison.OrdinalIgnoreCase);
        fileName = fileName.Replace("{日期}", rowData.ContainsKey("日期") ? rowData["日期"] : DateTime.Now.ToString("yyyy-MM-dd"), StringComparison.OrdinalIgnoreCase);
        foreach (var invalidChar in Path.GetInvalidFileNameChars()) fileName = fileName.Replace(invalidChar.ToString(), "_"); // Use ToString() for Replace
        return string.IsNullOrWhiteSpace(fileName) || fileName.All(c => c == '_') ? $"Document_{index + 1}_{DateTime.Now:yyyyMMddHHmmss}" : fileName;
    }

    [RelayCommand]
    private async Task CheckPlaceholders()
    {
        if (string.IsNullOrEmpty(WordTemplatePath) || !File.Exists(WordTemplatePath)) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请先选择有效的Word模板"); return; }
        try
        {
            _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = "正在检查Word模板中的占位符..."; });
            var placeholders = await _wordService.ExtractPlaceholdersNpoiAsync(WordTemplatePath);
            _dispatcherQueue.TryEnqueue(() => {
                if (placeholders.Any())
                {
                    StatusMessage = $"模板中找到 {placeholders.Count} 个占位符";
                    bool foundIdCardPlaceholders = false;
                    foreach (var p in PlaceholderConstants.AllPlaceholders) if (placeholders.Contains(p)) { foundIdCardPlaceholders = true; break; }
                    if (foundIdCardPlaceholders && !EnableIdCardExtraction)
                    {
                        EnableIdCardExtraction = true; StatusMessage += " (已自动启用身份证信息提取功能)";
                    }
                }
                else { StatusMessage = "模板中未找到任何占位符"; }
            });
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"检查占位符失败: {ex.Message}"); }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }

    [RelayCommand]
    private void OpenOutputFolder()
    {
        if (Directory.Exists(OutputDirectory))
            try { Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"\"{OutputDirectory}\"", UseShellExecute = true }); }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"打开输出文件夹失败: {ex.Message}"); }
        else _dispatcherQueue.TryEnqueue(() => StatusMessage = "输出目录不存在");
    }

    private string CalculateAge(string birthDateStr) { /* ... 原有逻辑 ... */ return "未知"; }

    [RelayCommand]
    private async Task CheckExcelPlaceholders()
    {
        if (string.IsNullOrEmpty(ExcelTemplatePath) || !File.Exists(ExcelTemplatePath)) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请先选择有效的Excel模板"); return; }
        try
        {
            _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = "正在检查Excel模板中的占位符..."; });
            var placeholders = await _excelTemplateService.ExtractPlaceholdersAsync(ExcelTemplatePath);
            _dispatcherQueue.TryEnqueue(() => {
                DetectedExcelPlaceholders.Clear();
                foreach (var placeholder in placeholders) DetectedExcelPlaceholders.Add(placeholder);
                StatusMessage = placeholders.Any() ? $"Excel模板中找到 {placeholders.Count} 个占位符" : "Excel模板中未找到任何占位符";
            });
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"检查Excel占位符失败: {ex.Message}"); }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }

    [RelayCommand] private void ClearExcelTemplate() { ExcelTemplatePath = string.Empty; _dispatcherQueue.TryEnqueue(() => DetectedExcelPlaceholders.Clear()); UseExcelTemplate = false; }
    [RelayCommand] private void RemoveImageDirectory(ImageSourceDirectory? directory) { if (directory != null && ImageDirectories.Remove(directory)) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"已移除图片目录: {directory.DirectoryName}"); if (ImageDirectories.Count == 0) UseImageReplacement = false; } }

    [RelayCommand]
    private async Task CheckDirectoryImages(ImageSourceDirectory? directory)
    {
        if (directory == null || string.IsNullOrEmpty(directory.DirectoryPath) || !Directory.Exists(directory.DirectoryPath)) return;
        try
        {
            _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = $"正在扫描目录 {directory.DirectoryName} 中的图片..."; });
            var imageFiles = await _imageProcessingService.ScanDirectoryForImagesAsync(directory.DirectoryPath);
            _dispatcherQueue.TryEnqueue(() => {
                if (directory.ImageFiles == null) directory.ImageFiles = new ObservableCollection<string>(); // Ensure initialized
                directory.ImageFiles.Clear();
                foreach (var file in imageFiles) directory.ImageFiles.Add(file);
                StatusMessage = $"目录 {directory.DirectoryName} 中找到 {imageFiles.Count} 个图片";
            });
        }
        catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"扫描图片目录失败: {ex.Message}"); }
        finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
    }
    [RelayCommand] private void ClearAllImageDirectories() { ImageDirectories.Clear(); UseImageReplacement = false; _dispatcherQueue.TryEnqueue(() => StatusMessage = "已清除所有图片目录"); }

    // On<PropertyName>Changed partial methods for [ObservableProperty]
    // These are called AFTER the property has changed.
    partial void OnAvailableColumnsChanged(ObservableCollection<string>? oldValue, ObservableCollection<string> newValue)
    {
        // Update SelectedImageMatchingColumn based on new available columns
        _dispatcherQueue.TryEnqueue(() => {
            if (!string.IsNullOrEmpty(SelectedImageMatchingColumn) && (newValue?.Contains(SelectedImageMatchingColumn) == true)) { /* Keep current */ }
            else if (newValue?.Count > 0)
            {
                var preferredColumns = newValue.Where(c =>
                    c.Contains("姓名", StringComparison.OrdinalIgnoreCase) || c.Contains("名字", StringComparison.OrdinalIgnoreCase) ||
                    c.Contains("ID", StringComparison.OrdinalIgnoreCase) || c.Contains("编号", StringComparison.OrdinalIgnoreCase) ||
                    c.Contains("身份证", StringComparison.OrdinalIgnoreCase)).ToList();
                SelectedImageMatchingColumn = preferredColumns.Any() ? preferredColumns.First() : newValue.First();
            }
            else { SelectedImageMatchingColumn = string.Empty; }
        });
    }

    partial void OnSelectedImageFillModeItemChanged(ImageFillModeItem? oldValue, ImageFillModeItem newValue)
    {
        if (newValue != null) ImageFillMode = newValue.Value;
    }
}