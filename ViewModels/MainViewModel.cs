using DocTransform.Constants;
using DocTransform.Models;
using DocTransform.Mvvm; // 辅助类命名空间
using DocTransform.Services;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input; // ICommand 接口
using Windows.ApplicationModel.DataTransfer;
using Windows.UI; // Color


namespace DocTransform.ViewModels
{
    public class MainViewModel : ObservableObject // 继承自创建的 ObservableObject
    {
        // 服务和 DispatcherQueue 保持不变
        private readonly ExcelService _excelService;
        private readonly WordService _wordService;
        private readonly IdCardService _idCardService;
        private readonly ImageProcessingService _imageProcessingService;
        private readonly ExcelTemplateService _excelTemplateService;
        private readonly DispatcherQueue _dispatcherQueue;

        // --- 后台字段 (Backing Fields) ---
        // 为所有之前使用 [ObservableProperty] 的属性创建私有字段
        private ExcelData _excelData = new();
        private string _excelFilePath = string.Empty;
        private bool _isMultiTableMode;
        private string _selectedKey = string.Empty;
        private MultiTableData _multiTableData; // 将在构造函数中初始化
        private ObservableCollection<string> _availableColumns = new();
        private ObservableCollection<string> _availableIdCardColumns = new();
        private ObservableCollection<string> _availableKeys = new();
        private ExcelData _currentExcelData = new();
        private ObservableCollection<string> _detectedExcelPlaceholders = new();
        private bool _enableIdCardExtraction;
        private string _excelTemplatePath = string.Empty;
        private List<string> _idCardPlaceholders = PlaceholderConstants.AllPlaceholders;
        private ObservableCollection<ImageSourceDirectory> _imageDirectories = new();
        private ImageFillMode _imageFillMode = ImageFillMode.Fit;
        private List<ImageFillModeItem> _imageFillModeItems = ImageFillModeItem.GetAll();
        private int _imageFillPercentage = 90;
        private bool _isProcessing;
        private string _outputFileNameTemplate = "{序号}_{姓名}_{时间}";
        private int _processedItems;
        private string _processResultText = string.Empty;
        private bool _processSuccess;
        private int _progressValue;
        private ObservableCollection<string> _selectedColumns = new();
        private string _selectedIdCardColumn = string.Empty;
        private ImageFillModeItem _selectedImageFillModeItem; // 将在构造函数中初始化
        private string _selectedImageMatchingColumn = string.Empty;
        private string _statusMessage = "准备就绪";
        private int _totalItems;
        private bool _useExcelTemplate;
        private bool _useImageReplacement;
        private string _wordTemplatePath = string.Empty;
        private string _outputDirectory = string.Empty; // 将在构造函数中初始化
        private int _valueA;
        private string _valueBText = string.Empty;


        // --- 命令属性 (ICommand Properties) ---
        public ICommand BrowseExcelFileCommand { get; }
        public ICommand BrowseWordTemplateCommand { get; }
        public ICommand BrowseOutputDirectoryCommand { get; }
        public ICommand BrowseExcelTemplateCommand { get; }
        public ICommand AddImageDirectoryCommand { get; }
        public ICommand ToggleTableModeCommand { get; }
        public ICommand RemoveTableCommand { get; }
        public ICommand CopyToClipboardCommand { get; }
        public ICommand GenerateDocumentsCommand { get; }
        public ICommand HandleDroppedFilesCommand { get; }
        public ICommand CheckPlaceholdersCommand { get; }
        public ICommand OpenOutputFolderCommand { get; }
        public ICommand CheckExcelPlaceholdersCommand { get; }
        public ICommand ClearExcelTemplateCommand { get; }
        public ICommand RemoveImageDirectoryCommand { get; }
        public ICommand CheckDirectoryImagesCommand { get; }
        public ICommand ClearAllImageDirectoriesCommand { get; }

        // 构造函数
        public MainViewModel(DispatcherQueue dispatcherQueue)
        {
            _dispatcherQueue = dispatcherQueue;
            _excelService = new ExcelService();
            _wordService = new WordService();
            _idCardService = new IdCardService();
            _imageProcessingService = new ImageProcessingService();
            _excelTemplateService = new ExcelTemplateService(_imageProcessingService);

            // 初始化属性
            OutputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // 这会调用 setter
            _selectedImageFillModeItem = _imageFillModeItems.FirstOrDefault(item => item.Value == ImageFillMode.Fit) ?? _imageFillModeItems.First();
            _multiTableData = new MultiTableData();
            SubscribeToMultiTableDataChanges(_multiTableData);

            // 初始化命令
            BrowseExcelFileCommand = new RelayCommand(_ => BrowseExcelFile());
            BrowseWordTemplateCommand = new RelayCommand(_ => BrowseWordTemplate());
            BrowseOutputDirectoryCommand = new RelayCommand(_ => BrowseOutputDirectory());
            BrowseExcelTemplateCommand = new RelayCommand(_ => BrowseExcelTemplate());
            AddImageDirectoryCommand = new RelayCommand(_ => AddImageDirectory());
            ToggleTableModeCommand = new RelayCommand(_ => ToggleTableMode());
            RemoveTableCommand = new RelayCommand(param => RemoveTable(param as ExcelData)); // 参数化命令
            CopyToClipboardCommand = new RelayCommand(param => CopyToClipboard(param as string)); // 参数化命令
            OpenOutputFolderCommand = new RelayCommand(_ => OpenOutputFolder());
            ClearExcelTemplateCommand = new RelayCommand(_ => ClearExcelTemplate());
            RemoveImageDirectoryCommand = new RelayCommand(param => RemoveImageDirectory(param as ImageSourceDirectory)); // 参数化命令
            ClearAllImageDirectoriesCommand = new RelayCommand(_ => ClearAllImageDirectories());

            // 异步命令
            GenerateDocumentsCommand = new AsyncRelayCommand(GenerateDocuments);
            HandleDroppedFilesCommand = new AsyncRelayCommand<string[]>(HandleDroppedFilesAsync); // 拖放的文件数组需要从View传递
            CheckPlaceholdersCommand = new AsyncRelayCommand(CheckPlaceholders);
            CheckExcelPlaceholdersCommand = new AsyncRelayCommand(CheckExcelPlaceholders);
            CheckDirectoryImagesCommand = new AsyncRelayCommand(async () => await CheckDirectoryImages(null)); // 目录对象需要从View传递
        }

        // --- 完整的公共属性 (Full Public Properties) ---
        // 为每个后台字段创建公共属性，并在 setter 中调用 OnPropertyChanged 和相关逻辑

        public string ExcelFilePath
        {
            get => _excelFilePath;
            set
            {
                if (SetProperty(ref _excelFilePath, value))
                {
                    OnExcelFilePathChanged(value); // 调用原 OnExcelFilePathChanged 逻辑
                    // 手动通知所有依赖此属性的计算属性
                    OnPropertyChanged(nameof(ExcelFilePathDisplayText));
                    OnPropertyChanged(nameof(ExcelFilePathIsPlaceholder));
                    OnPropertyChanged(nameof(SingleModeExcelDisplayText));
                    OnPropertyChanged(nameof(SingleModeExcelIsPlaceholder));
                    OnPropertyChanged(nameof(ExcelFilePathForegroundBrush));
                }
            }
        }

        public bool IsMultiTableMode
        {
            get => _isMultiTableMode;
            set
            {
                if (SetProperty(ref _isMultiTableMode, value))
                {
                    // 手动通知所有依赖此属性的计算属性
                    OnPropertyChanged(nameof(CurrentTableModeText));
                    OnPropertyChanged(nameof(CurrentTableModeForegroundBrush));
                    OnPropertyChanged(nameof(ExcelButtonText));
                    OnPropertyChanged(nameof(ExcelFilePathDisplayText));
                    OnPropertyChanged(nameof(ExcelFilePathIsPlaceholder));
                    OnPropertyChanged(nameof(MergeSummary));
                    OnPropertyChanged(nameof(ExcelFilePathForegroundBrush));
                }
            }
        }

        public string SelectedKey
        {
            get => _selectedKey;
            set
            {
                if (SetProperty(ref _selectedKey, value))
                {
                    OnSelectedKeyChanged(value); // 调用原 OnSelectedKeyChanged 逻辑
                    OnPropertyChanged(nameof(MergeSummary));
                }
            }
        }

        public MultiTableData MultiTableData
        {
            get => _multiTableData;
            set
            {
                var oldValue = _multiTableData;
                if (SetProperty(ref _multiTableData, value))
                {
                    OnMultiTableDataChanged(oldValue, value); // 调用原 OnMultiTableDataChanged 逻辑
                    // 手动通知依赖属性
                    OnPropertyChanged(nameof(ExcelFilePathDisplayText));
                    OnPropertyChanged(nameof(ExcelFilePathIsPlaceholder));
                    OnPropertyChanged(nameof(MultiModeExcelSummaryText));
                    OnPropertyChanged(nameof(MultiModeExcelIsPlaceholder));
                    OnPropertyChanged(nameof(MergeSummary));
                }
            }
        }

        public ObservableCollection<string> AvailableColumns
        {
            get => _availableColumns;
            set
            {
                if (SetProperty(ref _availableColumns, value))
                {
                    OnAvailableColumnsChanged(value); // 调用原 OnAvailableColumnsChanged 逻辑
                }
            }
        }

        public ObservableCollection<string> AvailableIdCardColumns
        {
            get => _availableIdCardColumns;
            set => SetProperty(ref _availableIdCardColumns, value);
        }

        public ObservableCollection<string> AvailableKeys
        {
            get => _availableKeys;
            set => SetProperty(ref _availableKeys, value);
        }

        public ExcelData CurrentExcelData
        {
            get => _currentExcelData;
            set
            {
                if (SetProperty(ref _currentExcelData, value))
                {
                    // 手动通知依赖属性
                    OnPropertyChanged(nameof(SingleModeExcelDisplayText));
                    OnPropertyChanged(nameof(SingleModeExcelIsPlaceholder));
                }
            }
        }

        public ObservableCollection<string> DetectedExcelPlaceholders
        {
            get => _detectedExcelPlaceholders;
            set => SetProperty(ref _detectedExcelPlaceholders, value);
        }

        public bool EnableIdCardExtraction
        {
            get => _enableIdCardExtraction;
            set
            {
                if (SetProperty(ref _enableIdCardExtraction, value))
                {
                    OnPropertyChanged(nameof(IdCardExtractionRelatedPropertiesMightChange));
                }
            }
        }

        public string ExcelTemplatePath
        {
            get => _excelTemplatePath;
            set
            {
                if (SetProperty(ref _excelTemplatePath, value))
                {
                    OnExcelTemplatePathChanged(value); // 调用原 OnExcelTemplatePathChanged 逻辑
                }
            }
        }

        public List<string> IdCardPlaceholders
        {
            get => _idCardPlaceholders;
            set
            {
                if (SetProperty(ref _idCardPlaceholders, value))
                {
                    OnPropertyChanged(nameof(IdCardPlaceholdersText));
                }
            }
        }

        public ObservableCollection<ImageSourceDirectory> ImageDirectories
        {
            get => _imageDirectories;
            set => SetProperty(ref _imageDirectories, value);
        }

        public ImageFillMode ImageFillMode
        {
            get => _imageFillMode;
            set => SetProperty(ref _imageFillMode, value);
        }

        public List<ImageFillModeItem> ImageFillModeItems
        {
            get => _imageFillModeItems;
            set => SetProperty(ref _imageFillModeItems, value);
        }

        public int ImageFillPercentage
        {
            get => _imageFillPercentage;
            set => SetProperty(ref _imageFillPercentage, value);
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set
            {
                if (SetProperty(ref _isProcessing, value))
                {
                    OnPropertyChanged(nameof(IsProcessingOverlayVisible));
                }
            }
        }

        public string OutputFileNameTemplate
        {
            get => _outputFileNameTemplate;
            set => SetProperty(ref _outputFileNameTemplate, value);
        }

        public int ProcessedItems
        {
            get => _processedItems;
            set => SetProperty(ref _processedItems, value);
        }

        public string ProcessResultText
        {
            get => _processResultText;
            set => SetProperty(ref _processResultText, value);
        }

        public bool ProcessSuccess
        {
            get => _processSuccess;
            set => SetProperty(ref _processSuccess, value);
        }

        public int ProgressValue
        {
            get => _progressValue;
            set => SetProperty(ref _progressValue, value);
        }

        public ObservableCollection<string> SelectedColumns
        {
            get => _selectedColumns;
            set => SetProperty(ref _selectedColumns, value);
        }

        public string SelectedIdCardColumn
        {
            get => _selectedIdCardColumn;
            set
            {
                if (SetProperty(ref _selectedIdCardColumn, value))
                {
                    OnPropertyChanged(nameof(IdCardExtractionRelatedPropertiesMightChange));
                }
            }
        }

        public ImageFillModeItem SelectedImageFillModeItem
        {
            get => _selectedImageFillModeItem;
            set
            {
                if (SetProperty(ref _selectedImageFillModeItem, value))
                {
                    OnSelectedImageFillModeItemChanged(value); // 调用原 OnSelectedImageFillModeItemChanged 逻辑
                }
            }
        }

        public string SelectedImageMatchingColumn
        {
            get => _selectedImageMatchingColumn;
            set => SetProperty(ref _selectedImageMatchingColumn, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int TotalItems
        {
            get => _totalItems;
            set => SetProperty(ref _totalItems, value);
        }

        public bool UseExcelTemplate
        {
            get => _useExcelTemplate;
            set => SetProperty(ref _useExcelTemplate, value);
        }

        public bool UseImageReplacement
        {
            get => _useImageReplacement;
            set => SetProperty(ref _useImageReplacement, value);
        }

        public string WordTemplatePath
        {
            get => _wordTemplatePath;
            set
            {
                if (SetProperty(ref _wordTemplatePath, value))
                {
                    OnWordTemplatePathChanged(value); // 调用原 OnWordTemplatePathChanged 逻辑
                    OnPropertyChanged(nameof(WordTemplatePathDisplayText));
                    OnPropertyChanged(nameof(WordTemplatePathIsPlaceholder));
                    OnPropertyChanged(nameof(WordTemplatePathForegroundBrush));
                }
            }
        }

        public string OutputDirectory
        {
            get => _outputDirectory;
            set
            {
                if (SetProperty(ref _outputDirectory, value))
                {
                    OnPropertyChanged(nameof(IsOutputFolderAvailable));
                    OnPropertyChanged(nameof(OutputDirectoryDisplayText));
                    OnPropertyChanged(nameof(OutputDirectoryIsPlaceholder));
                    OnPropertyChanged(nameof(OutputDirectoryForegroundBrush));
                }
            }
        }

        public int ValueA
        {
            get => _valueA;
            set
            {
                if (SetProperty(ref _valueA, value))
                {
                    OnPropertyChanged(nameof(IsAGreaterThanB)); // 依赖属性
                }
            }
        }

        public string ValueBText
        {
            get => _valueBText;
            set
            {
                if (SetProperty(ref _valueBText, value))
                {
                    OnPropertyChanged(nameof(IsAGreaterThanB)); // 依赖属性
                }
            }
        }


        // --- 计算属性 (Computed Properties - Read-Only) ---
        // 这些属性的 get 访问器包含逻辑，并且它们依赖的属性的 setter 会通知它们更新

        public bool IsAGreaterThanB
        {
            get
            {
                if (int.TryParse(ValueBText, out var valueB))
                {
                    return ValueA > valueB;
                }
                return false;
            }
        }

        public bool IsProcessingOverlayVisible => IsProcessing;

        public string CurrentTableModeText => IsMultiTableMode ? "多表格" : "单表格";

        public Brush CurrentTableModeForegroundBrush => new SolidColorBrush(IsMultiTableMode ? ((Color)Application.Current.Resources["SystemAccentColor"]) : ((Color)Application.Current.Resources["TextFillColorPrimary"]));

        public string ExcelButtonText => IsMultiTableMode ? "添加Excel" : "浏览";

        public string ExcelFilePathDisplayText => IsMultiTableMode ? MultiModeExcelSummaryText : SingleModeExcelDisplayText;

        public bool ExcelFilePathIsPlaceholder => IsMultiTableMode ? MultiModeExcelIsPlaceholder : SingleModeExcelIsPlaceholder;

        public string SingleModeExcelDisplayText => string.IsNullOrEmpty(ExcelFilePath) ? "未选择Excel文件" : Path.GetFileName(ExcelFilePath);

        public bool SingleModeExcelIsPlaceholder => string.IsNullOrEmpty(ExcelFilePath);

        public string MultiModeExcelSummaryText => (_multiTableData?.Tables?.Count ?? 0) == 0 ? "未添加Excel表格" : $"已添加 {_multiTableData.Tables.Count} 个表格";

        public bool MultiModeExcelIsPlaceholder => (_multiTableData?.Tables?.Count ?? 0) == 0;

        public bool IsOutputFolderAvailable => !string.IsNullOrEmpty(OutputDirectory) && Directory.Exists(OutputDirectory);

        public string MergeSummary =>
            (IsMultiTableMode && !string.IsNullOrEmpty(SelectedKey) && (_multiTableData?.MergedRows?.Count ?? 0) > 0)
                ? $"已使用 “{SelectedKey}” 列合并数据，共有 {_multiTableData.MergedRows.Count} 条记录"
                : (IsMultiTableMode && !string.IsNullOrEmpty(SelectedKey) ? "合并后无匹配数据或未选择有效的键列" : "请选择匹配键以生成预览...");

        public Brush ExcelFilePathForegroundBrush => ExcelFilePathIsPlaceholder
            ? (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
            : (Brush)Application.Current.Resources["TextFillColorPrimaryBrush"];

        public Brush WordTemplatePathForegroundBrush => WordTemplatePathIsPlaceholder
            ? (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
            : (Brush)Application.Current.Resources["TextFillColorPrimaryBrush"];

        public Brush OutputDirectoryForegroundBrush => OutputDirectoryIsPlaceholder
            ? (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
            : (Brush)Application.Current.Resources["TextFillColorPrimaryBrush"];

        public string WordTemplatePathDisplayText => string.IsNullOrEmpty(WordTemplatePath) ? "未选择Word模板文件 (*.docx)" : Path.GetFileName(WordTemplatePath);

        public bool WordTemplatePathIsPlaceholder => string.IsNullOrEmpty(WordTemplatePath);

        public string OutputDirectoryDisplayText => string.IsNullOrEmpty(OutputDirectory) ? "未选择输出文件夹" : OutputDirectory;

        public bool OutputDirectoryIsPlaceholder => string.IsNullOrEmpty(OutputDirectory);

        public string IdCardPlaceholdersText => string.Join(", ", _idCardPlaceholders);

        public bool IdCardExtractionRelatedPropertiesMightChange => true; // 这个可以保持为true，或者根据实际逻辑调整


        // --- 方法 (Methods) ---
        // 原先的 partial void On...Changed 方法现在是普通的私有方法，由属性的 setter 调用

        private void OnMultiTableDataChanged(MultiTableData? oldValue, MultiTableData newValue)
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
            OnPropertyChanged(nameof(MergeSummary));
            UpdateKeyColumns(); UpdateAvailableColumns(); UpdateIdCardColumns();
        }

        // 命令对应的方法
        private void BrowseExcelFile() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择Excel文件...");
        private void BrowseWordTemplate() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择Word模板...");
        private void BrowseOutputDirectory() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择输出文件夹...");
        private void BrowseExcelTemplate() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择Excel模板...");
        private void AddImageDirectory() => _dispatcherQueue.TryEnqueue(() => StatusMessage = "请通过文件对话框选择图片目录...");

        // 属性变化时调用的逻辑方法
        private async void OnExcelFilePathChanged(string? newValue) // 注意参数名可能与旧版不同
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

        private async void OnWordTemplatePathChanged(string? newValue)
        {
            if (!string.IsNullOrEmpty(newValue) && File.Exists(newValue))
            {
                var isValid = await _wordService.IsValidTemplateAsync(newValue);
                _dispatcherQueue.TryEnqueue(() => {
                    if (isValid) { CheckPlaceholdersCommand?.Execute(null); } // 假设 CheckPlaceholdersCommand 已初始化
                    else { StatusMessage = "选择的Word模板无效"; }
                });
            }
        }

        private async void OnExcelTemplatePathChanged(string? newValue)
        {
            if (!string.IsNullOrEmpty(newValue) && File.Exists(newValue))
            {
                var isValid = await ExcelTemplateService.IsValidTemplateAsync(newValue); // 假设 ExcelTemplateService 是静态或已实例化
                _dispatcherQueue.TryEnqueue(() => {
                    if (!isValid) { StatusMessage = "选择的Excel模板无效"; }
                    else { CheckExcelPlaceholdersCommand?.Execute(null); } // 假设 CheckExcelPlaceholdersCommand 已初始化
                });
            }
        }

        private void OnAvailableColumnsChanged(ObservableCollection<string> newValue)
        {
            _dispatcherQueue.TryEnqueue(() => {
                if (!string.IsNullOrEmpty(SelectedImageMatchingColumn) && (newValue?.Contains(SelectedImageMatchingColumn) == true)) { }
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

        private void OnSelectedImageFillModeItemChanged(ImageFillModeItem newValue)
        {
            if (newValue != null) ImageFillMode = newValue.Value;
        }


        public async Task ProcessAddedImageDirectory(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath)) return;
            try
            {
                _dispatcherQueue.TryEnqueue(() => IsProcessing = true);
                var directoryName = Path.GetFileName(directoryPath);
                if (string.IsNullOrEmpty(directoryName)) directoryName = new DirectoryInfo(directoryPath).Name;

                _dispatcherQueue.TryEnqueue(() => StatusMessage = "正在扫描图片目录...");
                var imageFiles = await ImageProcessingService.ScanDirectoryForImagesAsync(directoryPath);

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
                    CurrentExcelData = loadedData; _excelData = loadedData; // _excelData 赋值也应通过属性以触发通知（如果需要）
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
                    foreach (var sheet in allSheets) _multiTableData.Tables.Add(sheet); // MultiTableData 内部应处理 CollectionChanged
                    NotifyMultiTableDependentProperties(); // 确保相关UI更新
                    StatusMessage = $"添加了 {allSheets.Count} 个工作表，共有 {_multiTableData.TotalRowCount} 行数据";
                });
            }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"加载Excel文件失败: {ex.Message}"); }
            finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
        }

        private void UpdateAvailableColumns()
        {
            _dispatcherQueue.TryEnqueue(() => {
                var newAvailableColumns = new ObservableCollection<string>();
                var newSelectedColumns = new ObservableCollection<string>(); // 如果 SelectedColumns 也需要重置
                var headersToAdd = new HashSet<string>();

                if (IsMultiTableMode) foreach (var header in _multiTableData.AllHeaders) headersToAdd.Add(header);
                else if (CurrentExcelData != null) foreach (var header in CurrentExcelData.Headers) headersToAdd.Add(header);

                foreach (var h in headersToAdd.OrderBy(x => x)) newAvailableColumns.Add(h);

                AvailableColumns = newAvailableColumns; // 触发属性更改
                SelectedColumns = newSelectedColumns; // 触发属性更改
            });
        }

        private void UpdateKeyColumns()
        {
            _dispatcherQueue.TryEnqueue(() => {
                var newAvailableKeys = new ObservableCollection<string>();
                string tempSelectedKeyColumn = string.Empty;
                if (IsMultiTableMode && _multiTableData != null)
                {
                    foreach (var header in _multiTableData.CommonHeaders.OrderBy(x => x))
                    {
                        newAvailableKeys.Add(header);
                        if (string.IsNullOrEmpty(tempSelectedKeyColumn) &&
                            (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) || header.Contains("ID", StringComparison.OrdinalIgnoreCase) ||
                             header.Contains("编号", StringComparison.OrdinalIgnoreCase) || header.Contains("姓名", StringComparison.OrdinalIgnoreCase) ||
                             header.Contains("名字", StringComparison.OrdinalIgnoreCase)))
                            tempSelectedKeyColumn = header;
                    }
                }
                AvailableKeys = newAvailableKeys; // 触发属性更改
                SelectedKey = tempSelectedKeyColumn; // 触发属性更改
            });
        }

        private void UpdateIdCardColumns()
        {
            _dispatcherQueue.TryEnqueue(() => {
                var newAvailableIdCardColumns = new ObservableCollection<string>();
                string tempSelectedIdCardCol = string.Empty;

                List<string> headers;
                if (IsMultiTableMode)
                {
                    headers = _multiTableData?.AllHeaders.ToList() ?? new List<string>();
                }
                else
                {
                    headers = CurrentExcelData?.Headers?.ToList() ?? new List<string>();
                }

                foreach (var header in headers.OrderBy(x => x))
                {
                    if (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) || header.Contains("证件", StringComparison.OrdinalIgnoreCase) || header.Contains("ID", StringComparison.OrdinalIgnoreCase))
                    {
                        newAvailableIdCardColumns.Add(header);
                        if (string.IsNullOrEmpty(tempSelectedIdCardCol) && header.Contains("身份证", StringComparison.OrdinalIgnoreCase)) tempSelectedIdCardCol = header;
                    }
                }
                if (string.IsNullOrEmpty(tempSelectedIdCardCol) && newAvailableIdCardColumns.Count > 0) tempSelectedIdCardCol = newAvailableIdCardColumns[0];

                AvailableIdCardColumns = newAvailableIdCardColumns; // 触发属性更改
                SelectedIdCardColumn = tempSelectedIdCardCol; // 触发属性更改

                if (EnableIdCardExtraction && string.IsNullOrEmpty(SelectedIdCardColumn) && headers.Any()) StatusMessage = "警告：已启用身份证信息提取，但未自动匹配到身份证列";
            });
        }

        private void ToggleTableMode() // 命令方法
        {
            IsMultiTableMode = !IsMultiTableMode; // 这会触发 IsMultiTableMode 的 setter 中的所有依赖通知
            _dispatcherQueue.TryEnqueue(() => {
                if (IsMultiTableMode)
                {
                    _multiTableData.Clear(); // 假设 MultiTableData 有 Clear 方法并能正确通知
                    if (CurrentExcelData != null && CurrentExcelData.Rows.Any())
                    {
                        var newExcelData = new ExcelData
                        {
                            Headers = new ObservableCollection<string>(CurrentExcelData.Headers),
                            Rows = new ObservableCollection<Dictionary<string, string>>(CurrentExcelData.Rows.Select(r => new Dictionary<string, string>(r))),
                            SourceFileName = CurrentExcelData.SourceFileName
                        };
                        _multiTableData.Tables.Add(newExcelData); // 假设 Tables 是 ObservableCollection
                    }
                    NotifyMultiTableDependentProperties();
                    StatusMessage = "已切换到多表格模式";
                }
                else
                {
                    var firstTable = _multiTableData.Tables.FirstOrDefault();
                    CurrentExcelData = firstTable != null ? new ExcelData
                    {
                        Headers = new ObservableCollection<string>(firstTable.Headers),
                        Rows = new ObservableCollection<Dictionary<string, string>>(firstTable.Rows.Select(r => new Dictionary<string, string>(r))),
                        SourceFileName = firstTable.SourceFileName
                    } : new ExcelData();
                    _excelData = CurrentExcelData; // 直接赋值，如果 _excelData 也需要通知，则应通过属性
                    NotifyMultiTableDependentProperties();
                    StatusMessage = "已切换到单表格模式";
                }
            });
        }

        private void RemoveTable(ExcelData? table) // 命令方法
        {
            if (table != null && _multiTableData.Tables.Contains(table))
            {
                _multiTableData.Tables.Remove(table); // 假设 Tables 是 ObservableCollection
                _dispatcherQueue.TryEnqueue(() => {
                    NotifyMultiTableDependentProperties();
                    StatusMessage = $"已移除表格: {table.SourceFileName}";
                });
            }
        }

        private async void OnSelectedKeyChanged(string? newValue) // 属性变化时调用的逻辑
        {
            if (IsMultiTableMode && !string.IsNullOrEmpty(newValue))
            {
                await Task.Run(() => _multiTableData.MergeData(newValue)); // 假设 MergeData 存在
                _dispatcherQueue.TryEnqueue(() => {
                    StatusMessage = $"已使用 “{newValue}” 列合并数据，共有 {_multiTableData.MergedRows.Count} 条记录";
                    OnPropertyChanged(nameof(MergeSummary)); // 手动通知
                });
            }
            else if (IsMultiTableMode && string.IsNullOrEmpty(newValue))
            {
                if (_multiTableData.MergedRows.Any()) _multiTableData.MergedRows.Clear(); // 假设 MergedRows 是可清除的集合
                _dispatcherQueue.TryEnqueue(() => {
                    OnPropertyChanged(nameof(MergeSummary)); // 手动通知
                    StatusMessage = "请选择匹配键列以合并数据";
                });
            }
        }

        public async Task HandleDroppedFilesAsync(string[]? files)
        {
            if (files == null || files.Length == 0) return;
            try
            {
                var file = files[0]; // 只处理第一个文件作为示例
                if (Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    ExcelFilePath = file; // 假设 ExcelFilePath 属性的 setter 会处理通知
                }
                else if (Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    WordTemplatePath = file; // 假设 WordTemplatePath 属性的 setter 会处理通知
                }
                else
                {
                    _dispatcherQueue.TryEnqueue(() => StatusMessage = "不支持的文件类型");
                }
            }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"处理拖放文件时出错: {ex.Message}"); }
        }

        private void CopyToClipboard(string? text) // 命令方法
        {
            if (string.IsNullOrEmpty(text)) return;
            try
            {
                DataPackage dataPackage = new DataPackage(); dataPackage.SetText(text);
                Clipboard.SetContent(dataPackage);
                _dispatcherQueue.TryEnqueue(() => StatusMessage = $"已复制「{text}」到剪贴板");
                var timer = _dispatcherQueue.CreateTimer(); timer.Interval = TimeSpan.FromSeconds(3); timer.IsRepeating = false;
                timer.Tick += (s, e) => _dispatcherQueue.TryEnqueue(() => { if (StatusMessage.StartsWith($"已复制「{text}」")) StatusMessage = "准备就绪"; });
                timer.Start();
            }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"复制到剪贴板失败: {ex.Message}"); }
        }

        public async Task GenerateDocuments() // 命令方法
        {
            if (string.IsNullOrEmpty(OutputDirectory) || !Directory.Exists(OutputDirectory))
            {
                _dispatcherQueue.TryEnqueue(() => StatusMessage = "请选择有效的输出目录");
                return;
            }

            var hasWordTemplate = !string.IsNullOrEmpty(WordTemplatePath) && File.Exists(WordTemplatePath);
            var hasExcelTemplate = UseExcelTemplate && !string.IsNullOrEmpty(ExcelTemplatePath) && File.Exists(ExcelTemplatePath);

            if (!hasWordTemplate && !hasExcelTemplate)
            {
                _dispatcherQueue.TryEnqueue(() => StatusMessage = "请至少选择一个Word模板或Excel模板");
                return;
            }

            List<Dictionary<string, string>> dataRows = IsMultiTableMode ? _multiTableData.MergedRows : _excelData.Rows.ToList();

            if (dataRows == null || !dataRows.Any())
            {
                _dispatcherQueue.TryEnqueue(() => StatusMessage = "没有可处理的数据");
                return;
            }

            _dispatcherQueue.TryEnqueue(() =>
            {
                IsProcessing = true;
                ProgressValue = 0;
                TotalItems = dataRows.Count;
                ProcessedItems = 0;
                ProcessResultText = string.Empty;
            });

            var localSuccessCount = 0;
            var localFailCount = 0;

            try
            {
                await Task.Run(async () =>
                {
                    for (var i = 0; i < dataRows.Count; i++)
                    {
                        if (!IsProcessing) break; // 允许提前中止

                        var rowData = new Dictionary<string, string>(dataRows[i]); // 复制以防修改
                        rowData["序号"] = (i + 1).ToString();
                        rowData["时间"] = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                        rowData["日期"] = DateTime.Now.ToString("yyyy-MM-dd");

                        if (EnableIdCardExtraction && !string.IsNullOrEmpty(SelectedIdCardColumn) &&
                            rowData.TryGetValue(SelectedIdCardColumn, out var idCard) && !string.IsNullOrEmpty(idCard))
                        {
                            try
                            {
                                rowData["身份证性别"] = _idCardService.ExtractGender(idCard);
                            }
                            catch
                            {
                                // 可以选择记录日志或静默处理
                            }
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
                               ? await ExcelTemplateService.ProcessTemplateWithImagesAsync(ExcelTemplatePath, excelOutputPath, rowData, ImageDirectories, ImageFillMode, ImageFillPercentage, null)
                               : await ExcelTemplateService.ProcessTemplateAsync(ExcelTemplatePath, excelOutputPath, rowData, null);
                            if (excelResult.Success) localSuccessCount++; else localFailCount++;
                        }

                        // 更新UI需要在DispatcherQueue上进行
                        _dispatcherQueue.TryEnqueue(() => {
                            ProcessedItems = i + 1;
                            ProgressValue = (int)((double)ProcessedItems / TotalItems * 100);
                        });
                    }
                });

                _dispatcherQueue.TryEnqueue(() =>
                {
                    ProcessResultText = $"处理完成：成功 {localSuccessCount} 个，失败 {localFailCount} 个";
                    ProcessSuccess = localFailCount == 0;
                    StatusMessage = $"文档生成完成，输出到 {OutputDirectory}";
                    ProgressValue = 100; // 确保进度条满
                });
            }
            catch (Exception ex)
            {
                _dispatcherQueue.TryEnqueue(() =>
                {
                    StatusMessage = $"处理过程中出错: {ex.Message}";
                    ProcessSuccess = false;
                    ProcessResultText = $"处理失败: {ex.Message}";
                });
            }
            finally
            {
                _dispatcherQueue.TryEnqueue(() => IsProcessing = false);
            }
        }

        private string GenerateOutputFileName(Dictionary<string, string> rowData, int index)
        {
            var fileName = OutputFileNameTemplate;
            foreach (var item in rowData) fileName = fileName.Replace($"{{{item.Key}}}", item.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            fileName = fileName.Replace("{序号}", rowData.ContainsKey("序号") ? rowData["序号"] : (index + 1).ToString(), StringComparison.OrdinalIgnoreCase);
            fileName = fileName.Replace("{时间}", rowData.ContainsKey("时间") ? rowData["时间"] : DateTime.Now.ToString("yyyyMMdd-HHmmss"), StringComparison.OrdinalIgnoreCase);
            fileName = fileName.Replace("{日期}", rowData.ContainsKey("日期") ? rowData["日期"] : DateTime.Now.ToString("yyyy-MM-dd"), StringComparison.OrdinalIgnoreCase);
            foreach (var invalidChar in Path.GetInvalidFileNameChars()) fileName = fileName.Replace(invalidChar.ToString(), "_");
            return string.IsNullOrWhiteSpace(fileName) || fileName.All(c => c == '_') ? $"Document_{index + 1}_{DateTime.Now:yyyyMMddHHmmss}" : fileName;
        }

        public async Task CheckPlaceholders() // 命令方法
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
                            EnableIdCardExtraction = true; // 通过属性赋值
                            StatusMessage += " (已自动启用身份证信息提取功能)";
                        }
                    }
                    else { StatusMessage = "模板中未找到任何占位符"; }
                });
            }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"检查占位符失败: {ex.Message}"); }
            finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
        }

        private void OpenOutputFolder() // 命令方法
        {
            if (Directory.Exists(OutputDirectory))
                try { Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"\"{OutputDirectory}\"", UseShellExecute = true }); }
                catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"打开输出文件夹失败: {ex.Message}"); }
            else _dispatcherQueue.TryEnqueue(() => StatusMessage = "输出目录不存在");
        }

        // CalculateAge 方法在原始代码中未被使用，如果需要，可以保留或移除
        // private string CalculateAge(string birthDateStr) { return "未知"; }

        public async Task CheckExcelPlaceholders() // 命令方法
        {
            if (string.IsNullOrEmpty(ExcelTemplatePath) || !File.Exists(ExcelTemplatePath)) { _dispatcherQueue.TryEnqueue(() => StatusMessage = "请先选择有效的Excel模板"); return; }
            try
            {
                _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = "正在检查Excel模板中的占位符..."; });
                var placeholders = await ExcelTemplateService.ExtractPlaceholdersAsync(ExcelTemplatePath);
                _dispatcherQueue.TryEnqueue(() => {
                    var newPlaceholders = new ObservableCollection<string>();
                    foreach (var placeholder in placeholders) newPlaceholders.Add(placeholder);
                    DetectedExcelPlaceholders = newPlaceholders; // 通过属性赋值
                    StatusMessage = placeholders.Any() ? $"Excel模板中找到 {placeholders.Count} 个占位符" : "Excel模板中未找到任何占位符";
                });
            }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"检查Excel占位符失败: {ex.Message}"); }
            finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
        }

        private void ClearExcelTemplate() // 命令方法
        {
            ExcelTemplatePath = string.Empty; // 通过属性赋值
            _dispatcherQueue.TryEnqueue(() => DetectedExcelPlaceholders.Clear()); // 直接操作集合，如果需要通知，则应通过属性
            UseExcelTemplate = false; // 通过属性赋值
        }

        private void RemoveImageDirectory(ImageSourceDirectory? directory) // 命令方法
        {
            if (directory != null && ImageDirectories.Remove(directory)) // ImageDirectories 是 ObservableCollection
            {
                _dispatcherQueue.TryEnqueue(() => StatusMessage = $"已移除图片目录: {directory.DirectoryName}");
                if (ImageDirectories.Count == 0) UseImageReplacement = false; // 通过属性赋值
            }
        }

        public async Task CheckDirectoryImages(ImageSourceDirectory? directory) // 命令方法 (参数从View传递)
        {
            if (directory == null || string.IsNullOrEmpty(directory.DirectoryPath) || !Directory.Exists(directory.DirectoryPath)) return;
            try
            {
                _dispatcherQueue.TryEnqueue(() => { IsProcessing = true; StatusMessage = $"正在扫描目录 {directory.DirectoryName} 中的图片..."; });
                var imageFiles = await ImageProcessingService.ScanDirectoryForImagesAsync(directory.DirectoryPath);
                _dispatcherQueue.TryEnqueue(() => {
                    if (directory.ImageFiles == null) directory.ImageFiles = new ObservableCollection<string>();
                    directory.ImageFiles.Clear(); // 直接操作集合
                    foreach (var file in imageFiles) directory.ImageFiles.Add(file);
                    StatusMessage = $"目录 {directory.DirectoryName} 中找到 {imageFiles.Count} 个图片";
                });
            }
            catch (Exception ex) { _dispatcherQueue.TryEnqueue(() => StatusMessage = $"扫描图片目录失败: {ex.Message}"); }
            finally { _dispatcherQueue.TryEnqueue(() => IsProcessing = false); }
        }

        private void ClearAllImageDirectories() // 命令方法
        {
            ImageDirectories.Clear(); // 直接操作集合
            UseImageReplacement = false; // 通过属性赋值
            _dispatcherQueue.TryEnqueue(() => StatusMessage = "已清除所有图片目录");
        }
    }
}