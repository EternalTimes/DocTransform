using DocTransform.ViewModels;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Linq;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.Storage;
using WinRT.Interop;

namespace DocTransform
{
    public sealed partial class MainWindow : Window
    {
        public MainViewModel ViewModel { get; }
        private AppWindow _appWindow;
        private bool _isMaximized = false;
        private OverlappedPresenter? _presenter;

        public MainWindow()
        {
            this.InitializeComponent();

            // Pass the window's DispatcherQueue to the ViewModel's constructor
            ViewModel = new MainViewModel(this.DispatcherQueue);

            // Initialize window properties
            InitializeWindow();

            // Set up custom title bar
            ExtendsContentIntoTitleBar = true;
            SetTitleBar(AppTitleBar);

            // Add drag and drop handlers
            SetupDragAndDrop();

            // Add window state change handlers
            SetupWindowStateHandlers();
        }

        private void InitializeWindow()
        {
            /*
            // Get AppWindow for advanced window management
            IntPtr hWnd = WindowNative.GetWindowHandle(this);
            WindowId windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            _appWindow = AppWindow.GetFromWindowId(windowId);
            */

            // Get the presenter for window state management
            _presenter = _appWindow.Presenter as OverlappedPresenter;

            // Set window properties
            this.Title = "Excel��Word����ӳ�乤��";
            // ��ȡAppWindow�����ó�ʼ��С
            _appWindow = GetAppWindowForCurrentWindow();
            if (_appWindow != null)
            {
                _appWindow.Resize(new SizeInt32(950, 680)); // ���ó�ʼ��Ⱥ͸߶�
            }

            // Center the window
            CenterWindow();
        }

        private AppWindow GetAppWindowForCurrentWindow()
        {
            IntPtr hWnd = WindowNative.GetWindowHandle(this);
            WindowId wndId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            return AppWindow.GetFromWindowId(wndId);
        }

        private void CenterWindow()
        {
            var displayArea = DisplayArea.GetFromWindowId(_appWindow.Id, DisplayAreaFallback.Nearest);
            if (displayArea is not null)
            {
                var centeredPosition = _appWindow.Position;
                centeredPosition.X = (displayArea.WorkArea.Width - _appWindow.Size.Width) / 2;
                centeredPosition.Y = (displayArea.WorkArea.Height - _appWindow.Size.Height) / 2;
                _appWindow.Move(centeredPosition);
            }
        }

        private void SetupDragAndDrop()
        {
            // Enable drag and drop on the MainContentGrid
            // We'll set this up after the grid is loaded
        }

        private void SetupWindowStateHandlers()
        {
            // We'll track window state manually through button clicks
            // Alternative: Use a timer to periodically check state if needed
        }

        public void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            // Use the presenter to minimize the window
            if (_presenter != null)
            {
                _presenter.Minimize();
            }
        }

        public void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleMaximize();
        }

        private void ToggleMaximize()
        {
            if (_presenter != null)
            {
                if (_presenter.State == OverlappedPresenterState.Maximized)
                {
                    _presenter.Restore();
                    _isMaximized = false;
                }
                else
                {
                    _presenter.Maximize();
                    _isMaximized = true;
                }

            }
        }



        public void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MainContentGrid_Loaded(object sender, RoutedEventArgs e)
        {
            // Set up drag and drop after the grid is loaded
            if (sender is Grid grid)
            {
                grid.AllowDrop = true;
                // Drag & drop events will be added later when ViewModel is ready
            }
        }

        // Method to show ContentDialog (replaces MaterialDesign DialogHost)
        public async void ShowDialog(ContentDialog dialog)
        {
            dialog.XamlRoot = this.Content.XamlRoot;
            await dialog.ShowAsync();
        }

        private void RootGrid_DragOver(object sender, DragEventArgs e)
        {
            // ����ק������ļ�ʱ����ϵͳ�������ǽ��ܸ��Ʋ���
            // ���ʹ���ָ����ʾΪ�����ơ�ͼ��
            e.AcceptedOperation = DataPackageOperation.Copy;
        }

        private async void RootGrid_Drop(object sender, DragEventArgs e)
        {
            // �����ק�����Ƿ�����ļ���StorageItems��
            if (e.DataView.Contains(StandardDataFormats.StorageItems))
            {
                // �첽��ȡ���б�������ļ�
                var items = await e.DataView.GetStorageItemsAsync();
                if (items.Any())
                {
                    // ��ȡ�����ļ�������·��
                    var filePaths = items.OfType<StorageFile>().Select(i => i.Path).ToArray();

                    // ��� ViewModel �� Command ���ڣ���ִ������
                    if (ViewModel?.HandleDroppedFilesCommand != null)
                    {
                        ViewModel.HandleDroppedFilesCommand.Execute(filePaths);
                    }
                }
            }
        }

        // Clean up when needed (if you add any future event handlers)
        public void Cleanup()
        {
            // Reserved for future cleanup if needed
        }
    }
}