using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Windowing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

using Windows.ApplicationModel.DataTransfer; // 需要这个 using
using Windows.Storage;                     // 需要这个 using
using Windows.ApplicationModel.Resources; // 用于本地化


namespace DocTransform.Pages
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class HomePage : Page
    {
        private static ResourceLoader _resourceLoader = null;

        private static ResourceLoader GetResourceLoaderInstance()
        {
            if (_resourceLoader == null)
            {
                try { _resourceLoader = ResourceLoader.GetForViewIndependentUse(); }
                catch { _resourceLoader = new ResourceLoader(); }
            }
            return _resourceLoader;
        }
        public static string GetLocalizedString(string resourceKey) // 可以设为内部或公共静态
        {
            try { return GetResourceLoaderInstance().GetString(resourceKey); }
            catch (Exception) { return $"!{resourceKey}!"; }
        }

        public HomePage()
        {
            InitializeComponent();
        }

        // 从 MainWindow.xaml.cs 迁移过来的拖放事件处理
        private void RootGrid_DragOver(object sender, DragEventArgs e)
        {
            try
            {
                e.AcceptedOperation = DataPackageOperation.Copy;
                // 使用本地化字符串 (确保您的 .resw 文件中有 "DragFilesCaption")
                e.DragUIOverride.Caption = GetLocalizedString("DragFilesCaption");
                e.DragUIOverride.IsCaptionVisible = true;
                e.DragUIOverride.IsGlyphVisible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HomePage: 拖拽悬停处理失败: {ex.Message}");
            }
        }

        private async void RootGrid_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.DataView.Contains(StandardDataFormats.StorageItems))
                {
                    var items = await e.DataView.GetStorageItemsAsync();
                    if (items.Any())
                    {
                        var filePaths = items.OfType<StorageFile>().Select(i => i.Path).ToArray();
                        System.Diagnostics.Debug.WriteLine($"HomePage: 拖入文件: {string.Join(", ", filePaths)}");

                        // var viewModel = this.DataContext as MainViewModel; // 或者特定的 HomePageViewModel
                        // viewModel?.HandleDroppedFilesCommand?.Execute(filePaths);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HomePage: 文件拖拽处理失败: {ex.Message}");
            }
        }
    }
}
