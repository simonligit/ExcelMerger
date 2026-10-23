using CommunityToolkit.Mvvm.Messaging;
using ExcelMerger.Converters;
using ExcelMerger.Messages;
using ExcelMerger.Models;
using ExcelMerger.ViewModels;
using ExcelMerger.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;
        public MainWindow()
        {
            InitializeComponent();

            // 实例化 ViewModel 并设置为数据上下文
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;
            
            // 注册订阅者
            // 参数1：接收者（this，即当前窗口）
            // 参数2：当收到消息时执行的操作
            WeakReferenceMessenger.Default.Register<FilePathsSelectedMessage>(this, async (r, m) =>
            {
                /* 这是一个测试代码段，用于验证消息传递是否正确
                // m.Value 就是我们传递的 string[] 数组
                string file1 = m.Value[0];
                string file2 = m.Value[1];

                // 弹出对话框进行验证
                MessageBox.Show($"接收到文件路径：\n文件1: {file1}\n文件2: {file2}", "验证成功");
                */
                // 使用 await 异步等待加载和对比任务完成
                // 这样 UI 线程在等待期间依然可以处理 ProgressBar 的动画渲染
                try
                {
                    // 调用异步加载
                    await _viewModel.LoadAndCompareAsync(m.Value[0], m.Value[1]);

                    // 成功后自动切换到第一个 Sheet
                    if (MainTabControl.Items.Count > 0)
                    {
                        MainTabControl.SelectedIndex = 0;
                    }
                }
                catch (Exception ex)
                {
                    // 捕获其他未知错误
                    MessageBox.Show(LocalizationHelper.GetFormat("UnexpectedErrorFormat", ex.Message),
                                    LocalizationHelper.Get("ErrorTitle"),
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                }
            });

            // 列模式下重新比对完成后，对齐列数/表头可能已变化，
            // 而绑定集合是原地更新的（不会重新触发生成列），需主动重建 DataGrid 列
            WeakReferenceMessenger.Default.Register<SheetRebuiltMessage>(this, (r, m) =>
            {
                RebuildColumnsForSelectedTab(m.Value);
            });
        }
        private void MenuItem_New_Click(object sender, RoutedEventArgs e)
        {
            // 实例化并弹出对话框
            var dialog = new SelectFilesDialog();
            dialog.ShowDialog();
        }

        private void MenuItem_Exit_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();

        private void MenuItem_About_Click(object sender, RoutedEventArgs e)
        {
            var aboutWin = new AboutWindow();
            aboutWin.Owner = this; // 设置 Owner 使其在主窗口中心弹出
            aboutWin.ShowDialog();
        }
        private void DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            // 1. 拦截并取消不需要显示的辅助属性列。因为 DataGrid 默认会根据 ExcelRowModel 中的所有公开属性自动生成列。
            if (e.PropertyName == "DiffColumnIndices" ||
                e.PropertyName == "ColumnStatusMap" ||
                e.PropertyName == "OriginalIndex" ||
                e.PropertyName == "Status" ||
                e.PropertyName == "IsDifferent")
            {
                // 暂时隐藏这些辅助状态列，让界面只显示 Excel 内容
                e.Cancel = true;
                return;
            }

            // 2. 核心逻辑：处理 Cells 属性
            if (e.PropertyName == "Cells" && sender is DataGrid dg)
            {
                // 取消对 Cells 列表本身的默认生成（默认只会生成一个长度列）
                e.Cancel = true;

                BuildColumnsForGrid(dg);
            }
        }

        // 为指定 DataGrid 重建所有内容列（首次生成、切换 Tab、列模式重新比对后调用）
        private void BuildColumnsForGrid(DataGrid dg)
        {
            // 获取当前 DataGrid 所在 Tab 页的数据模型
            if (dg == null || dg.DataContext is not SheetComparisonViewModel viewModel) return;

            // 在生成新列前，先清空现有的列集合
            // 这样可以确保每次切换 Tab 时，列数都能根据当前 Sheet 的 MaxColumns 重新构建
            dg.Columns.Clear();

            if (viewModel.IsColumnMode)
            {
                BuildColumnModeColumns(dg, viewModel);
            }
            else
            {
                BuildRowModeColumns(dg, viewModel);
            }
        }

        // 行模式：按顺序生成列，差异单元格文字变红加粗
        private void BuildRowModeColumns(DataGrid dg, SheetComparisonViewModel viewModel)
        {
            var cellDiffConverter = new CellDiffConverter(); // 实例化转换器

            // 使用我们封装好的 MaxColumns
            for (int i = 0; i < viewModel.MaxColumns; i++)
            {
                // 创建一个新列
                var column = new DataGridTextColumn
                {
                    // 设置表头，例如 "Col 1", "Col 2" 或直接用字母 A, B, C
                    Header = LocalizationHelper.GetFormat("ColumnHeaderFormat", i + 1),
                    // 高级绑定技巧：绑定到 Cells 列表的特定索引
                    Binding = new Binding($"Cells[{i}]")
                };

                // 设置单元格样式：如果是差异列，文字变红并加粗
                var cellStyle = new Style(typeof(DataGridCell));

                // 1. 定义公共的 MultiBinding 结构
                Func<string, MultiBinding> createMultiBinding = (param) => {
                    // 多重绑定：需要行对象和当前列索引
                    var mb = new MultiBinding { Converter = cellDiffConverter, ConverterParameter = param };
                    mb.Bindings.Add(new Binding(".")); // 绑定到 RowModel
                    mb.Bindings.Add(new Binding { Source = i }); // 传入当前列索引 i
                    return mb;
                };

                // 2. 绑定前景色 (文字变红)
                cellStyle.Setters.Add(new Setter(DataGridCell.ForegroundProperty, createMultiBinding("Foreground")));

                // 3. 绑定粗细 (文字加粗)
                cellStyle.Setters.Add(new Setter(DataGridCell.FontWeightProperty, createMultiBinding("Weight")));

                column.CellStyle = cellStyle;

                dg.Columns.Add(column);
            }
        }

        // 列模式：表头使用原始列号，整列按对齐状态着色，修改列中的差异单元格文字变红加粗
        private void BuildColumnModeColumns(DataGrid dg, SheetComparisonViewModel viewModel)
        {
            var columnCellConverter = new ColumnCellConverter();

            for (int i = 0; i < viewModel.MaxColumns; i++)
            {
                var column = new DataGridTextColumn
                {
                    // 表头优先使用构建时记录的原始列号（如 "Col 3"），否则回退到顺序编号
                    Header = (viewModel.ColumnHeaderLabels != null && i < viewModel.ColumnHeaderLabels.Count)
                        ? viewModel.ColumnHeaderLabels[i]
                        : LocalizationHelper.GetFormat("ColumnHeaderFormat", i + 1),
                    Binding = new Binding($"Cells[{i}]")
                };

                var cellStyle = new Style(typeof(DataGridCell));

                Func<string, MultiBinding> createMultiBinding = (param) => {
                    var mb = new MultiBinding { Converter = columnCellConverter, ConverterParameter = param };
                    mb.Bindings.Add(new Binding(".")); // 绑定到 RowModel（含 ColumnStatusMap）
                    mb.Bindings.Add(new Binding { Source = i }); // 传入当前列索引 i
                    return mb;
                };

                // 整列背景：新增/删除/修改/占位 列分别着色
                cellStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, createMultiBinding("Background")));
                // 差异单元格：文字变红并加粗
                cellStyle.Setters.Add(new Setter(DataGridCell.ForegroundProperty, createMultiBinding("Foreground")));
                cellStyle.Setters.Add(new Setter(DataGridCell.FontWeightProperty, createMultiBinding("Weight")));

                column.CellStyle = cellStyle;

                dg.Columns.Add(column);
            }
        }

        private void DataGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // 获取当前触发滚动的 DataGrid
            var sourceDataGrid = sender as DataGrid;
            if (sourceDataGrid == null) return;

            // 获取父级 Grid，以便找到另一个 DataGrid
            var parentGrid = sourceDataGrid.Parent as Grid;
            if (parentGrid == null) return;

            // 确定目标 DataGrid（如果是左边在动，就同步给右边，反之亦然）
            DataGrid targetDataGrid = null;
            foreach (var child in parentGrid.Children)
            {
                if (child is DataGrid dg && dg != sourceDataGrid)
                {
                    targetDataGrid = dg;
                    break;
                }
            }

            if (targetDataGrid == null) return;

            // 核心：同步滚动偏移量
            // 使用 GetVisualChild 辅助函数获取内部的 ScrollViewer
            var sourceScrollViewer = GetVisualChild<ScrollViewer>(sourceDataGrid);
            var targetScrollViewer = GetVisualChild<ScrollViewer>(targetDataGrid);

            if (sourceScrollViewer != null && targetScrollViewer != null)
            {
                // 同步垂直滚动（如果需要水平滚动同步，也可以加上 HorizontalOffset）
                targetScrollViewer.ScrollToVerticalOffset(sourceScrollViewer.VerticalOffset);
                targetScrollViewer.ScrollToHorizontalOffset(sourceScrollViewer.HorizontalOffset);

                // 同步更新概览条的阴影块位置
                UpdateMinimapIndicator(sourceScrollViewer);
            }
        }

        // 辅助函数：在控件的视觉树中查找指定类型的子控件
        private T GetVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child != null && child is T t) return t;

                var childOfChild = GetVisualChild<T>(child);
                if (childOfChild != null) return childOfChild;
            }
            return null;
        }

        // 辅助函数：在控件的视觉树中查找所有指定类型的子控件
        private List<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
        {
            var result = new List<T>();
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) result.Add(t);
                result.AddRange(FindVisualChildren<T>(child));
            }
            return result;
        }

        // 重建当前 Tab 内所有 DataGrid 的列（列模式重新比对后调用）
        private void RebuildColumnsForSelectedTab(SheetComparisonViewModel sheetVM)
        {
            if (!ReferenceEquals(MainTabControl.SelectedItem, sheetVM)) return;

            foreach (var dg in FindVisualChildren<DataGrid>(MainTabControl))
            {
                BuildColumnsForGrid(dg);
            }
        }

        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            // 强制刷新 Header 避免虚拟化复用导致的行号错乱
            var rowModel = e.Row.Item as ExcelRowModel;
            if (rowModel != null && rowModel.Status != StatusCodes.Empty)
            {
                // 显示行号（从 1 开始）
                e.Row.Header = rowModel.OriginalIndex.ToString(); 
            }
            else
            {
                // 对齐用的空行不显示行号
                e.Row.Header = string.Empty;
            }
        }
        private void Minimap_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var minimapGrid = sender as FrameworkElement;
            // 获取当前选中的 Sheet
            var sheetVm = MainTabControl.SelectedItem as SheetComparisonViewModel;
            if (sheetVm == null || sheetVm.LeftRows.Count == 0) return;

            // 1. 计算点击位置占总高度的百分比
            Point clickPoint = e.GetPosition(minimapGrid);
            double percent = clickPoint.Y / minimapGrid.ActualHeight;

            // 2. 计算目标行索引
            int targetIndex = (int)(percent * sheetVm.LeftRows.Count);

            var dataGrid = GetVisualChild<DataGrid>(MainTabControl); // 查找当前激活的 DataGrid

            if (dataGrid != null && targetIndex >= 0 && targetIndex < dataGrid.Items.Count)
            {
                dataGrid.ScrollIntoView(dataGrid.Items[targetIndex]);
            }
        }

        // 水平 Minimap（列模式）：点击定位到对应列
        private void MinimapHorizontal_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var minimapGrid = sender as FrameworkElement;
            // 获取当前选中的 Sheet
            var sheetVm = MainTabControl.SelectedItem as SheetComparisonViewModel;
            if (sheetVm == null || sheetVm.ColumnStatuses.Count == 0) return;

            // 1. 计算点击位置占总宽度的百分比
            Point clickPoint = e.GetPosition(minimapGrid);
            double percent = clickPoint.X / minimapGrid.ActualWidth;

            // 2. 计算目标列索引
            int targetIndex = (int)(percent * sheetVm.ColumnStatuses.Count);

            var dataGrid = GetVisualChild<DataGrid>(MainTabControl); // 查找当前激活的 DataGrid
            if (dataGrid == null || targetIndex < 0 || targetIndex >= dataGrid.Columns.Count) return;

            // 3. 按目标列之前的列宽累加估算偏移量，水平滚动到目标列
            double targetOffset = dataGrid.Columns.Take(targetIndex).Sum(c => c.ActualWidth);
            var scrollViewer = GetVisualChild<ScrollViewer>(dataGrid);
            scrollViewer?.ScrollToHorizontalOffset(targetOffset);
        }

        private void UpdateMinimapIndicator(ScrollViewer sv)
        {
            if (sv == null) return;

            // 垂直 Minimap（行模式）：按垂直滚动更新阴影块
            if (sv.ExtentHeight > 0)
            {
                double canvasHeight = MinimapCanvas.ActualHeight;
                double viewPortFactor = sv.ViewportHeight / sv.ExtentHeight;
                double offsetFactor = sv.VerticalOffset / sv.ExtentHeight;
                ViewportIndicator.Height = canvasHeight * viewPortFactor;
                Canvas.SetTop(ViewportIndicator, canvasHeight * offsetFactor);
            }

            // 水平 Minimap（列模式）：按水平滚动更新阴影块
            if (sv.ExtentWidth > 0)
            {
                double canvasWidth = MinimapCanvasH.ActualWidth;
                double viewPortFactorW = sv.ViewportWidth / sv.ExtentWidth;
                double offsetFactorW = sv.HorizontalOffset / sv.ExtentWidth;
                ViewportIndicatorH.Width = Math.Max(5, canvasWidth * viewPortFactorW);
                Canvas.SetLeft(ViewportIndicatorH, canvasWidth * offsetFactorW);
            }
        }
        private async void Window_Drop(object sender, DragEventArgs e)
        {
            // 1. 检查拖入的是否是文件
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 2. 获取文件路径列表
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files.Length == 2)
                {
                    // 3. 如果拖入的是两个文件，直接比较
                    if (IsExcelFile(files[0]) && IsExcelFile(files[1])) 
                    {
                        // 直接调用 ViewModel 的比对方法。这与从 SelectFilesDialog 返回后的逻辑一致
                        await _viewModel.LoadAndCompareAsync(files[0], files[1]);
                        if (MainTabControl.Items.Count > 0) MainTabControl.SelectedIndex = 0;
                    }
                    else
                    {
                        MessageBox.Show(LocalizationHelper.Get("DragDrop_InvalidTwo"),
                                        LocalizationHelper.Get("FormatErrorTitle"));
                    }
                }
                else
                {
                    // 4. 如果拖入的是一个文件，弹出对话框并预设第一个文件
                    if (IsExcelFile(files[0]))
                    {
                        string filePath = files[0];

                        // 使用 Dispatcher 异步调用，确保 Drop 事件先完成
                        // 这会释放文件管理器的拖拽状态
                        _ = Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var dialog = new SelectFilesDialog();
                            dialog.SetFile1(filePath);
                            dialog.Owner = this;
                            dialog.ShowDialog();
                        }));
                    }
                    else
                    {
                        MessageBox.Show(LocalizationHelper.Get("DragDrop_PleaseDropExcel"),
                                        LocalizationHelper.Get("FormatErrorTitle"));
                    }
                }
            }
        }
        private bool IsExcelFile(string path)
        {
            // 验证文件扩展名（支持 xlsx 和 xls）
            string ext = System.IO.Path.GetExtension(path).ToLower();
            return ext == ".xlsx" || ext == ".xls";
        }
        private void Window_DragOver(object sender, DragEventArgs e)
        {
            // 如果拖入的是文件，显示“拷贝”图标，否则显示“禁止”图标
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true; // 表示事件已处理
        }

        private void SwitchToEnglish_Click(object sender, RoutedEventArgs e)
        {
            UpdateLanguage("Resources/StringResources.en.xaml");
        }

        private void SwitchToChinese_Click(object sender, RoutedEventArgs e)
        {
            UpdateLanguage("Resources/StringResources.zh-CN.xaml");
        }

        private void UpdateLanguage(string resourcePath)
        {
            try
            {
                // 1. 创建新资源字典
                ResourceDictionary newDict = new ResourceDictionary
                {
                    Source = new Uri(resourcePath, UriKind.Relative)
                };

                // 2. 查找并移除现有的语言字典（假设它是 MergedDictionaries 的第一个或特定项）
                // 简单做法：清空并重新添加。复杂做法：按路径查找特定字典替换。
                var mergedDicts = Application.Current.Resources.MergedDictionaries;

                // 这种方式会立即刷新所有使用了 {DynamicResource} 的 UI 元素
                mergedDicts.Clear();
                mergedDicts.Add(newDict);

                // 如果你还有其他基础资源字典（如样式、图标），记得在这里重新加回来
                // mergedDicts.Insert(0, new ResourceDictionary { Source = ... });

                // 保存设置
                Properties.Settings.Default.LastLanguage = resourcePath;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(LocalizationHelper.GetFormat("SwitchLanguageFailedFormat", ex.Message),
                                LocalizationHelper.Get("ErrorTitle"));
            }
        }
    }
}