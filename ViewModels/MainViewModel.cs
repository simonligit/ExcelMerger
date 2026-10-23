using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using ExcelMerger.Messages;
using ExcelMerger.Models;
using ExcelMerger.Services;
using ExcelMerger.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerger.ViewModels
{
   public partial class MainViewModel : ObservableObject
    {
        private readonly ExcelService _excelService = new();
        private readonly CompareService _compareService = new();

        [ObservableProperty]
        private bool _isBusy; // 当为 true 时，显示进度条

        [ObservableProperty]
        private bool _hasData = false; // 初始状态为 false

        [ObservableProperty]
        private int _selectedSheetIndex;
        
        // 状态信息，可以绑定到底部的 StatusBar
        [ObservableProperty]
        private string _statusMessage;

        // ObservableCollection 是 WPF 中最常用的集合，它在添加/删除项时会自动通知 UI 更新
        [ObservableProperty]
        private ObservableCollection<SheetComparisonViewModel> _sheets = new();

        // NEW: expose full file paths used for comparison (shown at top-level)
        [ObservableProperty]
        private string _leftFilePath = string.Empty;

        [ObservableProperty]
        private string _rightFilePath = string.Empty;

        // 对比模式：false = 行模式（左右并排比对行），true = 列模式（上下堆叠比对列）
        [ObservableProperty]
        private bool _isColumnMode;

        // 缓存已解析的工作表数据，切换对比模式时无需重新读取文件（也避免再次弹出密码框）
        private List<ExcelSheetModel> _cachedLeftSheets = new();
        private List<ExcelSheetModel> _cachedRightSheets = new();

        public MainViewModel()
        {
            // initialize with localized default message
            _statusMessage = LocalizationHelper.Get("StatusSelectFiles");
        }

        partial void OnIsColumnModeChanged(bool value)
        {
            // 模式切换后基于缓存数据重建比对结果
            if (_cachedLeftSheets.Count > 0 || _cachedRightSheets.Count > 0)
            {
                _ = RebuildComparisonAsync();
            }
        }

        public async Task LoadAndCompareAsync(string path1, string path2)
        {
            try
            {
                // set the full paths immediately so UI can show them while processing
                LeftFilePath = path1 ?? string.Empty;
                RightFilePath = path2 ?? string.Empty;

                IsBusy = true;
                HasData = false; // 开始处理时先重置状态
                StatusMessage = LocalizationHelper.Get("StatusProcessing");

                var (leftSheets, rightSheets) = await Task.Run(() => (
                    _excelService.ReadExcelContent(path1),
                    _excelService.ReadExcelContent(path2)
                ));

                _cachedLeftSheets = leftSheets;
                _cachedRightSheets = rightSheets;

                await BuildComparisonAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = LocalizationHelper.GetFormat("ErrorFormat", ex.Message);
            }
            finally
            {
                IsBusy = false; // 无论成功失败，最后都要关闭加载状态
            }
        }

        // 模式切换时基于缓存数据重建（不重新读取文件）
        private async Task RebuildComparisonAsync()
        {
            try
            {
                IsBusy = true;
                StatusMessage = LocalizationHelper.Get("StatusProcessing");

                await BuildComparisonAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = LocalizationHelper.GetFormat("ErrorFormat", ex.Message);
            }
            finally
            {
                IsBusy = false;
            }
        }

        // 根据缓存的工作表数据与当前模式构建比对结果并发布到 UI
        private async Task BuildComparisonAsync()
        {
            var leftSheets = _cachedLeftSheets;
            var rightSheets = _cachedRightSheets;
            var tempSheets = new List<SheetComparisonViewModel>();

            // 对比，假设按索引顺序一一对应比对。Task.Run 会开启一个新的后台线程
            await Task.Run(() =>
            {
                for (int i = 0; i < leftSheets.Count; i++)
                {
                    var leftSheet = leftSheets[i];
                    // 寻找右侧同名工作表，找不到则给个空对象
                    var rightSheet = rightSheets.FirstOrDefault(s => s.SheetName == leftSheet.SheetName)
                                     ?? new ExcelSheetModel { Rows = new List<ExcelRowModel>(), SheetName = leftSheet.SheetName };

                    var sheetVM = IsColumnMode
                        ? BuildColumnModeSheet(leftSheet, rightSheet)
                        : BuildRowModeSheet(leftSheet, rightSheet);

                    tempSheets.Add(sheetVM);
                }

                // 更新 UI 数据（必须回到主线程或重新实例化集合）
                App.Current.Dispatcher.Invoke(() =>
                {
                    Sheets = new ObservableCollection<SheetComparisonViewModel>(tempSheets);
                    // 重新加载后强制选中第一个 Tab
                    if (Sheets.Count > 0)
                    {
                        SelectedSheetIndex = 0;
                    }
                    // 数据填充完毕后，开启显示开关
                    HasData = Sheets.Count > 0;
                });
            });

            StatusMessage = LocalizationHelper.Get("StatusCompleted");

            HasData = Sheets.Count > 0;
        }

        // 行模式：左右并排，按行 LCS 对齐，关键字为选定的列
        private SheetComparisonViewModel BuildRowModeSheet(ExcelSheetModel leftSheet, ExcelSheetModel rightSheet)
        {
            // 计算当前 Sheet 组合需要的最大列数
            // 这样 DataGrid 就能知道最少需要生成多少列来显示所有内容
            int currentMax = Math.Max(leftSheet.MaxColumnCount, rightSheet.MaxColumnCount);

            // 默认将第一列设为 Key (索引 0)
            var defaultKeys = new List<int> { 0 };

            var sheetVM = new SheetComparisonViewModel
            {
                SheetName = leftSheet.SheetName,
                IsColumnMode = false,
                MaxColumns = currentMax,
                KeyIndices = defaultKeys
            };

            // 初始化列配置（默认勾选第一列）
            for (int colIndex = 0; colIndex < currentMax; colIndex++)
            {
                sheetVM.ColumnHeaderList.Add(new ColumnConfig
                {
                    HeaderName = LocalizationHelper.GetFormat("ColumnHeaderFormat", colIndex + 1),
                    Index = colIndex,
                    // 默认勾选第一列作为关键字列
                    IsKeyColumn = (colIndex == 0)
                });
            }

            // 使用 LCS 比对算法
            var (alignedLeft, alignedRight) = _compareService.LcsCompare(leftSheet.Rows, rightSheet.Rows, defaultKeys);

            sheetVM.LeftRows = new ObservableCollection<ExcelRowModel>(alignedLeft);
            sheetVM.RightRows = new ObservableCollection<ExcelRowModel>(alignedRight);

            return sheetVM;
        }

        // 列模式：上下堆叠，按列 LCS 对齐（复用转置 + LCS），关键字为选定的行（默认第一行/表头）
        private SheetComparisonViewModel BuildColumnModeSheet(ExcelSheetModel leftSheet, ExcelSheetModel rightSheet)
        {
            // 转置：每个原始列变成一个“列模型”，之后按列做 LCS 对齐
            var leftColumns = _compareService.SheetToColumns(leftSheet);
            var rightColumns = _compareService.SheetToColumns(rightSheet);

            int maxRowCount = Math.Max(leftSheet.Rows.Count, rightSheet.Rows.Count);

            // 默认将第一行（通常是表头）设为 Key (索引 0)
            var defaultKeys = new List<int> { 0 };

            var sheetVM = new SheetComparisonViewModel
            {
                SheetName = leftSheet.SheetName,
                IsColumnMode = true,
                KeyIndices = defaultKeys,
                LeftColumnModels = leftColumns,
                RightColumnModels = rightColumns
            };

            // 初始化关键行配置（默认勾选第一行）
            for (int rowIndex = 0; rowIndex < maxRowCount; rowIndex++)
            {
                sheetVM.ColumnHeaderList.Add(new ColumnConfig
                {
                    HeaderName = LocalizationHelper.GetFormat("RowHeaderFormat", rowIndex + 1),
                    Index = rowIndex,
                    // 默认勾选第一行作为关键行
                    IsKeyColumn = (rowIndex == 0)
                });
            }

            // 按列执行 LCS 比对
            var (alignedLeftCols, alignedRightCols) = _compareService.LcsCompare(leftColumns, rightColumns, defaultKeys);

            ApplyColumnModeResult(sheetVM, alignedLeftCols, alignedRightCols);

            return sheetVM;
        }

        // 将对齐后的列模型应用到 sheetVM：转置回显示行、生成列状态与表头
        private void ApplyColumnModeResult(
            SheetComparisonViewModel sheetVM,
            List<ExcelRowModel> alignedLeftCols,
            List<ExcelRowModel> alignedRightCols)
        {
            sheetVM.MaxColumns = alignedLeftCols.Count;
            sheetVM.ColumnHeaderLabels = BuildColumnHeaderLabels(alignedLeftCols, alignedRightCols);
            FillColumnStatuses(sheetVM, alignedLeftCols);

            sheetVM.LeftRows = new ObservableCollection<ExcelRowModel>(_compareService.ColumnsToRows(alignedLeftCols));
            sheetVM.RightRows = new ObservableCollection<ExcelRowModel>(_compareService.ColumnsToRows(alignedRightCols));
        }

        // 生成对齐列的表头：优先使用非占位一侧的原始列号
        private static List<string> BuildColumnHeaderLabels(
            List<ExcelRowModel> alignedLeft, List<ExcelRowModel> alignedRight)
        {
            var labels = new List<string>();
            for (int c = 0; c < alignedLeft.Count; c++)
            {
                var col = alignedLeft[c].Status != StatusCodes.Empty ? alignedLeft[c] : alignedRight[c];
                labels.Add(LocalizationHelper.GetFormat("ColumnHeaderFormat", col.OriginalIndex));
            }
            return labels;
        }

        private static void FillColumnStatuses(SheetComparisonViewModel sheetVM, List<ExcelRowModel> alignedColumns)
        {
            sheetVM.ColumnStatuses.Clear();
            foreach (var col in alignedColumns)
            {
                sheetVM.ColumnStatuses.Add(new ColumnStatusModel { Status = col.Status });
            }
        }

        [RelayCommand]
        private async Task ReCompareAsync(SheetComparisonViewModel sheetVM)
        {
            if (sheetVM == null) return;

            try
            {
                // 1. 开启进度条
                IsBusy = true;
                StatusMessage = LocalizationHelper.GetFormat("RecompareProcessingFormat", sheetVM.SheetName);

                // 2. 提取关键字索引（在 UI 线程准备好数据）
                var selectedKeys = sheetVM.ColumnHeaderList
                    .Where(c => c.IsKeyColumn)
                    .Select(c => c.Index)
                    .ToList();

                // 3. 在后台线程执行比对算法
                await Task.Run(() =>
                {
                    if (sheetVM.IsColumnMode)
                    {
                        // 列模式：基于缓存的原始列模型按选定关键行重新比对
                        var (alignedLeftCols, alignedRightCols) = _compareService.LcsCompare(
                                    sheetVM.LeftColumnModels,
                                    sheetVM.RightColumnModels,
                                    selectedKeys);

                        var newLabels = BuildColumnHeaderLabels(alignedLeftCols, alignedRightCols);
                        var newLeftRows = _compareService.ColumnsToRows(alignedLeftCols);
                        var newRightRows = _compareService.ColumnsToRows(alignedRightCols);

                        App.Current.Dispatcher.Invoke(() =>
                        {
                            sheetVM.MaxColumns = alignedLeftCols.Count;
                            sheetVM.ColumnHeaderLabels = newLabels;

                            sheetVM.LeftRows.Clear();
                            foreach (var row in newLeftRows)
                            {
                                sheetVM.LeftRows.Add(row);
                            }

                            sheetVM.RightRows.Clear();
                            foreach (var row in newRightRows)
                            {
                                sheetVM.RightRows.Add(row);
                            }

                            FillColumnStatuses(sheetVM, alignedLeftCols);

                            // 对齐列数/表头可能已变化，而集合是原地更新的，
                            // 需通知视图重建当前 Sheet 的 DataGrid 列
                            WeakReferenceMessenger.Default.Send(new SheetRebuiltMessage(sheetVM));
                        });
                    }
                    else
                    {
                        // 行模式：使用 LCS 比对算法
                        // 这里传入的 sheetVM 里的 Rows 已经包含用于对齐显示的的空行，
                        // 但 LcsCompare 内部已经处理了 Status != "Empty" 的过滤。
                        var (alignedLeft, alignedRight) = _compareService.LcsCompare(
                                    sheetVM.LeftRows.ToList(),
                                    sheetVM.RightRows.ToList(),
                                    selectedKeys);

                        // 在主线程更新 UI 绑定的集合
                        App.Current.Dispatcher.Invoke(() =>
                        {
                            // 不要 new 集合，而是操作现有集合，以保持数据绑定有效。暂不优化。
                            //sheetVM.LeftRows = new ObservableCollection<ExcelRowModel>(alignedLeft);
                            //sheetVM.RightRows = new ObservableCollection<ExcelRowModel>(alignedRight);
                            sheetVM.LeftRows.Clear();
                            foreach (var row in alignedLeft)
                            {
                                sheetVM.LeftRows.Add(row);
                            }

                            sheetVM.RightRows.Clear();
                            foreach (var row in alignedRight)
                            {
                                sheetVM.RightRows.Add(row);
                            }
                        });
                    }
                });

                StatusMessage = LocalizationHelper.GetFormat("RecompareCompletedFormat", sheetVM.SheetName);
            }
            catch (Exception ex)
            {
                StatusMessage = LocalizationHelper.GetFormat("CompareFailedFormat", ex.Message);
            }
            finally
            {
                // 4. 关闭进度条
                IsBusy = false;
            }
        }
        [RelayCommand]
        private void OpenSelectFiles()
        {
            // 使用完整路径来引用 Application
            var mainWindow = System.Windows.Application.Current.MainWindow;

            // 弹出选择文件对话框
            var dialog = new SelectFilesDialog();
            dialog.Owner = mainWindow;
            dialog.ShowDialog();
        }
    }
}
