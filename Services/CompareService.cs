using ExcelMerger.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using ExcelMerger.Utils;

namespace ExcelMerger.Services
{
    public class CompareService
    {
        /// <summary>
        /// 高层接口：对比两个工作表并直接填充结果数据
        /// </summary>
        public (List<ExcelRowModel> AlignedLeft, List<ExcelRowModel> AlignedRight) CompareSheets(
            ExcelSheetModel leftSheet,
            ExcelSheetModel rightSheet,
            List<int> keyIndices)
        {
            // 保护性编程：确保输入不为 null
            var leftRows = leftSheet?.Rows ?? new List<ExcelRowModel>();
            var rightRows = rightSheet?.Rows ?? new List<ExcelRowModel>();

            // 调用核心 LCS 算法
            return LcsCompare(leftRows, rightRows, keyIndices);
        }

        /// <summary>
        /// 使用 LCS 算法对比两个行列表，并返回对齐后的结果
        /// </summary>
        /// <param name="leftRows">左侧原始行数据</param>
        /// <param name="rightRows">右侧原始行数据</param>
        /// <param name="keyIndices">关键字列索引列表</param>
        /// <returns>对齐后的左侧和右侧列表元组</returns>
        public (List<ExcelRowModel> AlignedLeft, List<ExcelRowModel> AlignedRight) LcsCompare(
        List<ExcelRowModel> leftRows, List<ExcelRowModel> rightRows, List<int> keyIndices)
        {
            // 1. 同时处理 null 保护和占位行过滤。（过滤掉可能的空输入和当前列表中的占位空行，防止重复比对时空行堆积）
            // 使用 ?? 确保即使传入 null 也会变成空列表，避免后续产生 NullReferenceException
            var sourceLeft = (leftRows ?? new List<ExcelRowModel>())
                             .Where(r => r.Status != StatusCodes.Empty)
                             .ToList();

            var sourceRight = (rightRows ?? new List<ExcelRowModel>())
                              .Where(r => r.Status != StatusCodes.Empty)
                              .ToList();

            int m = sourceLeft.Count;
            int n = sourceRight.Count;

            // 0. 构建 DP 矩阵(Space: O(m*n))。dp[i,j] 表示 left 前 i 行和 right 前 j 行的最长公共子序列长度。
            // 即 dp[i,j] 存储 sourceLeft[0..i-1] 和 sourceRight[0..j-1] 的 LCS 长度
            int[,] dp = new int[m + 1, n + 1];

            // 1. 填充矩阵计算最大长度
            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    // 比较两行是否“匹配”（基于关键字或全行）
                    if (AreRowsEqual(sourceLeft[i - 1], sourceRight[j - 1], keyIndices))
                        dp[i, j] = dp[i - 1, j - 1] + 1;
                    else
                        dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                }
            }

            // 2. 回溯（Backtrack）找到具体的增删状态。调用对齐函数并返回结果
            return BuildAlignedLists(dp, sourceLeft, sourceRight, keyIndices);
        }

        /// <summary>
        /// 根据 DP 矩阵回溯，生成带占位符的对齐列表
        /// </summary>
        private (List<ExcelRowModel> AlignedLeft, List<ExcelRowModel> AlignedRight) BuildAlignedLists(
            int[,] dp, List<ExcelRowModel> leftRows, List<ExcelRowModel> rightRows, List<int> keyIndices)
        {
            var alignedLeft = new List<ExcelRowModel>();
            var alignedRight = new List<ExcelRowModel>();

            // 初始化：先把所有行标记为特殊状态，方便后续处理
            // 注意：回溯是从后往前走的
            int i = leftRows.Count;
            int j = rightRows.Count;

            // 获取列数，用于创建空行占位符
            int leftColCount = leftRows.FirstOrDefault()?.Cells.Count ?? 0;
            int rightColCount = rightRows.FirstOrDefault()?.Cells.Count ?? 0;

            // 从右下角回溯到左上角
            while (i > 0 || j > 0)
            {
                if (i > 0 && j > 0 && AreRowsEqual(leftRows[i - 1], rightRows[j - 1], keyIndices))
                {
                    // 路径 1: 行匹配 (LCS 路径)
                    // 情况 A: 匹配项 (相同或修改)
                    var lRow = leftRows[i - 1];
                    var rRow = rightRows[j - 1];

                    // 清除旧的差异记录
                    lRow.DiffColumnIndices = new HashSet<int>();
                    rRow.DiffColumnIndices = new HashSet<int>();
                    
                    bool fullySame = true;
                    int maxCols = Math.Max(lRow.Cells.Count, rRow.Cells.Count);

                    for (int col = 0; col < maxCols; col++)
                    {
                        string valL = col < lRow.Cells.Count ? lRow.Cells[col] : string.Empty;
                        string valR = col < rRow.Cells.Count ? rRow.Cells[col] : string.Empty;

                        if (valL != valR)
                        {
                            fullySame = false;
                            // 记录差异列索引
                            lRow.DiffColumnIndices.Add(col);
                            rRow.DiffColumnIndices.Add(col);
                        }
                    }

                    lRow.Status = rRow.Status = fullySame ? StatusCodes.Same : StatusCodes.Modified;
                    lRow.IsDifferent = rRow.IsDifferent = !fullySame;

                    alignedLeft.Insert(0, lRow);
                    alignedRight.Insert(0, rRow);
                    i--; j--;
                }
                else if (j > 0 && (i == 0 || dp[i, j - 1] >= dp[i - 1, j]))
                {
                    // 路径 2: (向上移动代价更小或已到左边界)
                    // 情况 B: 右侧新增 -> 左侧插空
                    var rRow = rightRows[j - 1];
                    rRow.Status = StatusCodes.Added;
                    rRow.IsDifferent = true;

                    alignedLeft.Insert(0, CreateEmptyRow(leftColCount)); // 左边补白
                    alignedRight.Insert(0, rRow);
                    j--;
                }
                else
                {
                    // 路径 3: 左侧删除
                    // 情况 C: 左侧删除 -> 右侧插空
                    var lRow = leftRows[i - 1];
                    lRow.Status = StatusCodes.Deleted;
                    lRow.IsDifferent = true;

                    alignedLeft.Insert(0, lRow);
                    alignedRight.Insert(0, CreateEmptyRow(rightColCount)); // 右边补白
                    i--;
                }
            }

            return (alignedLeft, alignedRight);
        }

        /// <summary>
        /// 比较两行是否“相等”（基于关键字或全行）
        /// </summary>
        private bool AreRowsEqual(ExcelRowModel left, ExcelRowModel right, List<int> keyIndices)
        {
            if (left == null || right == null) return false;

            // 如果定义了关键字，只比较关键字列
            if (keyIndices != null && keyIndices.Any())
            {
                foreach (int index in keyIndices)
                {
                    var leftVal = index < left.Cells.Count ? left.Cells[index]?.Trim() ?? string.Empty : string.Empty;
                    var rightVal = index < right.Cells.Count ? right.Cells[index]?.Trim() ?? string.Empty : string.Empty;
                    if (leftVal != rightVal) return false;
                }
                return true;
            }
            // 否则比较全行内容
            return left.GetFullContent() == right.GetFullContent();
        }

        /// <summary>
        /// 创建一个带有 "Empty" 状态的占位行，用于对齐显示
        /// </summary>
        private ExcelRowModel CreateEmptyRow(int colCount)
        {
            return new ExcelRowModel
            {
                Cells = Enumerable.Repeat(string.Empty, colCount).ToList(),
                Status = StatusCodes.Empty,
                IsDifferent = false
            };
        }

        /// <summary>
        /// 列模式：将工作表转置为“列模型”列表。每个 ExcelRowModel 代表原始文件的一列，
        /// Cells[r] 即第 r 行该列的值，OriginalIndex 为原始列号（从 1 开始）。
        /// 转置后即可复用 LcsCompare 按列比对（关键字索引对应原始行号-1）。
        /// </summary>
        public List<ExcelRowModel> SheetToColumns(ExcelSheetModel sheet)
        {
            var columns = new List<ExcelRowModel>();
            var rows = sheet?.Rows ?? new List<ExcelRowModel>();
            int colCount = sheet?.MaxColumnCount ?? 0;

            for (int c = 0; c < colCount; c++)
            {
                var column = new ExcelRowModel { OriginalIndex = c + 1 };
                for (int r = 0; r < rows.Count; r++)
                {
                    column.Cells.Add(c < rows[r].Cells.Count ? rows[r].Cells[c] : string.Empty);
                }
                columns.Add(column);
            }
            return columns;
        }

        /// <summary>
        /// 列模式：将对齐后的列模型转置回“显示行”，恢复原始行列方向用于 DataGrid 显示。
        /// 同时把每列的状态写入行的 ColumnStatusMap（整列高亮），
        /// 并把修改列中发生变化的单元格（行索引）记入 DiffColumnIndices（文字变红加粗）。
        /// </summary>
        public List<ExcelRowModel> ColumnsToRows(List<ExcelRowModel> alignedColumns)
        {
            var displayRows = new List<ExcelRowModel>();

            // 显示行数 = 该侧文件原始行数（占位列的 Cells 数即为本侧行数）
            int rowCount = 0;
            foreach (var col in alignedColumns)
            {
                rowCount = Math.Max(rowCount, col.Cells.Count);
            }

            for (int r = 0; r < rowCount; r++)
            {
                var row = new ExcelRowModel { OriginalIndex = r + 1 };
                for (int c = 0; c < alignedColumns.Count; c++)
                {
                    var col = alignedColumns[c];
                    row.Cells.Add(r < col.Cells.Count ? col.Cells[r] : string.Empty);
                    row.ColumnStatusMap[c] = col.Status ?? StatusCodes.Same;

                    if (col.Status == StatusCodes.Modified && col.DiffColumnIndices.Contains(r))
                    {
                        row.DiffColumnIndices.Add(c);
                    }
                }
                displayRows.Add(row);
            }
            return displayRows;
        }

    }
}
