using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerger.Models;
using System;
using System.Collections.Generic;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace ExcelMerger.Services
{
    public class ExcelService
    {
        // 使用 FileStream 以允许其他进程（如 Excel）同时打开文件
        public List<ExcelSheetModel> ReadExcelContent(string filePath)
        {
            var sheets = new List<ExcelSheetModel>();

            try
            {
                // Using 确保文件读取后正确关闭资源
                // 1. 使用 FileStream 以共享模式打开文件
                // FileMode.Open: 打开现有文件
                // FileAccess.Read: 仅请求读取权限
                // FileShare.ReadWrite: 允许其他进程同时进行读写操作（这是解决 Excel 占用问题的关键）
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 2. 将文件流传递给 XLWorkbook
                    using (var workbook = new XLWorkbook(stream))
                    {
                        // 3. 遍历所有工作表 (Sheet)
                        foreach (var worksheet in workbook.Worksheets)
                        {
                            var sheetModel = new ExcelSheetModel { SheetName = worksheet.Name };
                            var range = worksheet.RangeUsed();

                            if (range != null)
                            {
                                // 找到当前工作表中最右侧有数据的列号
                                int maxColumnInSheet = range.LastColumn().ColumnNumber();
                                sheetModel.MaxColumnCount = maxColumnInSheet;

                                // 找到该表数据范围的最后一列，确保每一行读取的列数一致
                                // int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

                                // 遍历该表中有数据的所有行
                                var rowOriginalIndex = 1;
                                foreach (var row in worksheet.RowsUsed())
                                {
                                    var rowModel = new ExcelRowModel();
                                    rowModel.OriginalIndex = rowOriginalIndex++;

                                    // 遍历从第 1 列到最后一列的所有单元格
                                    for (int i = 1; i <= maxColumnInSheet; i++)
                                    {
                                        // 如果某一行比其他行短，GetString() 依然能安全处理
                                        string cellValue = row.Cell(i).GetString().Trim();
                                        rowModel.Cells.Add(cellValue);
                                    }

                                    sheetModel.Rows.Add(rowModel);
                                }
                            }

                            sheets.Add(sheetModel);
                        }
                    }
                            

                }
            }
            catch (System.IO.IOException ex)
            {
                // 如果依然发生错误（例如文件被极其严格的权限锁定），则提示用户
                MessageBox.Show($"读取文件失败，可能文件被锁定。详情: {ex.Message}", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生意外错误: {ex.Message}", "错误");
            }

            return sheets;
        }
    }

}
