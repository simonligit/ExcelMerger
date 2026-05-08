using ExcelMerger.Models;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.Crypt.Standard;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
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
            IWorkbook workbook = null;
            string password = null;

            try
            {
                // 第一次尝试：无密码打开
                workbook = OpenWorkbook(filePath, null);
            }
            catch (NPOI.EncryptedDocumentException)
            {
                // 获取纯文件名（含路径）
                string fileName = System.IO.Path.GetFullPath(filePath);

                Application.Current.Dispatcher.Invoke(() =>
                {
                    // 传递文件名给构造函数
                    var pwdDlg = new PasswordDialog(fileName) { Owner = Application.Current.MainWindow };
                    if (pwdDlg.ShowDialog() == true)
                    {
                        password = pwdDlg.Password;
                    }
                });

                if (password == null) return sheets;

                // 第二次尝试：带密码打开（会自动重新建立流，避免 closed file 错误）
                workbook = OpenWorkbook(filePath, password);
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

            if (workbook != null)
            {
                ExtractData(workbook, sheets);
                workbook.Close();
            }

            return sheets;
        }

        // 辅助方法：确保每次调用都创建新鲜的流，防止流关闭问题
        private IWorkbook OpenWorkbook(string path, string password)
        {
            // Using 确保文件读取后正确关闭资源
            // 1. 使用 FileStream 以共享模式打开文件
            // FileMode.Open: 打开现有文件
            // FileAccess.Read: 仅请求读取权限
            // FileShare.ReadWrite: 允许其他进程同时进行读写操作（这是解决 Excel 占用问题的关键）
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return WorkbookFactory.Create(fs, password);
            }
        }

        private void ExtractData(IWorkbook workbook, List<ExcelSheetModel> sheets)
        {
            // 3. 遍历所有工作表 (Sheet)
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                var sheetModel = new ExcelSheetModel { SheetName = sheet.SheetName };
                int lastRowNum = sheet.LastRowNum;
                int maxCol = 0;

                // 预扫描获取最大列数
                for (int r = 0; r <= lastRowNum; r++)
                {
                    IRow row = sheet.GetRow(r);
                    if (row != null && row.LastCellNum > maxCol) maxCol = row.LastCellNum;
                }
                sheetModel.MaxColumnCount = maxCol;

                // 遍历该表中有数据的所有行
                for (int r = 0; r <= lastRowNum; r++)
                {
                    IRow row = sheet.GetRow(r);
                    var rowModel = new ExcelRowModel { OriginalIndex = r + 1 };

                    // 遍历从第 1 列到最后一列的所有单元格
                    for (int c = 0; c < maxCol; c++)
                    {
                        string cellValue = "";
                        if (row != null)
                        {
                            ICell cell = row.GetCell(c);
                            cellValue = cell?.ToString() ?? "";
                        }
                        rowModel.Cells.Add(cellValue.Trim());
                    }
                    sheetModel.Rows.Add(rowModel);
                }
                sheets.Add(sheetModel);
            }
        }
    }
}