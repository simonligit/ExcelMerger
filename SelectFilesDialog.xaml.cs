using CommunityToolkit.Mvvm.Messaging;
using ExcelMerger.Messages;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelMerger
{
    /// <summary>
    /// Interaction logic for SelectFilesDialog.xaml
    /// </summary>
    public partial class SelectFilesDialog : Window
    {
        public SelectFilesDialog()
        {
            InitializeComponent();
        }
        public string File1Path { get; private set; }
        public string File2Path { get; private set; }

        private void BrowseFile1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                // 显示文件路径并存入一个属性，方便主窗口读取
                TxtFile1.Text = openFileDialog.FileName;
                this.File1Path = openFileDialog.FileName;
            }
        }

        private void BrowseFile2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                // 显示文件路径并存入一个属性，方便主窗口读取
                TxtFile2.Text = openFileDialog.FileName;
                this.File2Path = openFileDialog.FileName;
            }
        }
        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtFile1.Text) && !string.IsNullOrEmpty(TxtFile2.Text))
            {
                // 发送消息，包含两个路径
                WeakReferenceMessenger.Default.Send(new FilePathsSelectedMessage(new[] { TxtFile1.Text, TxtFile2.Text }));

                // 关闭对话框
                this.DialogResult = true;
                this.Close();
            }
        }
        public void SetFile1(string path)
        {
            TxtFile1.Text = path;
            this.File1Path = path;
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // 逻辑 1：拖入两个文件
                if (files.Length >= 2)
                {
                    if (IsExcelFile(files[0]) && IsExcelFile(files[1]))
                    {
                        TxtFile1.Text = files[0];
                        this.File1Path = files[0];
                        TxtFile2.Text = files[1];
                        this.File2Path = files[1];
                    }
                }
                // 逻辑 2：拖入一个文件
                else if (files.Length == 1)
                {
                    if (IsExcelFile(files[0]))
                    {
                        // 如果第一个框为空，填入第一个；否则填入第二个
                        if (string.IsNullOrEmpty(TxtFile1.Text))
                        {
                            TxtFile1.Text = files[0];
                            this.File1Path = files[0];
                        }
                        else
                        {
                            TxtFile2.Text = files[0];
                            this.File2Path = files[0];
                        }
                    }
                }
            }
        }

        private bool IsExcelFile(string path)
        {
            string ext = System.IO.Path.GetExtension(path).ToLower();
            return ext == ".xlsx" || ext == ".xls";
        }
        private void SwapFiles_Click(object sender, RoutedEventArgs e)
        {
            // 1. 交换文本框中的显示内容
            string tempText = TxtFile1.Text;
            TxtFile1.Text = TxtFile2.Text;
            TxtFile2.Text = tempText;

            // 2. 交换后台存储的路径属性
            string tempPath = this.File1Path;
            this.File1Path = this.File2Path;
            this.File2Path = tempPath;
        }
    }
}
