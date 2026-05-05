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
    public partial class PasswordDialog : Window
    {
        public string Password => TxtPassword.Password;

        public PasswordDialog(string fileName)
        {
            InitializeComponent();

            // 从资源字典获取基础提示文字
            string baseMsg = Application.Current.TryFindResource("PasswordRequired") as string
                     ?? "This file is encrypted, please enter password:";

            // 格式化显示：提示文字加粗，文件名另起一行
            LblMessage.Text = $"{baseMsg}\n{fileName}";

            // 强制获得焦点
            TxtPassword.Focus();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}