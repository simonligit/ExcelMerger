using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger.Models
{
    public partial class ColumnConfig : ObservableObject
    {
        public required string HeaderName { get; set; }
        public int Index { get; set; }

        [ObservableProperty]
        private bool _isKeyColumn;
    }
}
