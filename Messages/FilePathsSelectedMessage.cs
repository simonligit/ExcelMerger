using CommunityToolkit.Mvvm.Messaging.Messages;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger.Messages
{
    public class FilePathsSelectedMessage : ValueChangedMessage<string[]>
    {
        public FilePathsSelectedMessage(string[] value) : base(value) { }
    }
}
