using CommunityToolkit.Mvvm.Messaging.Messages;
using ExcelMerger.ViewModels;

namespace ExcelMerger.Messages
{
    /// <summary>
    /// 列模式下重新比对完成后发送。因为绑定集合是原地更新的（Clear + Add），
    /// DataGrid 不会重新触发生成列，而对齐后的列数/表头可能已变化，
    /// 视图收到此消息后需主动重建当前 Sheet 的 DataGrid 列。
    /// </summary>
    public class SheetRebuiltMessage : ValueChangedMessage<SheetComparisonViewModel>
    {
        public SheetRebuiltMessage(SheetComparisonViewModel value) : base(value) { }
    }
}
