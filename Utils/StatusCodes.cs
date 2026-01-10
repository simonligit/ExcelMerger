namespace ExcelMerger.Utils
{
    /// <summary>
    /// Centralized internal status tokens used by CompareService and XAML triggers.
    /// These are internal identifiers and not user-facing labels.
    /// </summary>
    internal static class StatusCodes
    {
        /// <summary>Rows unchanged</summary>
        public const string Same = "Same";

        /// <summary>Rows modified</summary>
        public const string Modified = "Modified";

        /// <summary>Row added (present on right side)</summary>
        public const string Added = "Added";

        /// <summary>Row deleted (present on left side only)</summary>
        public const string Deleted = "Deleted";

        /// <summary>Placeholder empty row</summary>
        public const string Empty = "Empty";
    }
}