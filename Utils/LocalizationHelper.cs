using System;
using System.Windows;

namespace ExcelMerger.Utils
{
    internal static class LocalizationHelper
    {
        internal static string Get(string key)
        {
            if (string.IsNullOrEmpty(key)) return string.Empty;
            try
            {
                var value = Application.Current?.Resources?[key] as string;
                if (!string.IsNullOrEmpty(value)) return value;
            }
            catch
            {
                // swallow and fallback
            }
            return key; // fallback to key to make missing resource visible
        }

        internal static string GetFormat(string key, params object[] args)
        {
            var template = Get(key);
            try
            {
                return string.Format(template, args);
            }
            catch
            {
                return template;
            }
        }
    }
}