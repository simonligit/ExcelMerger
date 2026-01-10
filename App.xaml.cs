using System.Configuration;
using System.Data;
using System.Windows;

namespace ExcelMerger
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            SetLanguage();
        }

        private void SetLanguage()
        {
            string resourceName;

            // 1. 优先尝试从用户设置中读取
            // 假设你在 Settings 中定义了一个名为 LastLanguage 的字符串属性
            if (!string.IsNullOrEmpty(ExcelMerger.Properties.Settings.Default.LastLanguage))
            {
                resourceName = ExcelMerger.Properties.Settings.Default.LastLanguage;
            }
            else
            {
                string culture = System.Globalization.CultureInfo.CurrentCulture.Name;
                resourceName = culture.StartsWith("zh")
                                      ? "Resources/StringResources.zh-CN.xaml"
                                      : "Resources/StringResources.en.xaml";
                // 初始化时也可以顺便把这个默认值存入设置
                ExcelMerger.Properties.Settings.Default.LastLanguage = resourceName;
                ExcelMerger.Properties.Settings.Default.Save();
            }

            try
            {
                ResourceDictionary dict = new ResourceDictionary { Source = new Uri(resourceName, UriKind.Relative) };

                // 3. 替换现有的语言字典
                Application.Current.Resources.MergedDictionaries.Clear();
                Application.Current.Resources.MergedDictionaries.Add(dict);
            }
            catch (Exception)
            {
                // 防御性处理：如果读取设置的路径失效（比如文件被删了），回退到英文
                ResourceDictionary fallbackDict = new ResourceDictionary
                {
                    Source = new Uri("Resources/StringResources.en.xaml", UriKind.Relative)
                };
                Application.Current.Resources.MergedDictionaries.Clear();
                Application.Current.Resources.MergedDictionaries.Add(fallbackDict);
            }
        }
    }

}
