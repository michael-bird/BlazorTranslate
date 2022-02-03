using ThreeShape.SilverLake.Experiments.SIL165.Views;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace ThreeShape.SilverLake.Experiments.SIL165.Drivers
{
    public sealed class ChromeBrowser : ChromeDriver
    {
        private static ChromeBrowser? _browser;

        public static ChromeBrowser Browser
        {
            get
            {
                if (_browser != null)
                    return _browser;

                var options = new ChromeOptions
                {
                    PageLoadStrategy = PageLoadStrategy.Normal
                };
                options.AddArgument("start-maximized");
                options.AddArgument("ignore-certificate-errors");
                //options.AddArguments("headless");
                //options.AddArguments("--window-size=1920,1028");

                _browser = new ChromeBrowser(@"C:\Git_SilverLakeExperiments\ThreeShape.SilverLake.Experiments.SIL165\ThreeShape.SilverLake.Experiments.SIL165", options);
                return _browser;
            }
        }

        private ChromeBrowser(string chromeDriverDirectory, ChromeOptions options)
                                            : base(chromeDriverDirectory, options)
        {
            new DriverManager().SetUpDriver(new ChromeConfig());
            Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(20);
        }

        public T GetView<T>() where T : BaseView, new()
        {
            return new T
            {
                Browser = this
            };
        }
    }
}
