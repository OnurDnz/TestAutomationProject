using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System.Collections.Generic;

namespace TestAutomation.Core.Bases
{
    public class BrowserFactory
    {
        private static readonly IDictionary<string, IWebDriver> Drivers = new Dictionary<string, IWebDriver>();

        public static IWebDriver Driver { get; private set; }

        public static void InitBrowser(string browserName)
        {
            switch (browserName)
            {
                case "Firefox":
                    if (Driver == null)
                    {
                        Driver = new FirefoxDriver();
                        Drivers.Add("Firefox", Driver);
                    }
                    break;

                case "Chrome":
                    if (Driver == null)
                    {
                        Driver = new ChromeDriver();
                        Drivers.Add("Chrome", Driver);
                    }
                    break;
            }
        }
        public static void CloseAllDrivers()
        {
            foreach (var key in Drivers.Keys)
            {
                Drivers[key].Close();
                Drivers[key].Quit();
            }
        }
        public static IWebDriver GetDriver()
        {
            return Driver;
        }
    }

    public enum WebBrowsers
    {
        Firefox,
        Chrome
    }
}