using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using TestAutomation.Core.Helpers;

namespace TestAutomation.Core.Pages
{
    [CacheLookup]
    public static class MainPage
    {
        public static IWebElement firstResult => WaitClass.WaitUntilFind(By.Name("q"));
        public static IWebElement secondResult => WaitClass.WaitUntilFind(By.Name("btnK"));

    }
}