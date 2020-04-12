using NUnit.Framework;
using OpenQA.Selenium.Support.PageObjects;
using TestAutomation.Core.Bases;
using TestAutomation.Core.Pages;

namespace TestAutomation.Core.Tests
{
    [TestFixture]
    class Demo : BaseTest
    {
        [Test]
        public void Primary()
        {
            Driver.Navigate().GoToUrl("https://www.google.com.tr/");
            MainPage.firstResult.SendKeys("Test");
            MainPage.secondResult.Click();
        }
    }
}
