using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using ExpectedConditions = SeleniumExtras.WaitHelpers.ExpectedConditions;

namespace TestAutomation.Core.Helpers
{
    public static class WaitClass
    {
        public static IWebDriver Driver { get; set; }

        /// <summary>
        /// Bu fonksiyon elementi bulana kadar arar ve hata fırlatmaz.
        /// </summary>
        /// <param name="elementLocatorType">Aramak istedigin element tipini yazınız. Örnek kullanım By.İd("button")</param>
        /// <returns>Fonsiyon "IWebElement" tipinde elementi döndürecektir. Eger bulamazsa null deger döndürür.</returns>
        public static IWebElement WaitUntilFind(By elementLocatorType)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                wait.IgnoreExceptionTypes(
                   typeof(NotFoundException),
                   typeof(NoSuchElementException),
                   typeof(ElementNotVisibleException),
                   typeof(StaleElementReferenceException),
                   typeof(ElementNotInteractableException)
                );
                var foundElement = wait.Until(x => x.FindElement(elementLocatorType));
                wait.Until(ExpectedConditions.ElementToBeClickable(foundElement));
                return foundElement;
            }
            catch (Exception e)
            {
                Debug.WriteLine("Some Error: " + e.Message);
                return null;
            }
        }

        /// <summary>
        /// Bu fonksiyon elementi bulana kadar arar ve hata fırlatmaz.
        /// </summary>
        /// <param name="elementLocatorType"></param>
        /// <param name="locator">Aramak istedigin element tipini yazınız. Örnek kullanım By.İd("button")</param>
        /// <returns>Fonsiyon "IWebElement" tipinde elementi döndürecektir. Eger bulamazsa null deger döndürür.</returns>
        public static IWebElement WaitUntilFind(this IWebElement elementLocatorType, By locator)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                wait.IgnoreExceptionTypes(
                   typeof(NotFoundException),
                   typeof(NoSuchElementException),
                   typeof(ElementNotVisibleException),
                   typeof(StaleElementReferenceException),
                   typeof(ElementNotInteractableException)
                );
                var foundElement = wait.Until(x => x.FindElement(locator));
                wait.Until(ExpectedConditions.ElementToBeClickable(foundElement));
                return foundElement;
            }
            catch (Exception e)
            {
                Debug.WriteLine("Some Error: " + e.Message);
                return null;
            }
        }
    }
}