using NUnit.Framework;
using OpenQA.Selenium;
using System;

namespace TestAutomation.Core.Helpers
{
    public static class TakeScreenshot
    {
        static IWebDriver driver = WaitClass.Driver;

        /// <summary>
        /// Bu fonksiyon testin istenilen zamanda da ekran görüntüsü almasını saglar. 
        /// </summary>
        /// <param name="driver"></param>
        public static void Take()
        {
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var path = desktopPath + "\\" + TestContext.CurrentContext.Test.MethodName.Trim() + ".Jpeg";
            Screenshot image = ((ITakesScreenshot)driver).GetScreenshot();
            image.SaveAsFile(path, ScreenshotImageFormat.Jpeg);
        }

        /// <summary>
        /// Bu fonksiyon testin istenilen zamanda da ekran görüntüsü almasını saglar. 
        /// </summary>
        /// <param name="screenShotName">Vermek istediginiz ekran görüntüsü ismi.</param>
        public static void Take(string screenShotName)
        {
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var path = desktopPath + "\\" + screenShotName + ".Jpeg";
            Screenshot image = ((ITakesScreenshot)driver).GetScreenshot();
            image.SaveAsFile(path, ScreenshotImageFormat.Jpeg);
        }

        /// <summary>
        /// Bu fonksiyon testin istenilen zamanda da ekran görüntüsü almasını saglar. 
        /// </summary>
        /// <param name="exportPath">Ekran görüntüsünü çıkarmak istediginiz dosya yolu,Dosyalr arası çift ters slash '\\' kullanmanız gerekmektedir.</param>
        /// <param name="screenShotName">Vermek istediginiz ekran görüntüsü ismi.</param>
        public static void Take(string screenShotName, string exportPath)
        {
            var path = exportPath + "\\" + screenShotName + ".Jpeg";
            Screenshot image = ((ITakesScreenshot)driver).GetScreenshot();
            image.SaveAsFile(path, ScreenshotImageFormat.Jpeg);
        }
    }
}