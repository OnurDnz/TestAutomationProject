using NUnit.Framework;
using NUnit.Framework.Interfaces;
using System;
using TestAutomation.Core.Helpers;

namespace TestAutomation.Core.Bases
{
    public class BaseTest : BrowserFactory
    {
        string xlFilePath = @"C:\Users\onurd\source\repos\Simple\Simple\TestData.xlsx";
        protected ExtentReportsHelper extent;

        [SetUp]
        public void SetUp()
        {
            Driver.Manage().Window.Maximize();
            extent.CreateTest(TestContext.CurrentContext.Test.Name);
        }

        [OneTimeSetUp]
        public void SetUpReporter()
        {
            InitBrowser("Firefox");
            WaitClass.Driver = Driver;
            ExcelTool.ExcelFilePath = xlFilePath;
            extent = new ExtentReportsHelper();
        }

        [TearDown]
        public void AfterTest()
        {
            try
            {
                var status = TestContext.CurrentContext.Result.Outcome.Status;
                var stacktrace = TestContext.CurrentContext.Result.StackTrace;
                var errorMessage = "<pre>" + TestContext.CurrentContext.Result.Message + "</pre>";
                switch (status)
                {
                    case TestStatus.Failed:
                        extent.SetTestStatusFail($"<br>{errorMessage}<br>Stack Trace: <br>{stacktrace}<br>");
                        TakeScreenshot.Take();
                        break;
                    case TestStatus.Skipped:
                        extent.SetTestStatusSkipped();
                        break;
                    default:
                        extent.SetTestStatusPass();
                        TakeScreenshot.Take();
                        break;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                //driver.Quit();
                CloseAllDrivers();
            }
        }


        [OneTimeTearDown]
        public void CloseAll()
        {
            try
            {
                extent.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}