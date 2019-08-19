using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using IAlert = OpenQA.Selenium.IAlert;
using NUnit.Framework;

namespace AutoPOC.Utils
{
    class BaseFunctions
    {
        public IWebDriver driver;

        public string driverPath = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\")), "Drivers");
        public string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "Config.csv");
        public string elementsRepoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ElementsRepo", "ElementsRepo.csv");
        public string testDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestData", "TestData.csv");

        public string testCaseName = TestContext.CurrentContext.Test.Name;
        public string timeStamp = DateTime.Now.ToString("dd_MM_yyy_HH_mm_ss_fff");
        public string screenShotPath;
        public string testOutput;
        public StringBuilder csv;
        public int screenShotCount = 0;
        public int stepCount = 0;

        public BaseFunctions()
        {
            Directory.CreateDirectory(Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\Results")), testCaseName + "_" + timeStamp));
            testOutput = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\Results")), testCaseName + "_" + timeStamp, testCaseName + "_Report.csv");
            screenShotPath = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\Results")), testCaseName + "_" + timeStamp);

            csv = new StringBuilder();
            writeResultHeader();
        }

        public void getBrowser()
        {
            string browser = getConfigData("Browser");
            switch (browser.ToLower())
            {
                case "chrome":
                    ChromeOptions chromeOptions = new ChromeOptions();
                    chromeOptions.Proxy = null;
                    driver = new ChromeDriver(driverPath, chromeOptions);
                    break;
                case "firefox":
                    FirefoxOptions firefoxOptions = new FirefoxOptions();
                    firefoxOptions.Proxy = null;
                    driver = new FirefoxDriver(driverPath, firefoxOptions);
                    break;
                case "ie":
                    InternetExplorerOptions ieOptions = new InternetExplorerOptions();
                    ieOptions.Proxy = null;
                    driver = new InternetExplorerDriver(driverPath, ieOptions);
                    break;
                case "edge":
                    EdgeOptions edgeOptions = new EdgeOptions();
                    edgeOptions.Proxy = null;
                    driver = new EdgeDriver(driverPath, edgeOptions);
                    break;
            }
        }

        public void launchWebsite()
        {
            string URL = getConfigData("URL");
            driver.Navigate().GoToUrl(URL);
            driver.Manage().Window.Maximize();
            Thread.Sleep(TimeSpan.FromSeconds(1));
            takeScreenshot("Website", "URL");
        }

        public IWebElement getElementsFromRepo(string elementName)
        {
            string ex;
            string elementType = getDataFromCsv(elementsRepoPath, "ElementName", elementName, "ElementType", out ex);
            string elementIdentifier = getDataFromCsv(elementsRepoPath, "ElementName", elementName, "ElementIdentifier", out ex);
            string screenName = getDataFromCsv(elementsRepoPath, "ElementName", elementName, "ScreenName", out ex);

            IWebElement dynamicElement;
            switch (elementType.ToLower())
            {
                case "id":
                    dynamicElement = driver.FindElement(By.Id(elementIdentifier));
                    break;
                case "xpath":
                    dynamicElement = driver.FindElement(By.XPath(elementIdentifier));
                    break;
                case "linktext":
                    dynamicElement = driver.FindElement(By.LinkText(elementIdentifier));
                    break;
                case "cssselector":
                    dynamicElement = driver.FindElement(By.CssSelector(elementIdentifier));
                    break;
                case "classname":
                    dynamicElement = driver.FindElement(By.ClassName(elementIdentifier));
                    break;
                default:
                    return null;
            }
            HighlightElement(dynamicElement);
            takeScreenshot(screenName, elementName);
            UnHighlightElement(dynamicElement);
            return dynamicElement;
        }

        public string getDataFromCsv(string csvFilePath, string uniqueFieldName, string uniqueFieldValue, string targetFieldName, out string strException)
        {
            int count = 0;
            try
            {
                var reader = File.OpenText(csvFilePath);
                var csv = new CsvReader(reader);
                while (csv.Read())
                {
                    Console.WriteLine("count=" + count);
                    if (csv.GetField<string>(uniqueFieldName) == uniqueFieldValue)
                    {
                        strException = null;
                        string targetFieldValue = csv.GetField<string>(targetFieldName);
                        csv.Dispose();
                        return targetFieldValue;
                    }
                    count++;
                }
                strException = null;
                return null;
            }
            catch (Exception e)
            {
                strException = e.ToString();
                Console.WriteLine(strException);
                return null;
            }
        }

        public string getConfigData(string configName)
        {
            return getDataFromCsv(configPath, "ConfigName", configName, "ConfigValue", out string exception);
        }

        public string getTestData(string testCaseName, string testDataField)
        {
            return getDataFromCsv(testDataPath, "TestCaseName", testCaseName, testDataField, out string exception);
        }

        public void takeScreenshot(string screenName, string elementName)
        {
            string timestamp = DateTime.Now.ToString("dd_MM_yyy_HH_mm_ss_fff");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(screenShotPath + "\\" + ++screenShotCount + "_" + screenName + "_" + elementName + "_Screenshot.png");
        }

        public void seleniumExecuter(IWebElement element, string action, string elementName, string testData)
        {
            switch (action.ToLower())
            {
                case "input":
                    element.SendKeys(testData);
                    break;
                case "click":
                    element.Click();
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    break;
            }
            takeScreenshot(elementName, action);
        }

        public void KeyboardAction(string Action, int iteration)
        {
            Actions action = new Actions(driver);
            switch (Action)
            {
                case "PageDown":
                    for (int ite = 1; ite <= iteration; ite++)
                    {
                        action.SendKeys(OpenQA.Selenium.Keys.PageDown).Build().Perform();
                    }
                    break;

                case "PageUp":
                    for (int ite = 1; ite <= iteration; ite++)
                    {
                        action.SendKeys(OpenQA.Selenium.Keys.PageUp).Build().Perform();
                    }
                    break;

                case "Enter":
                    for (int ite = 1; ite <= iteration; ite++)
                    {
                        action.SendKeys(OpenQA.Selenium.Keys.Enter).Build().Perform();
                    }
                    break;

                case "ArrowDown":
                    for (int ite = 1; ite <= iteration; ite++)
                    {
                        action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Build().Perform();
                    }
                    break;

                case "ArrowUp":
                    for (int ite = 1; ite <= iteration; ite++)
                    {
                        action.SendKeys(OpenQA.Selenium.Keys.ArrowUp).Build().Perform();
                    }
                    break;

                case "Tab":
                    for (int ite = 1; ite <= iteration; ite++)
                    {
                        action.SendKeys(OpenQA.Selenium.Keys.Tab).Build().Perform();
                    }
                    break;
            }
         }

        public void Switch_Tab(int intTabNumber)
        {
            Actions builder = new Actions(driver);
            builder.KeyUp(OpenQA.Selenium.Keys.Control).SendKeys(OpenQA.Selenium.Keys.Tab);
            IAction switchTabs = builder.Build();
            switchTabs.Perform();
            Thread.Sleep(TimeSpan.FromSeconds(1));
            var driverHandles = new List<string>();
            foreach(var windowHandle in driver.WindowHandles)

            driverHandles.Add(windowHandle);
            try
            {
                driver.SwitchTo().Window(driverHandles[intTabNumber - 1]);
            }
            catch(Exception)
            {
                driver.SwitchTo().Window(driverHandles[intTabNumber - 2]);
            }
        }

        public string ReadAndHandleAlert()
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.AlertIsPresent());
                IAlert alert = driver.SwitchTo().Alert();
                string strWarning = alert.Text;
                alert.Accept();
                return strWarning;
            }

            catch(Exception)
            {
                return "Alert not displayed";
            }
        }

        public bool DropDownOptionAvailability(SelectElement selectElement, string optionTextToCheck)
        {
            IList<IWebElement> elementOptions = selectElement.Options;
            foreach (var option in elementOptions)
            {

                if (option.Text == optionTextToCheck)
                    return true;
            }
            return false;
         }
                
        public void HighlightElement(IWebElement element)
        {
            try
            {
                var jsDriver = (IJavaScriptExecutor)driver;
                string highlightJavascript = @"$(arguments[0]).css({ ""border-width"" : ""2px"", ""border-style"" : ""solid"", ""border-color"" : ""red"", ""background"" : """" });";
                jsDriver.ExecuteScript(highlightJavascript, new object[] { element });
            }
            catch (Exception e)
            {
                string strException = e.ToString();
                Console.WriteLine(strException);
            }            
        }

        public void UnHighlightElement(IWebElement element)
        {
            try
            {
                var jsDriver = (IJavaScriptExecutor)driver;
                string highlightJavascript = @"$(arguments[0]).css({ ""border-width"" : ""2px"", ""border-style"" : ""solid"", ""border-color"" : """", ""background"" : """" });";
                jsDriver.ExecuteScript(highlightJavascript, new object[] { element });
            }
            catch (Exception e)
            {
                string strException = e.ToString();
                Console.WriteLine(strException);
            }
        }

        public void elementScrollToView(IWebElement obj_element)
        {
            var js = (IJavaScriptExecutor)driver;
            IWebElement Welement = obj_element;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", Welement);
        }

        public void writeResultHeader()
        {            
            var header = string.Format("{0},{1},{2},{3},{4}", "Step No", "Scenario", "Expected Result", "Actual Result", "Status");
            csv.AppendLine(header);            
        }

        public void Results(string Scenario, string expectedResult, string actualResult, string Status)
        {
            var newLine = string.Format("{0},{1},{2},{3},{4}", ++stepCount, Scenario, expectedResult, actualResult, Status);
            csv.AppendLine(newLine); 
            File.WriteAllText(testOutput, csv.ToString());
        }

        public void closeBrowser()
        {
            takeScreenshot("LastStep", "ClosingBrowser");
            driver.Close();
            driver.Quit();
            driver.Dispose();
        }

    }
}
