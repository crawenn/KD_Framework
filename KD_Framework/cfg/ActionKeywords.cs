using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System.Configuration;
using System.Threading;
using System.Runtime.InteropServices;
using KD_Framework.cfg;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace KD_Framework.cfg
{
    public class ActionKeywords
    {
        public static IWebDriver driver;

        public static void openBrowser(String obj, String data)
        {
            try
            {
                if (data == @"((M|m)ozill?a)?|((F|f)ire(F|f)ox)?")
                {
                    driver = new FirefoxDriver();
                    Console.WriteLine("Launching" + data);
                }
                else if (data == @"((I|i)nternet)?|((E|e)xplorer)?|IE")
                {
                    driver = new InternetExplorerDriver();
                    Console.WriteLine("Launching" + data);
                }
                else if (data == @"(C|c)hrome")
                {
                    driver = new ChromeDriver();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to open browser " + e.Message + ". Please double check the test data.");
            }
        }

        public static void navigate(String obj, String data)
        {
            try
            {
                for (int second = 0; ; second++)
                {
                    if (second >= 60) Assert.Fail("timeout");
                    try
                    {
                        if (Regex.IsMatch(driver.FindElement(By.CssSelector("BODY")).Text, "^[\\s\\S]*$")) break;
                    }
                    catch (Exception)
                    {

                    }
                    Thread.Sleep(1000);
                }

                driver.Navigate().GoToUrl(Constants.baseURL);
            }
            catch (Exception e)
            {
                Console.WriteLine("The baseURL " + Constants.baseURL + " could not be reached.");
            }
        }

        public static void click(String obj, String data)
        {
            try
            {
                driver.FindElement(By.XPath(getKey(obj, ""))).Click();
            }
            catch (Exception e)
            {
                Console.WriteLine("Element not found: " + obj);
            }
        }

        public static string getKey(String obj, String data)
        {
            NotImplementedException e = new NotImplementedException();
            return e.ToString();
            //return cfg.Settings.Default.Properties[obj].DefaultValue.ToString();            
        }

        public static void input(String obj, String data)
        {
            try
            {
                IWebElement we = driver.FindElement(By.XPath(getKey(obj, "")));
                we.Clear();
                we.SendKeys(data);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void waitFor(String obj, String data)
        {
            try
            {
                Thread.Sleep(5000);
            }
            catch (ThreadInterruptedException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void waitUntil(String obj, String data)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(data)));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        public static void checkCheckBox(String obj, String data)
        {
            try
            {
                IWebElement chkbx = driver.FindElement(By.XPath(getKey(obj, "")));
                if (!chkbx.Selected)
                {
                    chkbx.Click();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void select(String obj, String data)
        {
            try
            {
                SelectElement select = new SelectElement(driver.FindElement(By.XPath(getKey(obj, ""))));
                select.SelectByText(data);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void submitForm(String obj, String data)
        {
            try
            {
                driver.FindElements(By.XPath(data))[0].Submit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void confirmOrder(String obj, String data)
        {
            NotImplementedException e = new NotImplementedException();
            Console.WriteLine(e.Message);
        }

        public static void closeBrowser(String obj, String data)
        {
            try
            {
                driver.Close();
                driver.Quit();
                driver.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine("Browser could not be closed: " + e.Message);
            }
        }

        public static void takeSnip(String data, String obj)
        {
            Screenshot prntscr = ((ITakesScreenshot)driver).GetScreenshot();
            string screenshot = prntscr.AsBase64EncodedString;
            byte[] screenShotAsByteArreay = prntscr.AsByteArray;

            prntscr.SaveAsFile("selTests_" + DateTime.Today.ToString() + data + obj, ScreenshotImageFormat.Png);
            prntscr.ToString();
        }
    }
}
