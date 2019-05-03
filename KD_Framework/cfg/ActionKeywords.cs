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
                    { }
                    Thread.Sleep(1000);
                }

                driver.Navigate().GoToUrl(Constants.baseURL);
            }
            catch (Exception e)
            {
                Console.WriteLine("The baseURL " + Constants.baseURL + " could not be reached.");
            }
        }
    }
}
