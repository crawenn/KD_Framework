using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using KD_Framework.cfg;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Chrome;

namespace KD_Framework.execEngine
{
    [TestClass]
    public class ExecutionScript
    {
        public static ActionKeywords actionKeywords;
        public static string sActionKeyword;
        public static string sPageObject;
        //reflection class object
        public static MethodInfo methodInfo;

        public static int iTestStep;
        public static int iTestLastStep;
        public static string sTestCaseID;
        public static string sRunMode;
        public static string sData;
        public static bool bResult;

        [TestMethod]
        public static void TestMain(TestContext testContext)
        {
            actionKeywords = new ActionKeywords();

            //path of the xls
            string sPath = Constants.testDataPath;

            XLUtils.setXlFile(sPath);

            execute_TestCase();
        }

        private static void execute_TestCase()
        {
            //get total testcase num
            int iTotalTestCases = XLUtils.getRowCount(Constants.Sheet_TestCases);
            //the loop will execute the number of times equal to total number of test cases
            for (int iTestCase = 1; iTestCase <= iTotalTestCases; iTestCase++)
            {
                bResult = true;
                sTestCaseID = XLUtils.getCellData(iTestCase, Constants.Col_TestCaseID, Constants.Sheet_TestCases);

                sRunMode = XLUtils.getCellData(iTestCase, Constants.Col_RunMode, Constants.Sheet_TestCases);

                if (sRunMode != null && sRunMode.Equals("Yes"))
                {
                    iTestStep = XLUtils.getRowContains(sTestCaseID, Constants.Col_TestCaseID, Constants.Sheet_TestSteps);

                    iTestLastStep = XLUtils.getTestStepsCount(Constants.Sheet_TestSteps, sTestCaseID, iTestStep);

                    bResult = true;

                    for (; iTestStep <= iTestLastStep; iTestStep++)
                    {
                        sActionKeyword = XLUtils.getCellData(iTestStep, Constants.Col_ActionKeyword, Constants.Sheet_TestSteps);

                        sPageObject = XLUtils.getCellData(iTestStep, Constants.Col_PageObject, Constants.Sheet_TestSteps);

                        sData = XLUtils.getCellData(iTestStep, Constants.Col_DataSet, Constants.Sheet_TestSteps);
                        execute_Actions();

                        if (!bResult)
                        {
                            XLUtils.setCellData(Constants.KEYWORD_FAIL, iTestCase, Constants.Col_CaseResult, Constants.Sheet_TestCases);
                            break;
                        }
                    }

                    if (bResult)
                    {
                        XLUtils.setCellData(Constants.KEYWORD_PASS, iTestCase, Constants.Col_CaseResult, Constants.Sheet_TestCases);
                    }
                }
            }            
        }

        private static void execute_Actions()
        {
            string sResult;

            if (sActionKeyword != null)
            {
                MethodInfo mi = actionKeywords.GetType().GetMethod(sActionKeyword);

                string[] args = { sPageObject, sData };
                mi.Invoke(sActionKeyword, args);

                sResult = bResult ? Constants.KEYWORD_PASS : Constants.KEYWORD_FAIL;

                XLUtils.setCellData(sResult, iTestStep, Constants.Col_TestStepResult, Constants.Sheet_TestSteps);
                if (!bResult)
                {
                    ActionKeywords.closeBrowser("", "");
                }
            }
        }

        [ClassCleanup]
        public static void TestCloseApp()
        {
            XLUtils.xlWB.Save();
            XLUtils.xlWB.Close(0);
            XLUtils.xlApp.Quit();
        }

        [ClassInitialize]
        public void TestSetup()
        {

        }

        [TestInitialize]
        public void InitTest()
        {

        }

        [TestCleanup]
        public void CleanupTest()
        {

        }
    }
}
