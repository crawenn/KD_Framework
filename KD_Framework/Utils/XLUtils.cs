using System;
using Excel = Microsoft.Office.Interop.Excel;
using KD_Framework.cfg;
using KD_Framework.execEngine;

namespace KD_Framework
{
    public class XLUtils
    {
        public static Excel.Application xlApp;
        public static Excel.Workbook xlWB;
        private static Excel.Worksheet xlWS;

        //opens the .xls file specified in string testDataPath in Constants.cs
        public static void setXlFile(String xlPath)
        {
            try
            {
                xlApp = new Excel.Application
                {
                    Visible = false
                };

                xlWB = xlApp.Workbooks.Open(Constants.testDataPath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: setXlFile | Exception: " + e.Message);
            }
        }

        //reading data from the excel cell
        //parameters are rowNum, colNum and sheetName
        public static string getCellData(int rowNum, int colNum, string sheetName)
        {
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                return (xlWS.Cells[rowNum + 1, colNum + 1] as Excel.Range).Value.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: getCellData | Exception: " + e.Message);
                return null;                
            }
        }

        //get rowCount
        public static int getRowCount(string sheetName)
        {
            int number = 0;
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                number = xlWS.Rows.Count;

                return number;
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: getRowCount | Exception: " + e);
                ExecutionScript.bResult = false;
                Console.WriteLine(ExecutionScript.bResult.ToString());
            }
            return number;
        }

        //get rowNumber of the test case
        //arguments are testCaseName, colNum, sheetName
        public static int getRowContains(string testCaseName, int colNum, string sheetName)
        {
            int rowNum = 0;
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                int rowCount = getRowCount(sheetName);

                for (; rowNum < rowCount; rowNum++)
                {
                    if (getCellData(rowNum + 1, colNum, sheetName).Equals(testCaseName))
                    {
                        rowNum++;
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: getRowContains | Exception: " + e.Message);
            }

            return rowNum;
        }

        //counting test steps
        //arguments: sheetName, testCaseID, testCaseStart
        public static int getTestStepsCount(string sheetName, string testCaseID, int testCaseStart)
        {
            int i = 0;
            try
            {
                for (int j = testCaseStart; j <= getRowCount(sheetName); j++)
                {
                    if (!testCaseID.Equals(value: getCellData(i, Constants.Col_TestCaseID, sheetName)))
                    {
                        i = j;
                    }
                }

                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                i = xlWS.UsedRange.Rows.Count + 1;
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: ExcelReder | Method: getRowContains | Exception: " + e.Message);
            }

            return i;
        }

        //writing xl values into cells
        //arguments: result, rowNum, colNum, sheetName
        public static void setCellData(string result, int rowNum, int colNum, string sheetName)
        {
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                var a = (xlWS.Cells[rowNum + 1, colNum + 1] as Excel.Range).Value.ToString();
                (xlWS.Cells[rowNum + 1, colNum + 1] as Excel.Range).Value = result;
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: setCellData | Exception: " + e.Message);
            }
        }
    }
}
