using System;
using Excel = Microsoft.Office.Interop.Excel;
using KD_Framework.cfg;

namespace KD_Framework
{
    public class XLUtils
    {
        public static Excel.Application xlApp;
        public static Excel.Workbook xlWB;
        public static Excel.Worksheet xlWS;

        //opens the .xls file specified in string testDataPath in Constants.cs
        public static void setXlFile(String xlPath)
        {
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;

                xlWB = xlApp.Workbooks.Open(Constants.testDataPath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: setXlFile | Exception: " + e.Message);
            }
        }

        //reading data from the excel cell
        //parameters are rowNum, colNum and sheetName
        public static string getCellData(int rowNum, int colNum, String sheetName)
        {
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                var cellValue = (xlWS.Cells[rowNum + 1, colNum + 1] as Excel.Range).Value.ToString();
                return cellValue;
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
            int i = 0;
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                i = xlWS.UsedRange.Rows.Count;
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: getRowCount | Exception: " + e.Message);
            }

            return i;
        }

        //get rowNumber of the test case
        //arguments are testCaseName, colNum, sheetName
        public static int getRowContains(String testCaseName, int colNum, String sheetName)
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
        public static int getTestStepsCount(String sheetName, String testCaseID, int testCaseStart)
        {
            int i = 0;
            try
            {
                for (int j = testCaseStart; j <= XLUtils.getRowCount(sheetName); j++)
                {
                    if (!testCaseID.Equals(XLUtils.getCellData(i, Constants.Col_TestCaseID, sheetName)))
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
        public static void setCellData(String result, int rowNum, int colNum, String sheetName)
        {
            try
            {
                xlWS = xlWB.Sheets[sheetName] as Excel.Worksheet;
                string a = (xlWS.Cells[rowNum + 1, colNum + 1] as Excel.Range).Value.ToString();
                (xlWS.Cells[rowNum + 1, colNum + 1] as Excel.Range).Value = result;
            }
            catch (Exception e)
            {
                Console.WriteLine("Class: XLUtils | Method: setCellData | Exception: " + e.Message);
            }
        }
    }
}
