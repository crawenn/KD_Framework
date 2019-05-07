using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace KD_Framework.cfg
{
    class Constants
    {
        //This is a list of our variables

        public static string baseURL;
        public static string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public static string projectDir = Path.GetDirectoryName(Path.GetDirectoryName(assemblyDir));
        public static string testDataPath = Path.Combine(projectDir, "dataEngine\\Data.xlsx");

        public static string testScriptPath = Path.Combine(projectDir, "extend-src.png");
        public static string reportPath; //TBD
        public static string cfgPath = Path.Combine(projectDir, "\\cfg\\", "cfg.xml");

        //xlDataSheet colnums
        public const int Col_TestCaseID = 0;
        public const int Col_TestScenarioID = 1;
        public const int Col_PageObject = 4;
        public const int Col_ActionKeyword = 5;
        public const int Col_DataSet = 6;
        public const int Col_RunMode = 2;

        //test result keywords
        public const string KEYWORD_FAIL = "FAIL";
        public const string KEYWORD_PASS = "PASS";

        //result columns
        public const int Col_CaseResult = 3;
        public const int Col_TestStepResult = 7;

        //dataEngine xls sheets
        public const string Sheet_TestSteps = "Test Steps";
        public const string Sheet_TestCases = "Test Cases";
    }
}
