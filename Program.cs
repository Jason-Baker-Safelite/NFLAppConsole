using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


//THIS IS COMMENT ON FEATURE BRANCH 1
//This is the feature branch JIRA 1234

namespace NFL
{
    class Program
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range MyRange = null;
        public enum PlayerPosition
        {
            QB,
            RB
        };
        public static void Main(string[] args)
        {
            string startTime = DateTime.Now.ToString();
            Console.WriteLine("Start time: {0}", startTime);
            ProcessSpreadsheet(SetUpExcel());
            string endTime = DateTime.Now.ToString();
            Console.WriteLine("End time: {0}", endTime);
            Console.ReadKey();

            
            Cleanup();
        }
        public static void ProcessSpreadsheet(object[,] ExcelArray)
        {
            Dictionary<string, Season> QBDictionary = new Dictionary<string, Season>();
            Dictionary<string, Season> RBDictionary = new Dictionary<string, Season>();

            for (int currentRow = 2; currentRow <= ExcelArray.GetLength(0); currentRow++)
            {
                string position = ExcelArray[currentRow, 22].ToString();
                Validate(position, currentRow);
                switch (position)
                {
                    case nameof(PlayerPosition.QB):
                        ProcessPosition(QBDictionary, position, currentRow, ExcelArray);
                        break;
                    case nameof(PlayerPosition.RB):
                        ProcessPosition(RBDictionary, position, currentRow, ExcelArray);
                        break;
                    default:
                        break;
                }
            }
            string bestQB = QBDictionary.OrderByDescending(s => s.Value.PassingYards).First().Key;
            double bestPassingYards = QBDictionary[bestQB].PassingYards;
            Console.WriteLine("The best QB is {0} with {1:0,0} passing yards\n", bestQB, bestPassingYards);
            string bestRB = RBDictionary.OrderByDescending(s => s.Value.RushingYards).First().Key;
            double bestRushingYards = RBDictionary[bestRB].RushingYards;
            Console.WriteLine("The best RB is {0} with {1:0,0} rushing yards", bestRB, bestRushingYards);
        }
        public static void ProcessPosition(Dictionary<string, Season> positionDictionary, string position, int currentRow, object[,] ExcelArray)
        {
            string player = ExcelArray[currentRow, 2].ToString();
            Validate(player, currentRow);
            double passingYards = (double)ExcelArray[currentRow, 11];
            Validate(passingYards, currentRow);
            double rushingYards = (double)ExcelArray[currentRow, 15];
            Validate(rushingYards, currentRow);
            if (positionDictionary.Count == 0)
            {
                Season season = AddSeason(currentRow, position, ExcelArray);
                positionDictionary.Add(season.FullName, season);
            }
            else
            {
                if (positionDictionary.ContainsKey(player))
                {
                    if (position == "QB")
                    {
                        positionDictionary[player].PassingYards += passingYards;
                    }
                    else
                    {
                        positionDictionary[player].RushingYards += rushingYards;
                    }
                }
                else
                {
                    Season season = AddSeason(currentRow, position, ExcelArray);
                    positionDictionary.Add(season.FullName, season);
                }
            }
        }
        public static object[,] SetUpExcel()
        {
            MyApp = new Excel.Application
            {
                Visible = false
            };
            //
            //string XLS_PATH = "C:\\Users\\Jason.Baker\\Projects\\NFLAppConsole\\NFLAppConsole\\NFL_Small_Set.xlsx";
            //string XLS_PATH = Environment.CurrentDirectory+"\\NFL_Small_Set.xlsx";
            //string Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string Path = System.AppContext.BaseDirectory;
            string Path = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            //string Path = AppDomain.CurrentDomain.BaseDirectory;
            string XLS_PATH = Path + "\\NFL_Small_Set.xlsx";
            MyBook = MyApp.Workbooks.Open(XLS_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets["1999-2013 data"];
            MyRange = MySheet.UsedRange;
            object[,] objectArray = (object[,])MyRange.Value2;
            return objectArray;
        }
        public static void Validate(string cellString, int cellRow)
        {
            if (String.IsNullOrEmpty(cellString))
            {
                string exceptionString = string.Format("invalid cell string in row {0}", cellRow.ToString());
                throw new ArgumentException(exceptionString);
            }
        }
        public static void Validate(double cellDouble, int cellRow)
        {
            if (double.IsNaN(cellDouble))
            {
                string exceptionString = string.Format("invalid yards value in row {0}", cellRow.ToString());
                throw new ArgumentException(exceptionString);
            }
        }
        private static Season AddSeason(int currentRow, string position, object[,] ExcelArray)
        {
            Season season = new Season
            {
                FullName = ExcelArray[currentRow, 2].ToString(),
                PassingYards = (double)ExcelArray[currentRow, 11],
                Position = position,
                RushingYards = (double)ExcelArray[currentRow, 15]
            };
            return season;
        }
        private static void Cleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(MyRange);
            Marshal.ReleaseComObject(MySheet);
            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);
        }
    }
}
