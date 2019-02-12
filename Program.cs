using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

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
        public enum SpreadsheetColumn
        {
            yearColumn = 1,
            playerColumn = 2,
            teamColumn = 6,
            passingYardsColumn = 11,
            rushingYardsColumn = 15,
            positionColumn = 22,
            collegeColumn = 26
        };
        public static void Main(string[] args)
        {
            PrintMenu();
            //string startTime = DateTime.Now.ToString();
            //Console.WriteLine("Start time: {0}", startTime);
            //ProcessSpreadsheet(SetUpExcel());
            //string endTime = DateTime.Now.ToString();
            //Console.WriteLine("End time: {0}", endTime);
            Console.ReadKey();
            Cleanup();
        }
        public static void PrintMenu()
        {
            Console.WriteLine("Select Menu Item");
            Console.WriteLine("1. Search by Player");
            Console.WriteLine("2. Search by Team");
            Console.WriteLine("3. Search by Position");
            Console.WriteLine("4. Search by Year");
            Console.WriteLine("5. Search by College");
            string menuSelection = Console.ReadLine();
            if (menuSelection == "1")
            {
                Console.WriteLine("You selected 1");
                SearchByPlayer(menuSelection);
            }
            else if (menuSelection == "2")
            {
                Console.WriteLine("You selected 2");
            }
            else if (menuSelection == "3")
            {
                Console.WriteLine("You selected 3");
            }
            else if (menuSelection == "4")
            {
                Console.WriteLine("You selected 4");
            }
            else if (menuSelection == "5")
            {
                Console.WriteLine("You selected Search by College");
                SearchByCollege(menuSelection);
            }
            else
            {
                Console.WriteLine("Invalid Menu Selection");
                PrintMenu();
            }
        }
        public static void SearchByPlayer(string menuNumber)
        {
            Console.WriteLine("Please Enter Player Name");
            string userInput = Console.ReadLine();
            Console.WriteLine("Search for " + userInput);
            object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(playerArray);
            List<Season> nameCollection = playerCollection.Where(s => s.FullName == userInput).ToList();
        }
        //Dennis college search - start
        public static void SearchByCollege(string menuNumber)
        {
            Console.WriteLine("Please Enter College Name");
            string input = Console.ReadLine();
            while (String.IsNullOrEmpty(input))
            {
                Console.WriteLine("Invalid College Name");
                input = Console.ReadLine();
            }

            // Extract the rows matching the college (name / position / college)
            object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(playerArray);
            List<Season> nameCollection = playerCollection.Where(s => s.College == input).ToList();
            if (nameCollection.Count == 0)
            {
                Console.WriteLine("No entries found for College Name");
                input = Console.ReadLine();
            }
             //Dennis college search - end
        }
        public static List<Season> LoadCollection(object[,] ExcelArray)
        {
            List<Season> spreadsheetCollection = new List<Season>();
            for (int currentRow = 2; currentRow <= ExcelArray.GetLength(0); currentRow++)
            {
                object[,] correctedArray = Validate(currentRow, ExcelArray);
                Season season = AddSeason(currentRow, correctedArray);
                spreadsheetCollection.Add(season);
            }
            return spreadsheetCollection;
        }
        public static object[,] SetUpExcel()
        {
            MyApp = new Excel.Application
            {
                Visible = false
            };
            string XLS_PATH = ConfigurationManager.AppSettings.Get("ExcelPath");
            MyBook = MyApp.Workbooks.Open(XLS_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[ConfigurationManager.AppSettings.Get("SheetName")];
            MyRange = MySheet.UsedRange;
            object[,] objectArray = (object[,])MyRange.Value2;
            return objectArray;
        }
        public static object[,] Validate(int spreadsheetRow, object[,] parmArray)
        {
            if (double.IsNaN((double)parmArray[spreadsheetRow, (int)SpreadsheetColumn.yearColumn]))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.yearColumn] = 0;
            }
            if (String.IsNullOrEmpty(parmArray[spreadsheetRow, (int)SpreadsheetColumn.playerColumn].ToString()))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.playerColumn] = " ";
            }
            if (String.IsNullOrEmpty(parmArray[spreadsheetRow, (int)SpreadsheetColumn.teamColumn].ToString()))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.teamColumn] = " ";
            }
            if (String.IsNullOrWhiteSpace(Convert.ToString(parmArray[spreadsheetRow, (int)SpreadsheetColumn.passingYardsColumn])))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.passingYardsColumn] = 0;
            }
            if (String.IsNullOrWhiteSpace(Convert.ToString(parmArray[spreadsheetRow, (int)SpreadsheetColumn.rushingYardsColumn])))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.rushingYardsColumn] = 0;
            }
            if (String.IsNullOrEmpty(parmArray[spreadsheetRow, (int)SpreadsheetColumn.positionColumn].ToString()))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.positionColumn] = " ";
            }
            if (String.IsNullOrEmpty(parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn].ToString()))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn] = " ";
            }
            return parmArray;
        }
        public static Season AddSeason(int currentRow, object[,] ExcelArray)
        {
            Season season = new Season()
            {
                Year = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.yearColumn]),
                FullName = ExcelArray[currentRow, (int)SpreadsheetColumn.playerColumn].ToString(),
                Team = ExcelArray[currentRow, (int)SpreadsheetColumn.teamColumn].ToString(),
                PassingYards = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.passingYardsColumn]),
                RushingYards = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.rushingYardsColumn]),
                Position = ExcelArray[currentRow, (int)SpreadsheetColumn.positionColumn].ToString(),
                College = ExcelArray[currentRow, (int)SpreadsheetColumn.collegeColumn].ToString()
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
