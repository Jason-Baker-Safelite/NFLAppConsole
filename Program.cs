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
                Console.WriteLine("Please enter 2-character position: ");
                string userInput = ValidateInput();
                List<Season> positionResultsCollection = SearchByPosition(userInput);
                if (positionResultsCollection.Count == 0)
                {
                    Console.WriteLine("No entries found for position " + userInput);
                }
                else
                {
                    DisplayPosition(positionResultsCollection);
                }
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
        public static string ValidateInput()
        {
            bool validPositon = false;
            string userEntry = Console.ReadLine().ToUpper();
            do
            {
                if (userEntry.Length == 2)
                {
                    validPositon = true;
                }
                else
                {
                    Console.WriteLine("Please enter 2-character position: ");
                    userEntry = Console.ReadLine().ToUpper();
                }
            } while (validPositon == false);
            return userEntry;
        }
        public static List<Season> SearchByPosition(string validInput)
        {
            Console.WriteLine("Searching for " + validInput + "...");
            object[,] positionArray = SetUpExcel();
            List<Season> positionCollection = LoadCollection(positionArray).Where(s => s.Position == validInput).OrderBy(o => o.FullName).ThenBy(o => o.Year).ToList();
            return positionCollection;
        }
        public static void DisplayPosition(List<Season> displayCollection)
        {
            Console.WriteLine("Position\t    Player Name\t\tTeam\tYear\tPassing Yds\tRushing Yds");
            foreach (Season selectedPosition in displayCollection)
            {
                Console.WriteLine("{0,5}\t{1,25}\t{2,3}\t{3,4}\t{4,8:n0}\t{5,8:n0}",
                    selectedPosition.Position,
                    selectedPosition.FullName,
                    selectedPosition.Team,
                    selectedPosition.Year,
                    selectedPosition.PassingYards,
                    selectedPosition.RushingYards);
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
        public static void SearchByCollege(string input)
        {
            Console.WriteLine("Please Enter College Name");
            input = Console.ReadLine();

            if (String.IsNullOrEmpty(input))
            {
                Console.WriteLine("Invalid College Name");
                // ?????? - how do I return to the top of "SearchByCollege"??????
            }
            else
            {
                // Extract the rows matching the college (name / position / college)
                object[,] PlayerArray = SetUpExcel();
                List<Season> collegeDict = LoadCollection(PlayerArray);
                List<Season> nameCollection = collegeDict.Where(s => s.College == input).ToList();
                int count1 = collegeDict.Count;
                if (count1 > 1)
                {
                    Console.WriteLine("College Name not unique");
                }
                if (count1 == 0)
                {
                    Console.WriteLine("College Name not found");
                }
                else
                {
                    Console.WriteLine("College data display");
                }
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
