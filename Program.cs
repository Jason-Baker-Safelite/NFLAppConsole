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
        public enum FullNameBreakdown
        {
            firstNameIndex = 0,
            lastNameIndex = 1
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
            string userInput = Console.ReadLine();
            if (userInput == "1")
            {
                Console.WriteLine("You selected 1");
                SearchByPlayer(userInput);
            }
            else if (userInput == "2")
            {
                Console.WriteLine("You selected 2");
                SearchByTeam(userInput);
            }
            else if (userInput == "3")
            {
                Console.WriteLine("You selected 3");
                Console.WriteLine("Please enter 2-character position: ");
                string positionRequest = ValidatePositionInput();
                List<Season> positionResultsCollection = SearchByPosition(positionRequest);
                if (positionResultsCollection.Count == 0)
                {
                    Console.WriteLine("No entries found for position " + positionRequest);
                }
                else
                {
                    DisplayResults(positionResultsCollection);
                }
            }
            else if (userInput == "4")
            {
                Console.WriteLine("You selected 4");
            }
            else if (userInput == "5")
            {
                Console.WriteLine("You selected Search by College");
                SearchByCollege(userInput);
            }
            else
            {
                Console.WriteLine("Invalid Menu Selection");
                PrintMenu();
            }
        }
        //Terry's Code
        private static void SearchByTeam(string menuSelection)
        //{
        //        Console.WriteLine("I'm here now what");
        //}
        {
            Console.WriteLine("Please Enter Team Name");
            string userInput = Console.ReadLine();
            Console.WriteLine("Search for " + userInput);
            object[,] TeamArray = SetUpExcel();
            List<Season> TeamCollection = LoadCollection(TeamArray);
            List<Season> TeamResultCollection = TeamCollection.Where(s => s.Team == userInput).ToList();
            DisplayResults(TeamResultCollection);
        }
        public static string ValidatePositionInput()
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
        public static void DisplayResults(List<Season> displayCollection)
        {

            int skipCount = 0;
            int takeCount = 25;
            int pageCount = 1;
            int totCount = displayCollection.Count;
            int pageTot = (totCount + 24) / 25; // determine number of pages needed to display list
            Console.WriteLine("Number of results " + totCount + " / page count " + pageTot);
            Console.WriteLine(" ");

            Console.WriteLine("Position\tPlayer Name\t\tTeam\tYear\tPassing Yds\tRushing Yds");
            do
            {
                foreach (Season selectedPosition in displayCollection.Skip<Season>(skipCount).Take<Season>(takeCount))
                {
                    Console.WriteLine("{0,5}\t" + "   " + "{1,-25}\t{2,3}\t{3,4}\t{4,8:n0}\t{5," + "" + "8:n0}",
                        selectedPosition.Position,
                        selectedPosition.FullName,
                        selectedPosition.Team,
                        selectedPosition.Year,
                        selectedPosition.PassingYards,
                        selectedPosition.RushingYards);
                }
                pageCount = (skipCount / 25) + 1;
                Console.WriteLine("Display Page " + pageCount + " 'N' Forward / 'P' Backward ' 'X' Exit or Specific Page");
                string pageInput = Console.ReadLine().ToUpper();
                if (pageInput == "N")
                {
                    skipCount += 25;
                }
                else if (pageInput == "P")
                {
                    skipCount -= 25;
                    if (skipCount < 0)
                    {
                        skipCount = 0;
                    }
                }
                else if (pageInput == "X")
                {
                    skipCount = totCount;
                }
                else if (pageInput != " ")
                {
                    int reqPage = int.Parse(pageInput);
                    skipCount = (reqPage * 25) - 25;
                }
                Console.WriteLine("Position\tPlayer Name\t\tTeam\tYear\tPassing Yds\tRushing Yds");
                pageCount = (skipCount * 25) -1;
            } while (skipCount < displayCollection.Count);
            Console.WriteLine("Done with list");
        }
        //public static void DisplayList(List<Season> displayCollection)
        //{
        //    int skipCount = 0;
        //    int takeCount = 25;
        //    int printCount = 0;
        //    int pageCount = 1;
        //    int totCount = displayCollection.Count;
        //    int pageTot = (totCount + 24) / 25; // determine number of pages needed to display list
        //    Console.WriteLine("Number of results " + totCount + " / page count " + pageTot);
        //    Console.WriteLine(" ");

        //    do
        //    {
        //        foreach (Season selectedPosition in displayCollection.Skip<Season>(skipCount).Take<Season>(takeCount))
        //        {

        //            Console.WriteLine("{0,5}\t",
        //                selectedPosition.College,
        //                selectedPosition.FullName);
        //        }
        //        pageCount = (skipCount / 25) + 1;
        //        Console.WriteLine("Display Page " + pageCount + " 'N' Forward / 'P' Backward ' 'X' Exit or Specific Page");
        //        string pageInput = Console.ReadLine().ToUpper();
        //        if (pageInput == "N")
        //        {
        //            skipCount += 25;
        //        }
        //        else if (pageInput == "P")
        //        {
        //            skipCount -= 25;
        //            if (skipCount < 0)
        //            {
        //                skipCount = 0;
        //            }
        //            printCount -= 25;
        //        }
        //        else if (pageInput == "X")
        //        {
        //            skipCount = totCount;
        //        }
        //        else if (pageInput != " ")
        //        {
        //            int reqPage = int.Parse(pageInput);
        //            skipCount = (reqPage * 25) - 25;
        //        }
        //        pageCount = (skipCount * 25) - 1;
        //    } while (skipCount < displayCollection.Count);
        //    Console.WriteLine("Done with list");
        //}
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
            Console.WriteLine("Please Enter College Name or '?' to list colleges");
            string input = Console.ReadLine();
            while (String.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter selection for colleges");
                input = Console.ReadLine();
            }
            if (input == "?")
            {
                Console.WriteLine("Searching for colleges...");

                object[,] playerArray = SetUpExcel();
                List<Season> playerCollection = LoadCollection(playerArray);

                var distCollege = (from z in playerCollection
                                   orderby z.College
                                   select z.College
                                   ).Distinct().ToList();

                int colCount = distCollege.Count;
                Console.WriteLine("Number of results " + colCount);

                foreach (var College in distCollege)
                {
                    Console.WriteLine(College);
                }
                Console.WriteLine("College List complete");
            }
            else
            {
                Console.WriteLine("Searching for players...");

                object[,] playerArray = SetUpExcel();
                List<Season> playerCollection = LoadCollection(playerArray);
                List<Season> nameCollection = playerCollection.Where(s => s.College == input).ToList();

                var distPlayer = (from z in nameCollection
                                  orderby z.FullName
                                  select new Season
                                  {
                                      Year = z.Year,
                                      FullName = z.FullName,
                                      Team = z.Team,
                                      PassingYards = z.PassingYards,
                                      RushingYards = z.RushingYards,
                                      College = z.College,
                                      Position = z.Position
                                  }
                                   ).Distinct().ToList();
                DisplayResults(distPlayer);
            }
            //Dennis college search - end
        }

        public static List<Season> LoadCollection(object[,] ExcelArray)
        {
            List<Season> spreadsheetCollection = new List<Season>();
            for (int currentRow = 2; currentRow <= ExcelArray.GetLength(0); currentRow++)
            {
                object[,] correctedArray = ValidateExcelData(currentRow, ExcelArray);
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
        public static object[,] ValidateExcelData(int spreadsheetRow, object[,] parmArray)
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
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn] = "Unknown";
            }
            if (String.Equals("0", (parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn].ToString())))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn] = "Unknown";
            }
            if (String.Equals("-2146826246", (parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn].ToString())))
            {
                parmArray[spreadsheetRow, (int)SpreadsheetColumn.collegeColumn] = "Unknown";
            }
            return parmArray;
        }
        public static Season AddSeason(int currentRow, object[,] ExcelArray)
        {
            string[] parseFullNameArray = ExcelArray[currentRow, (int)SpreadsheetColumn.playerColumn].ToString().Split(new char[] { ' ' });
            Season season = new Season()
            {
                Year = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.yearColumn]),
                FullName = parseFullNameArray[1] + ", " + parseFullNameArray[0],
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
