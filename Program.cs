using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace NFL
{
    class Program
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range MyRange = null;
        private static object[,] ExcelArray = SetUpExcel();

        public enum FullNameBreakdown
        {
            firstNameIndex = 0,
            lastNameIndex = 1
        };

        //************************************************************************
        // Set the EXCEL column numbers to English names
        //************************************************************************
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

        //************************************************************************
        // Main processing loop
        //************************************************************************
        public static void Main(string[] args)
        {
            NFLAppConsole.NFLSearch NflSearch = new NFLAppConsole.NFLSearch();


            Cleanup();
            bool Mainloop = true;
            do
            {
                Mainloop = PrintMenu(Mainloop);
            } while (Mainloop == true);

            //string startTime = DateTime.Now.ToString();
            //Console.WriteLine("Start time: {0}", startTime);
            //string endTime = DateTime.Now.ToString();
            //Console.WriteLine("End time: {0}", endTime);

            Console.ReadKey();
        }

        //************************************************************************
        // Selection menu
        //************************************************************************
        public static bool PrintMenu(bool Printloop)
        {
            Console.WriteLine("Select Menu Item");
            Console.WriteLine("1. Search by Player");
            Console.WriteLine("2. Search by Team");
            Console.WriteLine("3. Search by Position");
            Console.WriteLine("4. Search by Year");
            Console.WriteLine("5. Search by College");
            Console.WriteLine("6. Exit");

            string userInput = Console.ReadLine();
            if (userInput == "1")
            {
                Console.WriteLine("You selected Search by Player");
                SearchByPlayer(userInput);
            }
            else if (userInput == "2")
            {
                Console.WriteLine("You selected Search by Team");
                SearchByTeam(userInput);
            }
            else if (userInput == "3")
            {
                Console.WriteLine("You selected Search by Position");
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
                Console.WriteLine("You selected Search by Year");
                SearchByYear(userInput);
            }
            else if (userInput == "5")
            {
                Console.WriteLine("You selected Search by College");
                SearchByCollege(userInput);
            }
            else if (userInput == "6")
            {
                Console.WriteLine("Goodbye");
                Printloop = false;
             }
            else
            {
                Console.WriteLine("Invalid Menu Selection");
//                PrintMenu();
            }
            return Printloop;
        }

        //************************************************************************
        // Seach by year processing - Randy
        //************************************************************************
        private static void SearchByYear(string userInput)
        {
            Console.WriteLine("Please Enter Year");
            string SelectYear = Console.ReadLine();
            Console.WriteLine("Search for " + SelectYear);
            //object[,] YearArray = SetUpExcel();
            List<Season> YearCollection = LoadCollection(ExcelArray);
            List<Season> YearResultCollection = YearCollection.Where(s => s.Year.ToString() == SelectYear).ToList();
            DisplayResults(YearResultCollection);
        }

        //************************************************************************
        // Seach by position processing -Jason
        //************************************************************************
        public static List<Season> SearchByPosition(string validInput)
        {
            Console.WriteLine("Searching for " + validInput + "...");
            //object[,] positionArray = SetUpExcel();
            List<Season> positionCollection = LoadCollection(ExcelArray).Where(s => s.Position == validInput).OrderBy(o => o.FullName).ThenBy(o => o.Year).ToList();
            return positionCollection;
        }

        //************************************************************************
        // Seach by team processing - Terry
        //************************************************************************
        private static void SearchByTeam(string menuSelection)
        {

            Console.WriteLine("Please Enter Team Name or '?' to list teams");
            string input = Console.ReadLine();
            while (String.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter selection for teams");
                input = Console.ReadLine();
            }
            if (input == "?")
            {
                AllTeamList();
            }
            else
            {
                SelTeamPlayerList(input);
            }
        }

        //************************************************************************
        // Create collection of players for specific team
        //************************************************************************
        private static void SelTeamPlayerList(string input)
        {
            Console.WriteLine("Searching for players from " + input);

            //object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(ExcelArray);
            List<Season> nameCollection = playerCollection.Where(s => s.Team == input).ToList();

            var distPlayer = (from z in nameCollection
                              orderby z.FullName, z.Year, z.Team
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

        //************************************************************************
        // Create collection of unique teams
        //************************************************************************
        private static void AllTeamList()
        {
            Console.WriteLine("Searching for teams...");

            //object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(ExcelArray);
 
            var distTeam = (from z in playerCollection
                            orderby z.Team
                            select z.Team
                               ).Distinct().ToList();

            DisplayStringList(distTeam);

            int EntryMax = distTeam.Count;
            string listSelection = FindSelectedEntry(distTeam, EntryMax);

            SelTeamPlayerList(listSelection);
        }

        //************************************************************************
        // Seach by Player processing - Liping
        //************************************************************************
        public static void SearchByPlayer(string menuNumber)
        {
            Console.WriteLine("Please Enter Player Name or '?' to list all players");
            string input = Console.ReadLine();
            while (String.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter selection for player");
                input = Console.ReadLine();
            }
            if (input == "?")
            {
                AllPlayerList();
            }
            else
            {
                SelPlayerPlayerList(input);
            }
        }
        //************************************************************************
        // Create collection of unique players 
        //************************************************************************
        private static void AllPlayerList()
        {
            Console.WriteLine("Searching for players...");

            //object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(ExcelArray);

            var distPlayer = (from z in playerCollection
                              orderby z.FullName
                              select z.FullName
                               ).Distinct().ToList();
            DisplayStringList(distPlayer);

            int EntryMax = distPlayer.Count;
            string listSelection = FindSelectedEntry(distPlayer, EntryMax);

            SelPlayerPlayerList(listSelection);
        }

        //************************************************************************
        // Create collection for specific player
        //************************************************************************
        private static void SelPlayerPlayerList(string input)
        {
            Console.WriteLine("Searching for players from " + input);
            //object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(ExcelArray);
            List<Season> nameCollection = playerCollection.Where(s => s.FullName == input).ToList();

            var distPlayer = (from z in nameCollection
                              orderby z.FullName, z.Year, z.Team
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

        //************************************************************************
        // Seach by College processing - Dennis
        //************************************************************************
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
                AllCollegeList();
            }
            else
            {
                SelCollegePlayerList(input);
            }
        }

        //************************************************************************
        // Create collection of unique colleges
        //************************************************************************
        private static void AllCollegeList()
        {
            Console.WriteLine("Searching for colleges...");

            //object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(ExcelArray);

            var distCollege = (from z in playerCollection
                               orderby z.College
                               select z.College
                               ).Distinct().ToList();

            DisplayStringList(distCollege);
            int EntryMax = distCollege.Count;
            string listSelection = FindSelectedEntry(distCollege, EntryMax);
            SelCollegePlayerList(listSelection);
        }

        //************************************************************************
        // Create collection of players for specific college
        //************************************************************************
        private static void SelCollegePlayerList(string input)
        {
            Console.WriteLine("Searching for players from " + input);

            //object[,] playerArray = SetUpExcel();
            List<Season> playerCollection = LoadCollection(ExcelArray);
            List<Season> nameCollection = playerCollection.Where(s => s.College == input).ToList();

            var distPlayer = (from z in nameCollection
                              orderby z.FullName, z.Year, z.Team
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

        //************************************************************************
        // Create Load collection from spreadsheet
        //************************************************************************
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

        //************************************************************************
        // EXCEL path definition
        //************************************************************************
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

        //************************************************************************
        // Close EXCEL
        //************************************************************************
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

        //************************************************************************
        // Display player search results
        //************************************************************************
        public static void DisplayResults(List<Season> displayCollection)
        {
            int skipCount = 0;
            int pageSize = 25;
            int takeCount = pageSize;
            int pageCount = 1;
            int totCount = displayCollection.Count;
            int pageTot = (totCount + (pageSize - 1)) / pageSize; // determine number of pages needed to display list
            Console.WriteLine("Number of results " + totCount + " / page count " + pageTot);
            Console.WriteLine(" ");
            Console.WriteLine("Position\tPlayer Name\t\tTeam\tYear\tPassing Yds\tRushing Yds\tCollege");
            do
            {
                foreach (Season selectedPosition in displayCollection.Skip<Season>(skipCount).Take<Season>(takeCount))
                {
                    Console.WriteLine("{0,5}\t" + "   " + "{1,-25}\t{2,3}\t{3,4}\t{4,8:n0}\t{5," + "" + "8:n0}\t{6,-20}",
                        selectedPosition.Position,
                        selectedPosition.FullName,
                        selectedPosition.Team,
                        selectedPosition.Year,
                        selectedPosition.PassingYards,
                        selectedPosition.RushingYards,
                        selectedPosition.College);
                }
                pageCount = PageControl(ref skipCount, pageSize, totCount, pageTot, "Y");
            } while (skipCount < displayCollection.Count);
            Console.WriteLine("Done with list");
        }

        //************************************************************************
        // Paging logic
        //************************************************************************
        private static int PageControl(ref int skipCount, int pageSize, int totCount, int pageTot, string headerList)
        {
            int pageCount = (skipCount / pageSize) + 1;
            Console.WriteLine("Display Page " + pageCount + " of " + pageTot + " 'F' Forward / 'B' Backward ' 'X' Exit or Specific Page");
            string pageInput = Console.ReadLine().ToUpper();

            if (pageInput == "F")
            {
                skipCount += pageSize;
                if (headerList == "Y")
                {
                    Console.WriteLine("Position\tPlayer Name\t\tTeam\tYear\tPassing Yds\tRushing Yds");
                }
            }
            else if (pageInput == "B")
            {
                skipCount -= pageSize;
                if (skipCount < 0)
                {
                    skipCount = 0;
                }
                if (headerList == "Y")
                {
                    Console.WriteLine("Position\tPlayer Name\t\tTeam\tYear\tPassing Yds\tRushing Yds");
                }
            }
            else if (pageInput == "X")
            {
                skipCount = totCount;
            }
            else if (pageInput.All(char.IsDigit))
            {
                // ???????????????????????????????????????????
                // ??? add logic to check for non-valid chars
                // ???????????????????????????????????????????
                int reqPage = int.Parse(pageInput);
                if (reqPage > pageTot)
                {
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    Console.WriteLine("Entered page greater than number of pages in list");
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else if (reqPage == 0)
                {
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    Console.WriteLine("Entered page must be greater than zero");
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else
                    skipCount = (reqPage * pageSize) - pageSize;
            }
            pageCount = (skipCount * pageSize) - 1;
            return pageCount;
        }

        //************************************************************************
        // Seach list results
        //************************************************************************
        public static void DisplayStringList(List<string> displayCollection)
        {
            int skipCount = 0;
            int pageSize = 25;
            int takeCount = pageSize;
            int printCount = 0;
            int pageCount = 1;
            int totCount = displayCollection.Count;
            int pageTot = (totCount + (pageSize - 1)) / pageSize; // determine number of pages needed to display list
            Console.WriteLine("Number of results " + totCount + " / page count " + pageTot);
            Console.WriteLine(" ");
            do
            {
                if (pageCount == 1)
                {
                    printCount = 1;
                }
                else
                {
                    printCount = skipCount + 1;
                };

                foreach (string selectedString in displayCollection.Skip<string>(skipCount).Take<string>(takeCount))
                {
                    Console.WriteLine(printCount + ". " + selectedString);
                    printCount += 1;
                }
                pageCount = PageControl(ref skipCount, pageSize, totCount, pageTot, "N");
            } while (skipCount < displayCollection.Count);
            Console.WriteLine("Done with list");
        }

        //************************************************************************
        // Find selected entry from numbered list
        //************************************************************************
        private static string FindSelectedEntry(List<string> distList,int EntryMax)
        {
            Console.WriteLine("Enter number of desired entry");
            string SelInput = Console.ReadLine();
            int reqSel = 0;
            bool goodSel = false;

            do
            {
                reqSel = int.Parse(SelInput);
                if (SelInput == "0")
                {
                    Console.WriteLine("Entry must be greater than zero");
                    SelInput = Console.ReadLine();
                }
                else if (reqSel > EntryMax)
                {
                    Console.WriteLine("Entry cannot be greater than " + EntryMax);
                    SelInput = Console.ReadLine();
                }
                else if (reqSel <= EntryMax & reqSel > 0)
                {
                    goodSel = true;
                }

            } while (goodSel == false);

            //while (String.IsNullOrEmpty(SelInput))
            //{
            //    Console.WriteLine("Please enter number of desired entry");
            //    SelInput = Console.ReadLine();
            //}
            //reqSel = int.Parse(SelInput);

            //if (SelInput == "0")
            //{
            //    Console.WriteLine("Entry must be greater than zero");
            //    SelInput = Console.ReadLine();
            //}
            //else if (reqSel > EntryMax)
            //{
            //    Console.WriteLine("Entry cannot be greater than "+EntryMax);
            //    SelInput = Console.ReadLine();
            //}

            int skipCount = reqSel - 1;
            int takeCount = 1;
            string listSelection = " ";
            do
            {
                foreach (string selectedString in distList.Skip<string>(skipCount).Take<string>(takeCount))
                {
                    Console.WriteLine("Searching for results for selected entry " + selectedString);
                    listSelection = selectedString;
                }
            } while (skipCount < 1);
            return listSelection;
        }

        //************************************************************************
        // Validate EXCEL data
        //************************************************************************
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

        //************************************************************************
        // Load collection - season
        //************************************************************************
        public static Season AddSeason(int currentRow, object[,] ExcelArray)
        {
            string[] parseFullNameArray = ExcelArray[currentRow, (int)SpreadsheetColumn.playerColumn].ToString().Split(new char[] { ' ' });
            Season season = new Season()
            {
                Year = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.yearColumn]),
                FullName = parseFullNameArray[(int)FullNameBreakdown.lastNameIndex] + ", " + parseFullNameArray[(int)FullNameBreakdown.firstNameIndex],
                Team = ExcelArray[currentRow, (int)SpreadsheetColumn.teamColumn].ToString(),
                PassingYards = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.passingYardsColumn]),
                RushingYards = Convert.ToDouble(ExcelArray[currentRow, (int)SpreadsheetColumn.rushingYardsColumn]),
                Position = ExcelArray[currentRow, (int)SpreadsheetColumn.positionColumn].ToString(),
                College = ExcelArray[currentRow, (int)SpreadsheetColumn.collegeColumn].ToString()
            };
            return season;
        }
        
        //************************************************************************
        // Validate Position
        //************************************************************************
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

        //************************************************************************
        // Chris - play 
        //************************************************************************
        public static IEnumerable<string> Search(IEnumerable<string> data, string q)
        {
            string regexSearch = q
                .Replace("*", ".+")
                .Replace("%", ".+")
                .Replace("#", "\\d")
                .Replace("@", "[a-zA-Z]")
                .Replace("?", "\\w");

            Regex regex = new Regex(regexSearch);

            return data
                .Where(s => regex.IsMatch(s));
        }
    }

}
