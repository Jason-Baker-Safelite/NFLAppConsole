using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using NFL;

namespace NFLAppConsole
{
    public class NFLSearch
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range MyRange = null;
        public  object[,] ExcelArray { get; set; } //SetUpExcel();

        

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



        public void LoadUp()
        {
            ExcelArray = SetUpExcel();
        }
        //************************************************************************
        // Close EXCEL
        //************************************************************************
        public void Cleanup()
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
        // EXCEL path definition
        //************************************************************************
        private  object[,] SetUpExcel()
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
        // Validate EXCEL data
        //************************************************************************
        private  object[,] ValidateExcelData(int spreadsheetRow, object[,] parmArray)
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
        private  Season AddSeason(int currentRow, object[,] ExcelArray)
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

        public  bool PrintConsoleMenu(bool Printloop)
        {
            Console.WriteLine("Select Menu Item");
            Console.WriteLine("1. Search by Player");
            Console.WriteLine("2. Search by Team");
            Console.WriteLine("3. Search by Position");
            Console.WriteLine("4. Search by Year");
            Console.WriteLine("5. Search by College");
            Console.WriteLine("6. Exit");

            string userInput = Console.ReadLine();
            //if (userInput == "1")
            //{
            //    Console.WriteLine("You selected Search by Player");
            //    SearchByPlayer(userInput);
            //}
            //else if (userInput == "2")
            //{
            //    Console.WriteLine("You selected Search by Team");
            //    SearchByTeam(userInput);
            //}
            //else if (userInput == "3")
            //{
            //    Console.WriteLine("You selected Search by Position");
            //    Console.WriteLine("Please enter 2-character position: ");
            //    string positionRequest = ValidatePositionInput();
            //    List<Season> positionResultsCollection = SearchByPosition(positionRequest);
            //    if (positionResultsCollection.Count == 0)
            //    {
            //        Console.WriteLine("No entries found for position " + positionRequest);
            //    }
            //    else
            //    {
            //        DisplayResults(positionResultsCollection);
            //    }
            //}
            //else if (userInput == "4")
            //{
            //    Console.WriteLine("You selected Search by Year");
            //    SearchByYear(userInput);
            //}
            //else if (userInput == "5")
            //{
            //    Console.WriteLine("You selected Search by College");
            //    SearchByCollege(userInput);
            //}
            //else if (userInput == "6")
            //{
            //    Console.WriteLine("Goodbye");
            //    Printloop = false;
            //}
            //else
            //{
            //    Console.WriteLine("Invalid Menu Selection");
            //    //                PrintMenu();
            //}
            return Printloop;
        }
    }
}
