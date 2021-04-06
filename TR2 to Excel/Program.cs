using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace TR2_to_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("WHATEVER YOU DO, DON'T CLOSE THE APPLICATION BEFORE THE PROCESS IS FINISHED, THERE WILL EXCEL SHEETS OPEN, WHICH IS VERY INCONVENIENT AND CLOGS UP SYSTEM RESOURCES UNTIL CLOSED. IF YOU CRASH IT, ELEMINATE ALL INSTANCES OF EXCEL THROUGH THE TASK MANAGER.");

            string[] fileLocations = GetFileLocation();

            ResetCurrentLine();
            Console.WriteLine(@"Type save location (example: C:\Users\Public\Testfolder\EXCELSHEET.xlsx) NOTE: THE FILETYPE MUST BE .xlsx");
            string savePath = Console.ReadLine();

            Console.WriteLine("");
            List<Tab> tabs = new List<Tab>();
            for (int i = 0; i < fileLocations.Length; i++)
            {
                Feedback(i, fileLocations.Length);

                string[] allLines = GetLines(fileLocations[i]);

                tabs.Add(GetLabel(allLines));

                string[] goodLines = ReturnLinesBetweenStrings(allLines, "\"\";\"\";\"\";\"s\";\"mm\";\"N\";\"mm\"", "[ChannelParameters]");

                List<Array> tempArray = new List<Array>();
                for (int j = 0; j < goodLines.Length; j++)
                {
                    tempArray.Add(SeperateIntoDoubleElements(goodLines[j], "-?\\ *[0-9]+\\.?[0-9]*(?:[Ee]\\ *[-|+]?\\ *[0-9]+)?"));
                    Feedback(i, fileLocations.Length, j, goodLines.Length, "good lines.");
                }

                tabs[i].arrays = tempArray.ToArray();
            }

            Console.WriteLine("Done");
            Console.WriteLine("Now to write everything to an Excel file...");
            Console.WriteLine("");

            WriteToExcel(tabs.ToArray(), savePath);
        }

        private static string[] GetFileLocation()
        {
            List<string> fileLocations = new List<string>();
            while (true)
            {
                Console.WriteLine(@"Type the file location (example: C:\Users\Public\Testfolder\FILE.TR2)");
                string input = Console.ReadLine();
                if (!fileLocations.Contains(input) || !File.Exists(input))
                {
                    Console.WriteLine(string.Format("{0} added. Do you want to add more? Type \"y\" to add more. Type \"c\" to remove the last entry. And type \"n\" to process data.", input));
                    switch (Console.ReadLine())
                    {
                        case "y":
                            fileLocations.Add(input);
                            break;
                        case "c":
                            break;
                        default:
                            fileLocations.Add(input);
                            return fileLocations.ToArray();
                    }
                }
                else
                {
                    Console.WriteLine(string.Format("{0} was already added, or is invalid, latest entry not added.\nDo you want to add more? Type \"y\" to add more. And type \"n\" to process data.", input));
                    switch (Console.ReadLine())
                    {
                        case "y":
                            break;
                        default:
                            return fileLocations.ToArray();
                    }
                }

            }
        }

        private static string[] GetLines(string location)
        {
            return File.ReadAllLines(location);
        }

        private static void Feedback(int currentLocation, int totalNumber)
        {
            ResetCurrentLine();
            string output = "";
            for (int i = 0; i < totalNumber; i++)
            {
                if (currentLocation >= i)
                {
                    output += "█";
                }
                else 
                {
                    output += "░";
                }
            }
            output += string.Format(" {0} out of {1}", currentLocation + 1, totalNumber);
            Console.Write(output);
        }

        private static void Feedback(int currentLocation, int totalNumber, int currentGoodLines, int goodLines, string itemName)
        {
            ResetCurrentLine();
            string output = "";
            for (int i = 0; i < totalNumber; i++)
            {
                if (currentLocation >= i)
                {
                    output += "█";
                }
                else
                {
                    output += "░";
                }
            }
            output += string.Format(" {0} out of {1} and {2} out of {3} {4}", currentLocation + 1, totalNumber, currentGoodLines + 1, goodLines, itemName);
            Console.Write(output);
        }

        private static void ResetCurrentLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        private static string[] ReturnLinesBetweenStrings(string[] input, string beginString, string endsTring)
        {
            List<string> output = new List<string>();

            bool active = false;

            for (int i = 0; i < input.Length; i++)
            {
                string line = input[i];
                if (line == beginString)
                {
                    active = true;
                } else if (line == endsTring)
                {
                    active = false;
                } else if (active == true && line != "")
                {
                    output.Add(line);
                }
            }

            return output.ToArray();
        }

        //"  (.*);  (.*);  (.*);  (.*);  (.*)"
        private static Array SeperateIntoDoubleElements(string line, string pattern)
        {
            List<double> numbers = new List<double>();

            foreach (Match match in Regex.Matches(line, pattern))
            {
                numbers.Add(double.Parse(match.Value.ToString(), CultureInfo.InvariantCulture));
            }
            return new Array { doubles = numbers.ToArray() };
        }

        private static Tab GetLabel(string[] fileLines)
        {
            Tab output = new Tab();
            string labelLine = "";
            foreach (string line in fileLines)
            {
                if (line.Contains("Label"))
                {
                    labelLine = line;
                }
            }
            foreach (Match match in Regex.Matches(labelLine, "(?<=\\\")([A-Za-z]*?)(?=\\\")"))
            {
                if (match.Value != "Label")
                {
                    output.name = match.Value;
                    break;
                }
            }
            output.elements = GetElements();
            return output;
        }

        private static string[] GetElements()
        {
            return new string[] { "", "", "", "s", "mm", "N", "mm" };
        }

        private static void WriteToExcel(Tab[] tabs, string savePath)
        {
            Application excelApp = new Application();
            Workbook excelWorkbook = excelApp.Workbooks.Add(Type.Missing);
            Sheets excelSheets = excelWorkbook.Sheets;
            for (int i = 0; i < tabs.Length; i++)
            {
                Feedback(i, tabs.Length);
                Worksheet excelWorksheet = (Worksheet)excelSheets.Add(excelSheets[i + 1]);
                excelWorksheet.Name = tabs[i].name;
                for (int j = 0; j < tabs[i].elements.Length; j++)
                {
                    excelWorksheet.Cells[1, j + 1] = tabs[i].elements[j];
                }

                for (int j = 0; j < tabs[i].arrays.Length; j++)
                {
                    for (int k = 0; k < tabs[i].arrays[j].doubles.Length; k++)
                    {
                        excelWorksheet.Cells[j + 2, k + 1] = tabs[i].arrays[j].doubles[k];
                    }

                    Feedback(i, tabs.Length, j, tabs[i].arrays.Length, "lines.");
                }
                excelWorksheet.Columns.AutoFit();
            }

            Console.WriteLine("");
            Console.WriteLine("Now saving...");
            File.Delete(savePath);
            excelWorkbook.SaveAs(savePath);
            excelWorkbook.Close(0);
            excelApp.Quit();
        }
    }

    public class Array
    {
        public double[] doubles;
    }
    
    public class Tab
    {
        public string name;
        public string[] elements;
        public Array[] arrays;
    }
}
