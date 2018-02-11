using OfficeOpenXml; // This is the top namespace for EPPlus, if your reference isn't found use the command -> Update-Package -reinstall in the NuGet Console
using System;
using System.IO;
using System.Collections.Generic;
using WaterAnalysisTool.Loader;
using WaterAnalysisTool.Components;
using System.Text.RegularExpressions;
using WaterAnalysisTool.Analyzer;

namespace WaterAnalysisTool
{
    class Program
    {
        public const double threshold = 0.7;
        public const double version = 1.0;

        public static void Main(string[] args)
        {
            // The functionality of main:
            // 1. Awaits input in from user
            //  1.1. Accepts a command to parse the ICP-AES file (parse <location/name of input> <location/name for output>)
            //      1.1.1. Create a new ExcelPackage
            //      1.1.2. Set the title in the packages properties to the name of the output file (sans the extension)
            //      1.1.3. Create a new DataLoader and call its load function
            //  1.2. Accepts a command to create correlation matrices (analyze <location/name of input> <r^2 threshold>)

            FileInfo infile = null, outfile = null; ;
            Double r2val;
            bool flag;

            String stringArgs = null;
            Regex r = new Regex("[^\\s\"']+|\"([^\"]*)\"|'([^']*)'");
            MatchCollection arguments = null;

            // Startup Message
            Console.WriteLine("ICP-AES Text File Parser version " + version + ".\nType \"usage\" for a list of commands.\n");

            do
            {
                try
                {
                    // Reseting
                    flag = false;
                    stringArgs = null;
                    arguments = null;

                    Console.Write("Enter command: ");
                    stringArgs = Console.ReadLine();

                    if (stringArgs.ToLower().Equals("usage"))
                        Console.WriteLine("\tparse <location/name of input> <location/name for output>\n\tanalyze <location/name of input> <r^2 threshold>\n\tType \"exit\" to exit.");

                    else
                    {
                        // Check if stringArgs matches expected command structure
                        arguments = r.Matches(stringArgs);

                        if (arguments.Count > 1)
                        {
                            #region Parse Command
                            if (arguments[0].Value.ToLower().Equals("parse"))
                            {
                                if (arguments.Count > 2)
                                {
                                    // Input file cleaning
                                    String file = arguments[1].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes
                                    if (!file.Contains(".")) // If it has no extension, add ".txt"
                                        file = file + ".txt";
                                    infile = new FileInfo(file);

                                    // Output file cleaning
                                    file = arguments[2].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes
                                    if (!file.Contains(".")) // If it has no extension, add ".xlsx"
                                        file = file + ".xlsx";
                                    outfile = new FileInfo(file);

                                    if (infile.Exists)
                                    {
                                        if (outfile.Exists)
                                        {
                                            Console.WriteLine("\tA file of the name " + outfile.Name + " already exists at " + (outfile.ToString().Substring(0, outfile.Name.Length)) + ".");
                                            Console.Write("\tThis operation will overwrite this file. Continue? (y/n): ");

                                            if (Console.ReadLine().ToLower().Equals("n"))
                                            {
                                                Console.WriteLine("\tParse operation cancelled.");
                                            }

                                            else
                                            {
                                                outfile.Delete();

                                                using (ExcelPackage p = new ExcelPackage(new FileInfo(@arguments[2].Value)))
                                                {
                                                    p.Workbook.Properties.Title = arguments[2].Value.Split('.')[0];

                                                    DataLoader loader = new DataLoader(infile.OpenText(), p);
                                                    loader.Load();
                                                }
                                            }
                                        }

                                        else
                                        {
                                            using (ExcelPackage p = new ExcelPackage(new FileInfo(@arguments[2].Value)))
                                            {
                                                p.Workbook.Properties.Title = arguments[2].Value.Split('.')[0];

                                                DataLoader loader = new DataLoader(infile.OpenText(), p);
                                                loader.Load();
                                            }
                                        }
                                    }

                                    else
                                        Console.WriteLine("\tCould not locate " + infile.ToString());
                                }

                                else
                                    Console.WriteLine("\t" + stringArgs + " is an invalid command. For a list of valid commands enter \"usage\".");
                            }
                            #endregion

                            #region Analyze Command
                            else if (arguments[0].Value.ToLower().Equals("analyze"))
                            {
                                String file = arguments[1].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes
                                if (!file.Contains(".")) // If it has no extension, add ".xlsx"
                                    file = file + ".xlsx";
                                infile = new FileInfo(file);

                                if (infile.Exists)
                                {
                                    if (arguments.Count > 2)
                                    {
                                        // Optional threshold argument entered
                                        if (Double.TryParse(arguments[2].Value, out r2val))
                                        {
                                            if (r2val >= 0.0 && r2val <= 1)
                                            {
                                                using (ExcelPackage p = new ExcelPackage(infile))
                                                {
                                                    foreach (ExcelWorksheet sheet in p.Workbook.Worksheets)
                                                    {
                                                        // Check if correlation worksheet already exists
                                                        if(sheet.Name.Equals("Correlation"))
                                                        {
                                                            Console.WriteLine("\tA correlation worksheet already exists for this file.");
                                                            Console.Write("\tThis operation will overwrite it. Continue? (y/n): ");

                                                            if (Console.ReadLine().ToLower().Equals("n"))
                                                            {
                                                                Console.WriteLine("\tAnalyze operation cancelled.");
                                                                flag = true;
                                                                break;
                                                            }

                                                            else
                                                            {
                                                                p.Workbook.Worksheets.Delete(sheet);
                                                                break;
                                                            }
                                                        }
                                                    }

                                                    if (!flag)
                                                    {
                                                        AnalyticsLoader analyticsLoader = new AnalyticsLoader(p, threshold);
                                                        analyticsLoader.Load();
                                                    }
                                                }
                                            }

                                            else
                                                Console.WriteLine("\t" + arguments[2] + " is an invalid threshold. Threshold must be a value between 0 and 1 inclusive.");
                                        }

                                        else
                                            Console.WriteLine("\t" + arguments[2] + " is an invalid threshold. Threshold must be numeric and a value between 0 and 1 inclusive.");
                                    }

                                    else
                                    {
                                        using (ExcelPackage p = new ExcelPackage(infile))
                                        {
                                            foreach (ExcelWorksheet sheet in p.Workbook.Worksheets)
                                            {
                                                // Check if correlation worksheet already exists
                                                if (sheet.Name.Equals("Correlation"))
                                                {
                                                    Console.WriteLine("\tA correlation worksheet already exists for this file.");
                                                    Console.Write("\tThis operation will overwrite it. Continue? (y/n): ");

                                                    if (Console.ReadLine().ToLower().Equals("n"))
                                                    {
                                                        Console.WriteLine("\tAnalyze operation cancelled.");
                                                        flag = true;
                                                        break;
                                                    }

                                                    else
                                                    {
                                                        p.Workbook.Worksheets.Delete(sheet);
                                                        break;
                                                    }
                                                }
                                            }

                                            if (!flag)
                                            {
                                                AnalyticsLoader analyticsLoader = new AnalyticsLoader(p, threshold);
                                                analyticsLoader.Load();
                                            }
                                        }
                                    }
                                }

                                else
                                    Console.WriteLine("\tCould not locate " + infile.ToString());
                            }
                            #endregion

                            else
                                Console.WriteLine("\t" + stringArgs + " is an invalid command. For a list of valid commands enter \"usage\".");
                        }

                        else
                            Console.WriteLine("\t" + stringArgs + " is an invalid command. For a list of valid commands enter \"usage\".");
                    }
                }

                // Ideally, we would never get here...
                catch(Exception e)
                {
                    //Console.WriteLine("\t" + e.Message); // Uncomment this when done and get rid of stuff below
                    Console.WriteLine("\t" + e.GetType() + " " + e.Message);
                    Console.WriteLine("\t" + e.ToString());
                }

                Console.WriteLine(); // Some formatting
            } while (!stringArgs.ToLower().Equals("exit"));

            Console.WriteLine("Exiting...");
        }
    }
}
