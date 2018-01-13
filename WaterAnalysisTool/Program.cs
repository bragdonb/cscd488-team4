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
        /* The ~~real~~ main */
        //TODO change so files can be opened without needing to type their extension
        //TODO handle filenames with spaces (surrounded by "")
        static void Main(string[] args)
        {
            // The functionality of main:
            // 1. Awaits input in from user
            //  1.1. Accepts a command to parse the ICP-AES file (parse <location/name of input> <location/name for output>)
            //      1.1.1. Create a new ExcelPackage
            //      1.1.2. Create each necessary worksheet in the package (Data, Calibration Standards, Graphs)
            //      1.1.3. Set the title in the packages properties to the name of the output file (sans the extension)
            //      1.1.4. Create a new DataLoader and call its load function
            //  1.2. Accepts a command to create correlation matrices (analyze <location/name of input> <r^2 threshold>)

            String stringArgs = null;
            FileInfo infile = null, outfile;
            Double r2val, threshold;

            do
            {
                try
                {
                    Console.Write("Enter command: ");
                    stringArgs = Console.ReadLine();

                    if (stringArgs.ToLower().Equals("usage"))
                        Console.WriteLine("parse <location/name of input> <location/name for output>\nanalyze <location/name of input> <r^2 threshold>");

                    #region Testing
                    else if (stringArgs.ToLower().Equals("test loader"))
                    {
                         SampleGroup CalibrationSamples;
                        SampleGroup CalibrationStandards;
                        SampleGroup QualityControlSamples;
                        SampleGroup CertifiedValueSamples_1;
                        SampleGroup CertifiedValueSamples_2;
                        SampleGroup Samples_1;
                        SampleGroup Samples_2;

                        Sample s = null;
                        List<Sample> list = new List<Sample>();

                        Random random = new Random();

                        FileInfo fi = new FileInfo(@"tester.xlsx");
                        if (fi.Exists)
                            fi.Delete();

                        using (var p = new ExcelPackage(new FileInfo(@"tester.xlsx")))
                        {     
                            p.Workbook.Properties.Title = "Title of Workbook";
                            p.Workbook.Worksheets.Add("Data");
                            p.Workbook.Worksheets.Add("Calibration Standards");
                            p.Workbook.Worksheets.Add("Graphs");


                            DataLoader loader = new DataLoader(null, p);

                            #region Sample & Sample Group Creation
                            for(int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Calibration Sample #" + i, DateTime.Now.ToString(), "QC", 3);

                                for(int j = 0; j < 10; j++)
                                    s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);
                            }

                            CalibrationSamples = new SampleGroup(list, "CalibrationSamples", false);
                            list.Clear();

                            for (int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Calibration Sample #" + i, DateTime.Now.ToString(), "QC", 3);

                                for(int j = 0; j < 10; j++)
                                s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);
                            }

                            CalibrationStandards = new SampleGroup(list, "CalibrationSamples", false);
                            list.Clear();

                            for (int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Quality Control Sample #" + i, DateTime.Now.ToString(), "QC", 3);

                                for (int j = 0; j < 10; j++)
                                    s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);

                            }

                            QualityControlSamples = new SampleGroup(list, "QualityControlSamples", true);
                            list.Clear();

                            for(int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Certified Value (1) Sample #" + i, DateTime.Now.ToString(), "QC", 3);

                                for (int j = 0; j < 10; j++)
                                    s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);
                            }

                            CertifiedValueSamples_1 = new SampleGroup(list, "CertifiedValueSamples_1", true);
                            list.Clear();

                            for(int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Certified Value (2) Sample #" + i, DateTime.Now.ToString(), "QC", 3);

                                for (int j = 0; j < 10; j++)
                                    s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);
                            }

                            CertifiedValueSamples_2 = new SampleGroup(list, "CertifiedValueSamples_2", true);
                            list.Clear();

                            for(int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Generic (1) Sample #" + i, DateTime.Now.ToString(), "Unk", 3);

                                for (int j = 0; j < 10; j++)
                                    s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);
                            }

                            Samples_1 = new SampleGroup(list, "Samples_1", false);
                            list.Clear();

                            for(int i = 0; i < 10; i++)
                            {
                                s = new Sample("Method Name", "Generic (2) Sample #" + i, DateTime.Now.ToString(), "QC", 3);

                                for (int j = 0; j < 10; j++)
                                    s.AddElement(new Element("Elem. #" + j, "Units", j * random.NextDouble(), 1.0, 1.0));

                                list.Add(s);
                            }

                            Samples_2 = new SampleGroup(list, "Samples_2", false);
                            list.Clear();

                            loader.AddCalibrationSampleGroup(CalibrationSamples);
                            loader.AddCalibrationStandard(CalibrationStandards);
                            loader.AddQualityControlSampleGroup(QualityControlSamples);
                            loader.AddCertifiedValueSampleGroup(CertifiedValueSamples_1);
                            loader.AddCertifiedValueSampleGroup(CertifiedValueSamples_2);
                            loader.AddSampleGroup(Samples_1);
                            loader.AddSampleGroup(Samples_2);
                            #endregion

                            loader.Load(); // Load calls Parse, don't need to in main; also saves the workbook
                        }
                    }
                    #endregion

                    else
                    {
                        Regex r = new Regex(@"("".*?"")|(\S+)");
                        MatchCollection arguments = r.Matches(stringArgs);

                        if (arguments.Count > 1)
                            infile = new FileInfo(@arguments[1].Value);

                        if (infile.Exists)
                        {
                            if (arguments[0].Value.ToLower().Equals("parse"))
                            {
                                if (arguments.Count > 2)
                                {
                                    outfile = new FileInfo(@arguments[2].Value);
                                    if (outfile.Exists)
                                        outfile.Delete();

                                    using (ExcelPackage p = new ExcelPackage(new FileInfo(@arguments[2].Value)))
                                    {
                                        p.Workbook.Properties.Title = arguments[2].Value.Split('.')[0];
                                        p.Workbook.Worksheets.Add("Data");
                                        p.Workbook.Worksheets.Add("Calibration Standards");
                                        p.Workbook.Worksheets.Add("Graphs"); //maybe rename

                                        DataLoader loader = new DataLoader(infile.OpenText(), p);
                                        loader.Load();
                                    }
                                }
                            }

                            else if (arguments[0].Value.ToLower().Equals("analyze"))
                            {
                                threshold = 0.7;
                                if (arguments.Count > 2)
                                {
                                    if (Double.TryParse(arguments[2].Value, out r2val))
                                    {
                                        if (r2val <= 1 && r2val >= 0)
                                            threshold = r2val;
                                        else
                                            threshold = -1;
                                    }

                                    else
                                        threshold = -1;
                                }

                                if (threshold != -1)
                                {
                                    //threshold now has correct value
                                    //Console.WriteLine("threshold is not -1, it is " + threshold);
                                    using (ExcelPackage p = new ExcelPackage(infile))//TODO see if this works with a file that isn't an xlsx file
                                    {
                                        AnalyticsLoader analyticsLoader = new AnalyticsLoader(p, threshold);
                                        analyticsLoader.Load();
                                    }
                                }
                            }

                        }//end if(infile.Exists)

                        else
                        {
                            Console.WriteLine("Input file does not exist.");
                        }
                    }
                }

                catch(Exception e)//TODO make these messages more specific so that she knows exactly what went wrong
                {
                    Console.WriteLine(e.GetType() + " " + e.Message) ;
                }

            } while (!stringArgs.ToLower().Equals("exit"));

            Console.WriteLine("Exiting...");
        }
    }
}
