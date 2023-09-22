using logging;
using mdbtocsv_util;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using CsvHelper;
using System.Globalization;
using System.Text.RegularExpressions;

namespace mdbtocsv
{
    internal class mdbtocsv
    {
        public enum QuoteIdentifier
        {
            None, DoubleQuote
        }

        public enum CSVDelimiter
        { 
            comma, tab, pipe
        }

        public enum FileNameCase
        { 
            none, lower, upper
        }

        private static string FileToProcess { get; set; }
        private static bool GenerateLogFile { get; set; }
        private static bool DEBUGMODE { get; set; }
        private static bool AllowOverWrite { get; set; }
        private static string OutputDirectory { get; set; }
        private static CSVDelimiter DelimiterToUse { get; set; }
        private static bool CleanFieldNames { get; set; }
        private static FileNameCase FileNameCaseToUse { get; set; }
        private static int ExitCodeStatus { get; set; }
        private static bool AppendCreateDateToOutputFiles { get; set; }

        static void Main(string[] args)
        {
            InitApplicationVariables();

            Assembly thisAssem = typeof(mdbtocsv).Assembly;
            AssemblyName thisAssemName = thisAssem.GetName();
            Version ver = thisAssemName.Version;

            System.IO.FileInfo fi = new System.IO.FileInfo(thisAssem.Location);
            OutputDirectory = fi.DirectoryName;

            var log_file = fi.DirectoryName + Path.DirectorySeparatorChar + "mdbtocsv_log.txt";
            Debug.Print($"LogFilename: {log_file}");
            Log.Init(log_file);
            Log.TrimLogFile();

            Console.WriteLine($"********** MDB To CSV {ver.Major}.{ver.Minor}.{ver.Build} **********");
            Log.WriteToLogFile($"********** MDB To CSV {ver.Major}.{ver.Minor}.{ver.Build} **********");
            Console.WriteLine($"*");
            Console.WriteLine($"* BASE2 Software - https://bentonvillebase2.com");
            Console.WriteLine($"*");
            Console.WriteLine($"************************************");

            // process all command line parameters
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (DEBUGMODE)
                    {
                        Log.WriteToLogFile($"Arg {i}: '{args[i].ToLower()}'", true);
                    }

                    if (args[i].ToLower() == "-nolog")
                    {
                        Console.WriteLine("* startup param: disabling log...");
                        GenerateLogFile = false;
                    }
                    else if (args[i].ToLower().StartsWith("-s:"))
                    {
                        string fileName = args[i].Substring(3);
                        if (File.Exists(fileName))
                        {
                            Log.WriteToLogFile($"* startup param: SOURCE_FILENAME={fileName.ToLower()}", true);
                            FileToProcess = fileName;
                            OutputDirectory = Path.GetDirectoryName(fileName);
                        }
                        else
                        {
                            Log.WriteToLogFile($"* startup param: ERROR. Invalid Source File '{fileName.ToLower()}' ...", true);
                        }
                    }
                    else if (args[i].ToLower().StartsWith("-o:"))
                    {
                        string outputDirectory = args[i].Substring(3);
                        if (Directory.Exists(outputDirectory))
                        {
                            Log.WriteToLogFile($"* startup param: Output Folder Updated To '{outputDirectory.ToLower()}' ...", true);
                            OutputDirectory = outputDirectory;
                        }
                        else
                        {
                            Log.WriteToLogFile($"* startup param: ERROR. Invalid Output Directory '{outputDirectory.ToLower()}' ...", true);
                            Log.WriteToLogFile($"* using default path. {OutputDirectory}", true);
                        }
                    }
                    else if (args[i].ToLower() == "-nooverwrite")
                    {
                        Log.WriteToLogFile("* startup param: disabling over-write of output file.", true);
                        AllowOverWrite = false;
                    }
                    else if (args[i].ToLower() == "-lower")
                    {
                        Log.WriteToLogFile("* startup param: output filenames will be lowercase.", true);
                        FileNameCaseToUse = FileNameCase.lower;
                    }
                    else if (args[i].ToLower() == "-upper")
                    {
                        Log.WriteToLogFile("* startup param: output filenames will be UPPERCASE.", true);
                        FileNameCaseToUse = FileNameCase.upper;
                    }
                    else if (args[i].ToLower() == "-debug")
                    {
                        Log.WriteToLogFile("* startup param: DEBUG Mode Enabled By Startup Parameter.", true);
                        DEBUGMODE = true;
                        GenerateLogFile = true;
                    }
                    else if (args[i].ToLower() == "-t")
                    {
                        Log.WriteToLogFile("* startup param: Using TAB Delimiter.", true);
                        DelimiterToUse = CSVDelimiter.tab;
                    }
                    else if (args[i].ToLower() == "-p")
                    {
                        Log.WriteToLogFile("* startup param: Using PIPE Delimiter.", true);
                        DelimiterToUse = CSVDelimiter.tab;
                    }
                    else if (args[i].ToLower() == "-c")
                    {
                        Log.WriteToLogFile("* startup param: Field Name Cleanup ENABLED.", true);
                        CleanFieldNames = true;
                    }
                    else if (args[i].ToLower() == "-adddate")
                    {
                        Log.WriteToLogFile("* startup param: Append File Create date to output file name option ENABLED.", true);
                        AppendCreateDateToOutputFiles = true;
                    }
                    else if (args[i].ToLower().Contains("?") || args[i].ToLower().Contains("-help"))
                    {
                        Console.WriteLine();
                        Console.WriteLine("## BASE2 Software MDBTOCSV Command Line Options ##");
                        Console.WriteLine("USAGE:   mdbtocsv -s:\"<sourceFileName>\" -noprompt -lower");
                        Console.WriteLine("The above example will extract all tables in source file as separate CSVs");

                        Console.WriteLine("");
                        Console.WriteLine("OPTIONS:");
                        Console.WriteLine("-s:<sourceFileName> : required file parameter. <filemask>.mdb files will be processed.");
                        Console.WriteLine("-noprompt : disables the 'Continue' prompt. Program will run without the need for user interaction.");
                        Console.WriteLine("-o:\"<full path>\" : user specified output path. Default = current folder.");

                        Console.WriteLine("-nolog : disables the runtime log. !!! Set as very first parameter to fully disable log.");
                        Console.WriteLine("-lower : force all filenames to be lowercase.");

                        Console.WriteLine("-p : use PIPE '|' Delimiter in output file.");
                        Console.WriteLine("-t : use TAB Delimiter in output file.");
                        Console.WriteLine("-c : Replaces all symbol chars with '_' and converts FieldName to UPPER case");
                        Console.WriteLine("-adddate : Appends source file create date to output file(s).");
                        Console.WriteLine("-debug : enables debug 'verbose' mode.");
                        Console.WriteLine();
                        Console.WriteLine("Press Any Key To Continue.");
                        Console.ReadKey();
                        return;
                    }
                    else
                    {
                        Log.WriteToLogFile($"* Unknown Argument '{args[i]}'", true);
                    }
                }
            }

            Console.WriteLine();

            if (File.Exists(FileToProcess))
            {
                ExportDataFromAccessFile(FileToProcess, new List<string>());
            }
            else
            {
                Console.WriteLine("ERROR: File To Process was not found. Please check filename.");
            }

#if DEBUG
            Console.WriteLine($"{Environment.NewLine}Press Any Key To Continue.");
            Console.ReadKey();
#endif

            Environment.Exit(ExitCodeStatus);
            
        }
        /// <summary>
        /// Initialize Global Variables
        /// </summary>
        private static void InitApplicationVariables()
        {
            ExitCodeStatus = 0;
            FileToProcess = string.Empty;
            GenerateLogFile = true;
            DEBUGMODE = false;
            AllowOverWrite = true;
            OutputDirectory = string.Empty;
            DelimiterToUse = CSVDelimiter.comma;
            CleanFieldNames = false;
            FileNameCaseToUse = FileNameCase.none;
            AppendCreateDateToOutputFiles = false;
        }

        /// <summary>
        /// Process single Access file. Exporting specific tables or all tables (default)
        /// </summary>
        /// <param name="sourceFileName">The filename of mdb file to process</param>
        private static void ExportDataFromAccessFile(string sourceFileName, List<string> tableNames)
        {
            Log.WriteToLogFile($"# Processing mdb file: {Path.GetFileName(sourceFileName).ToLower()}");

            FileInfo sourceFileInfo = new FileInfo(sourceFileName);
            

            List<string> mdbUserTableNames = new List<string>();

            DataTable userTables = null;

            //var accODBCCon = new System.Data.Odbc.OdbcConnection();
            string accODBCConnectStr, activeODBCDriverName;

            activeODBCDriverName = Util.GetOdbcAccessDriverName();

            if (activeODBCDriverName == null)
            {
                Log.WriteToLogFile($"INFO: Access ODBC Driver not found on system or error reading registry. Will use default value.");
                activeODBCDriverName = "Microsoft Access Driver (*.mdb, *.accdb)";
            }

            Log.WriteToLogFile($"INFO: Access ODBC Driver Name = '{activeODBCDriverName}'");

            accODBCConnectStr = $"Driver={{{activeODBCDriverName}}};DBQ=" + sourceFileName + ";";

            Log.WriteToLogFile($"INFO: Access ODBC Connect String = '{accODBCConnectStr}'");
            try
            {
                using (OdbcConnection accODBCCon = new OdbcConnection(accODBCConnectStr))
                {

                    // open access connection
                    accODBCCon.Open();
                    userTables = accODBCCon.GetSchema("Tables");

                    if(DEBUGMODE) { 
                        Log.WriteToLogFile("#### MDB Table List ####");
                        Util.DisplayDataTable(userTables, true);
                        Log.WriteToLogFile("#### MDB Table List ####");
                    }

                    Log.WriteToLogFile($"Detecting tables in source file...", true);

                    foreach (DataRow row in userTables.Rows)
                    {
                        if (row["TABLE_TYPE"].ToString() == "TABLE")
                        {
                            mdbUserTableNames.Add(row["TABLE_NAME"].ToString());
                        }
                    }

                    Log.WriteToLogFile($"Found {mdbUserTableNames.Count} tables to process...", true);

                    foreach (string tableName in mdbUserTableNames)
                    {
                        Console.WriteLine("");
                        Log.WriteToLogFile($"### Processing table [{tableName}] ###", true);

                        string outputFilename = $"\\{tableName}";

                        switch (FileNameCaseToUse)
                        {
                            case FileNameCase.none:
                                break;
                            case FileNameCase.lower:
                                outputFilename = outputFilename.ToLower();
                                break;
                            case FileNameCase.upper:
                                outputFilename = outputFilename.ToUpper();
                                break;
                            default:
                                break;
                        }

                        if (AppendCreateDateToOutputFiles)
                        {
                            outputFilename = $"{OutputDirectory}{outputFilename}_{sourceFileInfo.CreationTime.ToString("yyyy-MM-dd")}.txt";
                        }
                        else
                        {
                            outputFilename = $"{OutputDirectory}{outputFilename}.txt";
                        }
                        

                        Console.WriteLine($"OUTPUT_FILENAME={outputFilename}");

                        if (File.Exists(outputFilename)) 
                        {
                            if(AllowOverWrite) 
                            {
                                Log.WriteToLogFile($"INFO: Output file exists and will be over-written.", true);
                            }
                            else 
                            {
                                Log.WriteToLogFile($"INFO: Output file exists. Skipping output for table {tableName}", true);
                                Console.WriteLine($"  The -nooverwrite option prevented write to file.");
                                continue; // skip table output
                            }
                            
                        }

                        Console.WriteLine();

                        OdbcCommand command = new OdbcCommand($"select [{tableName}].* from [{tableName}]", accODBCCon);

                        Log.WriteToLogFile($"INFO: SQL command={command.CommandText}");

                        var csv_config = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
                        {
                            TrimOptions = CsvHelper.Configuration.TrimOptions.Trim,
                        };

                        switch (DelimiterToUse)
                        {
                            case CSVDelimiter.comma:
                                csv_config.Delimiter = ",";
                                break;
                            case CSVDelimiter.tab:
                                csv_config.Delimiter = "\t";
                                break;
                            case CSVDelimiter.pipe:
                                csv_config.Delimiter = "|";
                                break;
                            default:
                                break;
                        }

                        using (var reader = command.ExecuteReader())
                        using (var writer = new StreamWriter(outputFilename))
                        using (var csv = new CsvWriter(writer, csv_config))
                        {
                            Console.WriteLine($"Starting output...");
                            Console.WriteLine($"Found {reader.FieldCount} Fields in table.");

                            // Write column headers
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (CleanFieldNames)
                                    csv.WriteField(Regex.Replace(reader.GetName(i).Trim(), @"[^0-9a-zA-Z]+", "_").ToUpper().TrimEnd('_'));
                                else
                                    csv.WriteField(reader.GetName(i));
                            }
                            csv.NextRecord();

                            var rowsWritten = 0;
                            // Write data rows
                            while (reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    csv.WriteField(reader[i]);
                                }
                                
                                rowsWritten++;
                                if(rowsWritten % 10000 == 0)
                                    Console.WriteLine($"{rowsWritten.ToString("#,#")} rows written.");

                                csv.NextRecord();
                            }

                            Console.WriteLine($"{rowsWritten.ToString("#,#")} total rows written to output file.");
                            Console.WriteLine("");
                        }

                    } // foreach table found
                    

                }

            }
            catch (Exception ex)
            {
                Log.WriteToLogFile($"ERROR caught while processing source MDB file: {sourceFileName}", true);
                Log.WriteToLogFile(ex.Message);
                Console.WriteLine(ex.Message);
                ExitCodeStatus = 99; // critical error
                return;
            }

        }


    }
}

