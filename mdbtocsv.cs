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


        private static string FileToProcess { get; set; }
        private static bool ForceLowercaseFilename { get; set; }
        private static bool GenerateLogFile { get; set; }
        private static bool DEBUGMODE { get; set; }
        private static bool EnableOverWriteWarning { get; set; }
        private static string OutputDirectory { get; set; }
        private static CSVDelimiter DelimiterToUse { get; set; }
        private static bool CleanFieldNames { get; set; }

        static void Main(string[] args)
        {
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
                    else if (args[i].ToLower() == "-noprompt")
                    {
                        Log.WriteToLogFile("* startup param: disabling over-write warning.", true);
                        EnableOverWriteWarning = false;
                    }
                    else if (args[i].ToLower() == "-lower")
                    {
                        Log.WriteToLogFile("* startup param: output filenames will be lowercase.", true);
                        ForceLowercaseFilename = true;
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
                        Log.WriteToLogFile("* startup param: Field Name Cleanup Option ENABLED.", true);
                        CleanFieldNames = true;
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

            // TODO: disable prompt if -noprompt options is enabled.
            Console.WriteLine($"{Environment.NewLine}Press Any Key To Continue.");
            Console.ReadKey();

        }
        /// <summary>
        /// Initialize Global Variables
        /// </summary>
        private static void InitApplicationVariables()
        {
            FileToProcess = string.Empty;
            GenerateLogFile = true;
            DEBUGMODE = false;
            EnableOverWriteWarning = true;
            ForceLowercaseFilename = false;
            OutputDirectory = string.Empty;
            DelimiterToUse = CSVDelimiter.comma;
            CleanFieldNames = false;
        }

        /// <summary>
        /// Process single Access file. Exporting specific tables or all tables (default)
        /// </summary>
        /// <param name="sourceFileName">The filename of mdb file to process</param>
        private static void ExportDataFromAccessFile(string sourceFileName, List<string> tableNames)
        {
            Log.WriteToLogFile($"# Processing mdb file: {Path.GetFileName(sourceFileName).ToLower()}");

            List<string> mdbUserTableNames = new List<string>();

            DataTable userTables = null;

            //var accODBCCon = new System.Data.Odbc.OdbcConnection();
            string accODBCConnectStr, installedDriver;

            installedDriver = Util.GetOdbcAccessDriverName();

            if (installedDriver == null)
            {
                Log.WriteToLogFile($"INFO: Access ODBC Driver not found on system or error reading registry. Will use default value.");
                installedDriver = "Microsoft Access Driver (*.mdb, *.accdb)";
            }

            Log.WriteToLogFile($"INFO: Access ODBC Driver Name = '{installedDriver}'");

            accODBCConnectStr = $"Driver={{{installedDriver}}};DBQ=" + sourceFileName + ";";

            Log.WriteToLogFile($"INFO: Access ODBC Connect String = '{accODBCConnectStr}'");
            try
            {
                using (OdbcConnection accODBCCon = new OdbcConnection(accODBCConnectStr))
                {

                    // open access connection
                    accODBCCon.Open();

                    userTables = accODBCCon.GetSchema("Tables");

                    if (DEBUGMODE)
                    {
                        Log.WriteToLogFile("#### MDB Table List ####");
                        DisplayDataTable(userTables, true);
                        Log.WriteToLogFile("#### MDB Table List ####");
                    }

                    foreach (DataRow row in userTables.Rows)
                    {
                        if (row["TABLE_TYPE"].ToString() == "TABLE")
                        {
                            mdbUserTableNames.Add(row["TABLE_NAME"].ToString());
                        }
                    }

                    // TODO: Wrap csvwriter code in foreach loop for each table
                    // TODO: Filter tables which are selected for output

                    foreach (string t in mdbUserTableNames)
                        Log.WriteToLogFile($"Detected table [{t}]", true);

                    string outputFilename = OutputDirectory + $"\\{mdbUserTableNames[0]}.txt";

                    

                    if(ForceLowercaseFilename)
                        outputFilename = outputFilename.ToLower();

                    Console.WriteLine($"{Environment.NewLine}OUTPUT_FILENAME={outputFilename}{Environment.NewLine}");

                    OdbcCommand command = new OdbcCommand($"select [{mdbUserTableNames[0]}].* from [{mdbUserTableNames[0]}]", accODBCCon);
                    
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

                    // TODO: check if output file exists and if -noprompt option is enabled

                    using (var reader = command.ExecuteReader())
                    using (var writer = new StreamWriter(outputFilename))
                    using (var csv = new CsvWriter(writer, csv_config))
                    {
                        Console.WriteLine($"Found {reader.FieldCount} Fields in source table.");

                        // Write column headers
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if(CleanFieldNames)
                                csv.WriteField(Regex.Replace(reader.GetName(i).Trim().ToUpper(), @"[^0-9a-zA-Z]+", "_").TrimEnd('_'));
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
                            csv.NextRecord();
                        }

                        Console.WriteLine($"Wrote {rowsWritten.ToString("#,#")} rows to output file.");
                    }

                }

            }
            catch (Exception ex)
            {
                Log.WriteToLogFile($"ERROR caught while reading source MDB file: {sourceFileName}", true);
                Log.WriteToLogFile(ex.Message);
                Console.WriteLine(ex.Message);
                return;
            }

        }


        /// <summary>
        /// Displays data from a data table to console. Used for Debuging.
        /// </summary>
        /// <param name="table"></param>
        /// <remarks></remarks>
        private static void DisplayDataTable(DataTable table, bool writeToLogFileOnly = false)
        {
            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn col in table.Columns)
                {
                    if (writeToLogFileOnly)
                    { Log.WriteToLogFile($"{col.ColumnName} = {row[col]}"); }
                    else
                    { Console.WriteLine("{0} = {1}", col.ColumnName, row[col]); }
                }
                Log.WriteToLogFile("============================");
            }
        }
    }
}

