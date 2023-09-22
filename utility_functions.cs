
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using logging;
using Microsoft.Win32;

namespace mdbtocsv_util
{

    public static class Util
    {
        /// <summary>
        /// Determine if given filename is locked by other processes
        /// </summary>
        /// <param name="filePath">file to evaluate</param>
        /// <returns>TRUE if file is locked</returns>
        public static bool IsFileLocked(string filePath)
        {
            try
            {
                using (File.Open(filePath, FileMode.Open)) { }
                Debug.Print("[IsFileLocked] File is NOT Locked.");
            }
            catch (IOException e)
            {
                var errorCode = Marshal.GetHRForException(e) & ((1 << 16) - 1);

                Debug.Print($"[IsFileLocked] File Is Locked!");
                System.Threading.Thread.Sleep(100);
                return errorCode == 32 || errorCode == 33;
            }

            return false;
        }

        /// <summary>
        /// Gets the name of an ODBC driver for Microsoft Access giving preference to the most recent one.
        /// </summary>
        /// <returns>the name of an ODBC driver for Microsoft Access, if one is present; null, otherwise.</returns>
        public static string GetOdbcAccessDriverName()
        {
            string driverName = null;

            try
            {

                List<string> driverPrecedence = new List<string>() { "Microsoft Access Driver (*.mdb, *.accdb)", "Microsoft Access Driver (*.mdb)" };
                string[] availableOdbcDrivers = GetOdbcDriverNames();

                if (availableOdbcDrivers != null)
                {
                    driverName = driverPrecedence.Intersect(availableOdbcDrivers).FirstOrDefault();
                }

            }
            catch (Exception ex)
            {
                Log.WriteToLogFile($"[GetOdbcAccessDriverName] CAUGHT ERROR : {ex.Message}");

            }

            return driverName;
        }


        /// <summary>
        /// Gets the ODBC driver names from the registry.
        /// </summary>
        /// <returns>a string array containing the ODBC driver names, if the registry key is present; null, otherwise.</returns>
        public static string[] GetOdbcDriverNames(bool debugmode = false)
        {

            string[] odbcDriverNames = null;

            try
            {
                using (RegistryKey localMachineHive = Registry.LocalMachine)
                using (RegistryKey odbcDriversKey = localMachineHive.OpenSubKey(@"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"))
                {
                    if (odbcDriversKey != null)
                    {
                        odbcDriverNames = odbcDriversKey.GetValueNames();
                    }
                }

                if (debugmode)
                {
                    foreach (var d in odbcDriverNames)
                    {
                        Log.WriteToLogFile($"INFO:ODBC_DRIVERS:{d}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.WriteToLogFile($"[GetOdbcDriverNames] CAUGHT ERROR : {ex.Message}");
            }

            return odbcDriverNames;
        }


        /// <summary>
        /// Displays data from a data table to console. Used for Debuging.
        /// </summary>
        /// <param name="table"></param>
        /// <remarks></remarks>
        public static void DisplayDataTable(DataTable table, bool writeToLogFileOnly = false)
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