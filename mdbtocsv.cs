using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace mdbtocsv
{
    internal class mdbtocsv
    {
        public enum B2QuoteIdentifier
        {
            None, DoubleQuote
        }

        private static ArrayList FileMask { get; set; }
        private static bool ProcessExcelFiles { get; set; }
        private static bool ProcessAccessFiles { get; set; }
        private static bool ForceLowercaseFilename { get; set; }
        private static bool GenerateLogFile { get; set; }
        private static bool DEBUGMODE { get; set; }
        private static bool EnableOverWriteWarning { get; set; }
        private static string FolderNameToProcess { get; set; }
        public static string LogFileName { get; set; }

        private static readonly int MAXLOGSIZEKB = 128;

        static void Main(string[] args)
        {
            Assembly thisAssem = typeof(mdbtocsv).Assembly;
            AssemblyName thisAssemName = thisAssem.GetName();
            Version ver = thisAssemName.Version;

            System.IO.FileInfo fi = new System.IO.FileInfo(thisAssem.Location);
            FolderNameToProcess = fi.DirectoryName;

            LogFileName = fi.DirectoryName + Path.DirectorySeparatorChar + "rename activity log.txt";
            Debug.Print($"LogFilename: {LogFileName}");
            //TrimLogFile();



        }
    }
}
