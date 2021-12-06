using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using ExcelDataReader;

namespace LoadFiles
{

    public class FileLoader
    {
        private string _logFolder;
        private string _loadFolder;

        public FileLoader(string logFolder="", string loadFolder="")
        {
            _logFolder = (!string.IsNullOrEmpty(logFolder)) ? logFolder : ConfigurationManager.AppSettings["LogPath"];
            _loadFolder = (!string.IsNullOrEmpty(loadFolder)) ? loadFolder : ConfigurationManager.AppSettings["LoadDirectory"];
        }

     
        public int LoadFiles()
        {
            int count = 0;

            // If processing sub folders
            string[] subFolders = Directory.GetDirectories(_loadFolder);

            // Now process files in load folder
            foreach (var file in Directory.EnumerateFiles(_loadFolder, "*.*"))
            {
                Console.WriteLine($"Processing file '{file}'");
                if (0 == LoadFile(file)) count++;
            }

            return count;
                
        }

        // Returns zero if file successfully loaded
        private int LoadFile(string filePath)
        {
            int retVal = 0;


            FileInfo info = new FileInfo(filePath);
            Console.WriteLine($"Extension '{info.Extension}' Size={info.Length.ToString()}");
            if (0 != OnFile(filePath, info)) return retVal;

            switch (info.Extension.ToLower())
            {
                case ".csv":
                    retVal = LoadCsvFile(filePath, info);
                    break;
                case ".json": break;
                case ".txt": break;
                case ".xls": break;
                case ".xlsx":
                    //retVal = LoadXlsxFile(filePath, info);
                    break;
                case ".zip": break;
            }


            return retVal;
        }


        private int LoadXlsxFile(string filePath, FileInfo info)
        {
            int retVal = 0;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                //var cfg = new ExcelReaderConfiguration(){}
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    Console.WriteLine($"{info.Name} contains {reader.ResultsCount.ToString()} sheets");
                    do
                    {
                        if(0 == OnSheet(reader.Name))
                        {
                            int row = 0;
                            while (reader.Read())
                            {
                                if (0 == OnRow(row++, reader.FieldCount, reader))
                                {
                                    for (int col = 0; col < reader.FieldCount; col++)
                                    {
                                        if (0 != OnCol(col, row, reader))
                                            return retVal;
                                    }
                                }
                                // reader.GetDouble(0);
                            }
                        }
                    } while (reader.NextResult());  // Next Sheet
                }
            }
            return retVal;
        }
        private int LoadCsvFile(string filePath, FileInfo info)
        {
            int retVal = 0;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var cfg = new ExcelReaderConfiguration()
                {
                    AnalyzeInitialCsvRows = 1
                };
                using (var reader = ExcelReaderFactory.CreateCsvReader(stream, cfg))
                {
                    Console.WriteLine($"{info.Name} contains {reader.ResultsCount.ToString()} sheets");
                    do
                    {
                        if (0 == OnSheet(reader.Name))
                        {
                            int row = 0;
                            while (reader.Read())
                            {
                                if (0 == OnRow(row++, reader.FieldCount, reader))
                                {
                                    for (int col = 0; col < reader.FieldCount; col++)
                                    {
                                        if (0 != OnCol(col, row, reader))
                                            return retVal;
                                    }
                                }
                                // reader.GetDouble(0);
                            }
                        }
                    } while (reader.NextResult());  // Next Sheet
                }
            }
            return retVal;
        }


        private int OnCol(int col, int row, IExcelDataReader reader)
        {
            string value = reader.GetValue(col) == null ? "null" : reader.GetValue(col).ToString();
            string type = "null";
            if (reader.GetValue(col) != null)
                type = reader.GetFieldType(col).FullName;

            Console.WriteLine($"Column {col.ToString()} Type {type} Value {value}");
            return 0;
        }

        private int OnRow(int rowNumber, int readerFieldCount, IExcelDataReader reader)
        {
            if (rowNumber == 0)
            {
                if (null != reader.HeaderFooter)
                {
                    HeaderFooter h = reader.HeaderFooter;
                    Console.WriteLine($"Contains header!");
                }
            }
            Console.WriteLine($"Row {rowNumber.ToString()} FieldCount {readerFieldCount}");
            return 0;
        }

        // return non zero to skip the file
        // or zero to process file
        private int OnFile(string filePath, FileInfo info)
        {
            Console.WriteLine($"File {info.Name}");
            return 0;
        }

        // return non zero to skip the sheet
        // or zero to process sheet
        private int OnSheet(string sheetName)
        {
            Console.WriteLine($"Sheet '{sheetName}'");
            return 0;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            FileLoader fl = new FileLoader();
            fl.LoadFiles();

            Console.WriteLine("Press a key to exit");
            Console.ReadKey();
        }
    }
}
