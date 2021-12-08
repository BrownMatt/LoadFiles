using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using ExcelDataReader;

namespace LoadFiles
{
    public interface IFileLoader
    {
        // Called once before any processing
        void OnInit();
        // Called once when all processing is complete
        void OnFinish();
        int OnCol(int col, int row, IExcelDataReader reader);
        int OnRow(int rowNumber, int readerFieldCount, IExcelDataReader reader);
        int OnFile(string filePath, FileInfo info);
        int OnSheet(string sheetName);
    }

    public class TraceData : IFileLoader
    {
        private Action<string> _logger;
        private Stopwatch stopwatch = new Stopwatch();

        public TraceData(Action<string> logger=null)
        {
            _logger = logger ?? Console.WriteLine;
        }

        public void OnInit()
        {
            stopwatch.Start();
        }

        public void OnFinish()
        {
            stopwatch.Stop();
            TimeSpan t = stopwatch.Elapsed;
            _logger.Invoke($"Finished {DateTime.Now.ToString()} {t:hh\\:mm\\:ss}");
        }

        public int OnCol(int col, int row, IExcelDataReader reader)
        {
            string value = reader.GetValue(col) == null ? "null" : reader.GetValue(col).ToString();
            string type = "null";
            if (reader.GetValue(col) != null)
                type = reader.GetFieldType(col).FullName;

            _logger.Invoke($"Column {col.ToString()} Type {type} Value {value}");
            return 0;
        }

        public int OnRow(int rowNumber, int readerFieldCount, IExcelDataReader reader)
        {
            if (rowNumber == 0)
            {
                if (null != reader.HeaderFooter)
                {
                    HeaderFooter h = reader.HeaderFooter;
                    _logger.Invoke($"Contains header!");
                }
            }
            _logger.Invoke($"Row {rowNumber.ToString()} FieldCount {readerFieldCount}");
            return 0;
        }

        public int OnFile(string filePath, FileInfo info)
        {
            _logger.Invoke($"File {info.Name}");
            return 0;
        }

        public int OnSheet(string sheetName)
        {
            _logger.Invoke($"Sheet '{sheetName}'");
            return 0;
        }
    }

    

    #region FileLoaderClass

    public class FileLoader
    {
        private string _logFolder;
        private string _loadFolder;
        private IFileLoader _loader;
        private Action<string> _logger;

        public FileLoader(string logFolder = "", string loadFolder = "", IFileLoader loader = null, Action<string> logger = null)
        {
            _logFolder = (!string.IsNullOrEmpty(logFolder)) ? logFolder : ConfigurationManager.AppSettings["LogPath"];
            _loadFolder = (!string.IsNullOrEmpty(loadFolder)) ? loadFolder : ConfigurationManager.AppSettings["LoadDirectory"];
            if (null == _loadFolder)
            {
                _loadFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                _loadFolder = Path.Combine(_loadFolder, "incoming");
            }

            _logger = logger ?? Console.WriteLine;
            _loader = loader ?? new TraceData(_logger);
        }


        public int LoadFiles()
        {
            _loader.OnInit();

            int count = 0;

            string searchPattern = ConfigurationManager.AppSettings["SearchPattern"];
            if (string.IsNullOrEmpty(searchPattern))
                searchPattern = "*.*";

            List<string> matchingFiles = GetMatchingFiles(_loadFolder, true, searchPattern);

            // Now process files in load folder
            foreach (var file in matchingFiles)
            {
                _logger.Invoke($"Processing file '{file}'");
                if (0 == LoadFile(file)) count++;
            }

            _loader.OnFinish();

            return count;

        }

        private List<string> GetMatchingFiles(string rootFolder, bool searchSubfolders, string searchPattern)
        {
            Queue<string> folders = new Queue<string>();
            List<string> files = new List<string>();
            folders.Enqueue(rootFolder);
            while (folders.Count != 0)
            {
                string currentFolder = folders.Dequeue();
                try
                {
                    string[] filesInCurrent = Directory.GetFiles(currentFolder, searchPattern, System.IO.SearchOption.TopDirectoryOnly);
                    if (filesInCurrent.Length > 0)
                        files.AddRange(filesInCurrent);
                }
                catch (Exception ex)
                {
                    _logger.Invoke($"Directory.GetFiles Exception {ex.ToString()}");
                }
                try
                {
                    if (searchSubfolders)
                    {
                        string[] foldersInCurrent = Directory.GetDirectories(currentFolder, "*.*", System.IO.SearchOption.TopDirectoryOnly);
                        foreach (string _current in foldersInCurrent)
                        {
                            folders.Enqueue(_current);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Invoke($"Directory.GetDirectories Exception {ex.ToString()}");
                }
            }
            return files;
        }

        // Returns zero if file successfully loaded
        private int LoadFile(string filePath)
        {
            int retVal = 0;


            FileInfo info = new FileInfo(filePath);
            _logger.Invoke($"Extension '{info.Extension}' Size={info.Length.ToString()}");
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
                    _logger.Invoke($"{info.Name} contains {reader.ResultsCount.ToString()} sheets");
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
                    _logger.Invoke($"{info.Name} contains {reader.ResultsCount.ToString()} sheets");
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




        public int OnCol(int col, int row, IExcelDataReader reader)
        {
            return _loader.OnCol(col, row, reader);
        }

        public int OnRow(int rowNumber, int readerFieldCount, IExcelDataReader reader)
        {
            return _loader.OnRow(rowNumber, readerFieldCount, reader);
        }

        // return non zero to skip the file
        // or zero to process file
        public int OnFile(string filePath, FileInfo info)
        {
            return _loader.OnFile(filePath, info);
        }

        // return non zero to skip the sheet
        // or zero to process sheet
        public int OnSheet(string sheetName)
        {
            return _loader.OnSheet(sheetName);
        }
    }

    #endregion

    class Program
    {
        static void Main(string[] args)
        {
            IFileLoader l = null;
            switch (Environment.OSVersion.VersionString.ToLower())
            {
                case string m when m.Contains("windows"):
                    l = new LoadDataToSqlServer(ConfigurationManager.AppSettings["SqlSvrCon"]);
                    break;
            }

            FileLoader fl = new FileLoader(loader: l);
            fl.LoadFiles();
            
            if(System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Press a key to exit");
                Console.ReadKey();
            }
        }
    }


    public class LoadDataToSqlServer : IFileLoader
    {
        private Action<string> _logger;
        private Stopwatch stopwatch = new Stopwatch();
        private SqlConnection connection;
        private string currentFile;
        private string currentSheet;

        public LoadDataToSqlServer(string connectionString, Action<string> logger = null)
        {
            _logger = logger ?? Console.WriteLine;
            connection = new SqlConnection(connectionString);
            connection.Open();
        }

        public void OnInit()
        {
            stopwatch.Start();
        }

        public void OnFinish()
        {
            connection?.Close();

            stopwatch.Stop();
            TimeSpan t = stopwatch.Elapsed;
            _logger.Invoke($"Finished {DateTime.Now.ToString()} {t:hh\\:mm\\:ss}");
        }

        public int OnCol(int col, int row, IExcelDataReader reader)
        {
            string value = reader.GetValue(col) == null ? "null" : reader.GetValue(col).ToString();
            string type = "null";
            if (reader.GetValue(col) != null)
                type = reader.GetFieldType(col).FullName;

            _logger.Invoke($"Column {col.ToString()} Type {type} Value {value}");
            return 0;
        }

        public int OnRow(int rowNumber, int readerFieldCount, IExcelDataReader reader)
        {
            if (rowNumber == 0)
            {
                if (null != reader.HeaderFooter)
                {
                    HeaderFooter h = reader.HeaderFooter;
                    _logger.Invoke($"Contains header!");
                }
            }
            _logger.Invoke($"Row {rowNumber.ToString()} FieldCount {readerFieldCount}");
            return 0;
        }

        public int OnFile(string filePath, FileInfo info)
        {
            currentFile = filePath;
            _logger.Invoke($"File {info.Name}");
            return 0;
        }

        public int OnSheet(string sheetName)
        {
            currentSheet = sheetName;
            // check if we are reloading or loading
            // Do a file hash and check if it changed


            _logger.Invoke($"Sheet '{sheetName}'");
            return 0;
        }
    }

}
