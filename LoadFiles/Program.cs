using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using ExcelDataReader;
using System.Security.Cryptography;

namespace LoadFiles
{
    #region FileLoaderClass

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
        private FileInfo currentFileInfo;
        private string fileHash;
        private string currentSheet;
        private int currentSheetId;
        private DataTable fileSheetData;

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
            if(null != fileSheetData && null != connection)
                BulkCopyToFileSheetData(currentFile, currentFileInfo);

            connection?.Close();

            stopwatch.Stop();
            TimeSpan t = stopwatch.Elapsed;
            _logger.Invoke($"Finished {DateTime.Now.ToString()} {t:hh\\:mm\\:ss}");
        }

        public int OnCol(int col, int row, IExcelDataReader reader)
        {
            string value = reader.GetValue(col) == null ? null : reader.GetValue(col).ToString();
            string type = "null";
            if (reader.GetValue(col) != null)
                type = reader.GetFieldType(col).FullName;
            if (null != fileSheetData)
            {
                DataRow dr = fileSheetData.NewRow();
                dr["FSD_FLE_Id"] = currentSheetId;
                dr["FSD_Row"] = row;
                dr["FSD_Col"] = col;
                dr["FSD_Data"] = value;
                fileSheetData.Rows.Add(dr);
            }
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
            if (null != fileSheetData)
            {
                BulkCopyToFileSheetData(filePath, info);
            }
            else
            {
                fileSheetData = new DataTable("NewColumnData");
                fileSheetData.Columns.Add("FSD_Id", typeof(int));
                fileSheetData.Columns.Add("FSD_FLE_Id", typeof(int));
                fileSheetData.Columns.Add("FSD_Row", typeof(int));
                fileSheetData.Columns.Add("FSD_Col", typeof(int));
                fileSheetData.Columns.Add("FSD_Data", typeof(string));

                fileSheetData.Columns["FSD_Id"].AutoIncrement = true;

                // don't allow null for these columns
                fileSheetData.Columns["FSD_FLE_Id"].AllowDBNull = false;
                fileSheetData.Columns["FSD_Row"].AllowDBNull = false;
                fileSheetData.Columns["FSD_Col"].AllowDBNull = false;
            }
            currentFile = filePath;
            currentFileInfo = info;

            // create a file hash 
            if (!string.IsNullOrEmpty(currentFile) && File.Exists(currentFile))
            {
                SHA256 Sha256 = SHA256.Create();
                using (FileStream stream = File.OpenRead(currentFile))
                {
                    byte[] hashBytes = Sha256.ComputeHash(stream);
                    fileHash = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
                }
            }

            _logger.Invoke($"File {info.Name}");
            return 0;
        }

        private void BulkCopyToFileSheetData(string filePath, FileInfo info)
        {
            using (SqlBulkCopy bulk = new SqlBulkCopy(connection))
            {
                bulk.DestinationTableName = "dbo.FileSheetData";
                try
                {
                    fileSheetData.AcceptChanges();
                    bulk.WriteToServer(fileSheetData);
                }
                catch (Exception ex)
                {
                    _logger.Invoke(
                        $"Exception calling SqlBulkCopy.WriteToServer OnFile('{filePath}', '{info.ToString()}');\n{ex.ToString()}");
                }
            }

            fileSheetData.Rows.Clear();
        }

        public int OnSheet(string sheetName)
        {
            currentSheet = sheetName;
            // check db and see if we are reloading or loading
            // using hash and check if it changed
            int fileSheetId = GetSheetId(currentFile, currentSheet, fileHash);

            if (fileSheetId < 1)  // skip the file, either it was already loaded or there is a problem
            {
                string fileName = Path.GetFileName(currentFile);
                _logger.Invoke($"Skipping sheet GetSheetId('{fileName}', '{currentSheet}', '{fileHash}') returned {fileSheetId.ToString()}");
                return 1;
            }

            currentSheetId = fileSheetId;

            _logger.Invoke($"Sheet '{sheetName}'");
            return 0;
        }

        private int GetSheetId(string f, string s, string h)
        {
            try
            {
                if (connection.State == ConnectionState.Open)
                {
                    string fileName = Path.GetFileName(f);
                    string sql = "dbo.GetSheetId";
                    SqlCommand cmd = new SqlCommand(sql, connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FileName", fileName);
                    cmd.Parameters.AddWithValue("@Sheet", s);
                    cmd.Parameters.AddWithValue("@Hash", h);
                    Int32 id = Convert.ToInt32(cmd.ExecuteScalar());
                    cmd.Dispose();
                    return id;
                }
                else return -1;
            }
            catch (Exception ex)
            {
                _logger.Invoke($"Exception calling GetSheetId('{f}', '{s}', '{h}');\n{ex.ToString()}");
                return -2;
            }
        }
    }

}
