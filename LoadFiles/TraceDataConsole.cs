using System;
using System.Diagnostics;
using System.IO;
using ExcelDataReader;

namespace LoadFiles
{
    public class TraceDataConsole : IFileLoader
    {
        private Action<string> _logger;
        private Stopwatch stopwatch = new Stopwatch();

        public TraceDataConsole(Action<string> logger=null)
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
}