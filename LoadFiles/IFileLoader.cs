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
}