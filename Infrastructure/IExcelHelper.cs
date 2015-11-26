using Aspose.Cells;
using System.Data;

namespace Infrastructure
{
    public interface IExcelHelper
    {
        void CreatSheet(string sheetName);

        void InsertCell(string sheetName, int row, int column, object value);

        void InsertCell(string sheetName, int row, int column, object value, Style style);
        void InsertDataTable(string sheetName, int startRow, DataTable dt, bool isDuration);

        void SaveExcel(string path);
        void InsertFormula(string sheetName, int row, int column, object value);
    }
}