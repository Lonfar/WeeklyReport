using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Infrastructure
{
    public interface IExcelHelper
    {
        void CreatSheet(string sheetName);
        void InsertCell(string sheetName, int row, int column, string value);
        void InsertCell(string sheetName, int row, int column, string value, Style style);
        void InsertDataTable(string sheetName, int startRow, DataTable dt,bool isDuration);
        void SaveExcel(string path);
    }
}
