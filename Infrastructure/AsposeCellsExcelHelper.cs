using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Data;

namespace Infrastructure
{
    public class AsposeCellsExcelHelper : IExcelHelper
    {
        Workbook xls;
        public AsposeCellsExcelHelper()
        {
            xls = new Workbook();
        }
        public AsposeCellsExcelHelper(string path)
        {
            xls = new Workbook(path);
        }
        public void CreatSheet(string sheetName)
        {
            xls.Worksheets.Add(sheetName);
        }
        public void ChangeSheetName(string formName, string toName)
        {
            xls.Worksheets[formName].Name = toName;
        }
        public void InsertCell(string sheetName, int row, int column, string value, Style style)
        {
            xls.Worksheets[sheetName].Cells[row, column].Value = value;
            xls.Worksheets[sheetName].Cells[row, column].SetStyle(style);
        }
        public void InsertDataTable(string sheetName, int startRow, DataTable dt, bool isDuration)
        {
            if (xls.Worksheets[0].Name != sheetName && xls.Worksheets[0].Name != sheetName)
                CreatSheet(sheetName);
            int index = 0;
            for (int c = 1; c <= dt.Columns.Count; c++)
            {
                if (c == 1)
                {
                    InsertCell(sheetName, startRow - 2, 0, "No.", AddStyle());
                }
                InsertCell(sheetName, startRow - 2, c + 0, dt.Columns[index++].ColumnName, AddStyle());
                if (c == dt.Columns.Count && isDuration)
                {
                    InsertCell(sheetName, startRow - 2, dt.Columns.Count + 1, "Duration", AddStyle());
                }
            }
            int num = 1;
            foreach (DataRow rows in dt.Rows)
            {
                int column = 1;
                int start = 1;
                foreach (var item in rows.ItemArray)
                {
                    if (start == 1)
                    {
                        InsertCell(sheetName, startRow - 1, 0, num.ToString(), AddStyle());
                    }
                    switch (item.GetType().ToString())
                    {
                        case "System.String":
                            InsertCell(sheetName, startRow - 1, column++, item.ToString(), AddStyle());
                            break;
                        case "System.DateTime":
                            InsertCell(sheetName, startRow - 1, column++, ((DateTime)item).ToString("yyyy-MM-dd"), AddStyle());
                            break;
                        case "System.Decimal":
                            InsertCell(sheetName, startRow - 1, column++, Convert.ToDecimal(item).ToString("N2"), AddStyle());
                            break;
                        default:
                            InsertCell(sheetName, startRow - 1, column++, item.ToString(), AddStyle());
                            break;
                    }
                    //ws.Cells.Add(startRow, column++, item, xf);
                    start++;
                    if (start == rows.ItemArray.Length + 1 && isDuration)
                    {
                        //InsertCell(sheetName, startRow - 1, rows.ItemArray.Length + 1, "=TODAY()-J" + (startRow), AddStyle());
                        xls.Worksheets[sheetName].Cells[startRow - 1, rows.ItemArray.Length + 1].Formula = "=TODAY()-J" + (startRow);
                        xls.Worksheets[sheetName].Cells[startRow - 1, rows.ItemArray.Length + 1].SetStyle(AddStyle());
                    }
                }
                startRow++;
                num++;
            }
        }
        public void SaveExcel(string path)
        {
            xls.Save(path);
            xls.Dispose();

        }

        private Style AddStyle()
        {
            Style style = new Style();
            style.Borders.SetColor(System.Drawing.Color.Black);
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Font.Size = 9;
            return style;
        }

        public void InsertCell(string sheetName, int row, int column, string value)
        {
            throw new NotImplementedException();
        }
    }
}
