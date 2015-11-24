﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using org.in2bits.MyXls;
using System.Data;

namespace Infrastructure
{
    public class MyxlsExcelHelper
    {
        XlsDocument xls;
        public MyxlsExcelHelper()
        {
            xls = new XlsDocument();
        }

        private Worksheet CreateSheet(string sheetName)
        {
            return xls.Workbook.Worksheets.Add(sheetName);
        }

        public void InserCell(string sheetName, int row, int column, string value)
        {
            CreateSheet(sheetName).Cells.Add(row, column, value);
        }
        public void ImportDataTable(string sheetName, int startRow, DataTable dt)
        {
            Worksheet ws = CreateSheet(sheetName);
            XF xf = CellStyle();
            int index = 0;
            for (int c = 1; c <= dt.Columns.Count; c++)
            {
                ws.Cells.Add(startRow - 1, 1, "No.", xf);
                ws.Cells.Add(startRow - 1, c + 1, dt.Columns[index++].ColumnName, xf);
                //ws.Cells.Add(startRow - 1, dt.Columns.Count + 2, "Duration", xf);
            }
            int num = 1;
            foreach (DataRow rows in dt.Rows)
            {
                int column = 2;
                int start = 1;
                foreach (var item in rows.ItemArray)
                {
                    if (start == 1)
                    {
                        ws.Cells.Add(startRow, 1, num, xf);
                    }
                    switch (item.GetType().ToString())
                    {
                        case "System.String":
                            ws.Cells.Add(startRow, column++, item.ToString(), xf); break;
                        case "System.DateTime":
                            ws.Cells.Add(startRow, column++, ((DateTime)item).ToString("yyyy-MM-dd"), xf); break;
                        case "System.Decimal":
                            ws.Cells.Add(startRow, column++, Convert.ToDecimal(item).ToString("N2"), xf); break;
                    }
                    //ws.Cells.Add(startRow, column++, item, xf);
                    start++;
                    //if (start == rows.ItemArray.Length + 1)
                    //{
                    //    ws.Cells.Add(startRow, rows.ItemArray.Length + 2, "=TODAY()-J" + startRow, xf);
                    //}
                }
                startRow++;
                num++;
            }
        }
        public XF CellStyle()
        {
            XF xf = xls.NewXF();
            xf.HorizontalAlignment = HorizontalAlignments.Centered;
            xf.VerticalAlignment = VerticalAlignments.Centered;
            xf.UseBorder = true;
            xf.TopLineStyle = 1;
            xf.TopLineColor = Colors.Black;
            xf.BottomLineStyle = 1;
            xf.BottomLineColor = Colors.Black;
            xf.LeftLineStyle = 1;
            xf.LeftLineColor = Colors.Black;
            xf.RightLineStyle = 1;
            xf.RightLineColor = Colors.Black;
            xf.Font.Bold = false;
            xf.Font.Height = 9 * 20;
            return xf;
        }
        public void SaveExcel(string path)
        {
            Stream st = new System.IO.FileStream(path, FileMode.Create);
            xls.Save(st);
            st.Close();
        }
    }
}