using Infrastructure;
using System;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;

namespace WindowsForms
{
    public partial class index : Form
    {
        private ODBC odbc = new ODBC();

        public index()
        {
            InitializeComponent();
        }

        private void index_Load(object sender, EventArgs e)
        {
            StartDateTime.Value = DateTime.Now.AddDays(-7);
        }

        private void materialInventryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportInventory(folderBrowserDialog1.SelectedPath + @"material inventry report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }

        private void materialIssueReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportIssue(folderBrowserDialog1.SelectedPath + @"material issue report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }

        private void materialReceiveReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportReceiving(folderBrowserDialog1.SelectedPath + @"material Receive report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }

        private void materialTransferBackReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportTransferBack(folderBrowserDialog1.SelectedPath + @"material TransferBack report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }

        private void materialWellToWellReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportTransferWellToWell(folderBrowserDialog1.SelectedPath + @"material WellToWell report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }

        private void sDTWHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                SDTWH();
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }
        private void 各部门库存情况ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportInventoryInfor(folderBrowserDialog1.SelectedPath + @"各部门库存情况 " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
            }
        }
        private void 全部导出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportInventory(folderBrowserDialog1.SelectedPath + @"material inventry report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                ExportReceiving(folderBrowserDialog1.SelectedPath + @"material Receive report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                ExportIssue(folderBrowserDialog1.SelectedPath + @"material issue report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                ExportTransferWellToWell(folderBrowserDialog1.SelectedPath + @"material WellToWell report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                ExportTransferBack(folderBrowserDialog1.SelectedPath + @"material TransferBack report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
                SDTWH();
                MessageBox.Show("已成功导出");
                Process.Start("explorer.exe", folderBrowserDialog1.SelectedPath);
            }
        }


        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        #region 私有方法

        private void ExportInventory(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [INVENTORY]";
            //初始化excel操作类，新建Excel
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\material inventry report.xlsx");
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            excel.InsertDataTable("ALL", 8, dt, true);
            excel.InsertFormula("ALL", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Production Dept.'");
            excel.InsertDataTable("Production Dept.", 8, dt, true);
            //excel.InsertFormula("Production Dept.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Field Operation Dept.' or [Depart Desc]='Base camp'");
            excel.InsertDataTable("Field Operation Dept.", 8, dt, true);
            //excel.InsertFormula("Field Operation Dept.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Pipeline'");
            excel.InsertDataTable("Pipeline", 8, dt, true);
            //excel.InsertFormula("Pipeline", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Development Dep.'");
            excel.InsertDataTable("Development Dep.", 8, dt, true);
            //excel.InsertFormula("Development Dep.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Operation Dept.'");
            excel.InsertDataTable("Operation Dept.", 8, dt, true);
            //excel.InsertFormula("Operation Dept.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Engineering & Construction'");
            excel.InsertDataTable("Engineering & Construction", 8, dt, true);
            //excel.InsertFormula("Engineering & Construction", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='HSE Dept.'");
            excel.InsertDataTable("HSE Dept.", 8, dt, true);
            //excel.InsertFormula("HSE Dept.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Administration Dept.'");
            excel.InsertDataTable("Administration Dept.", 8, dt, true);
            //excel.InsertFormula("Administration Dept.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='P&L Dept.'");
            excel.InsertDataTable("P&L Dept.", 8, dt, true);
            //excel.InsertFormula("P&L Dept.", dt.Rows.Count + 8, 6, "=SUM(G1:G" + (dt.Rows.Count + 7) + ")");

            excel.SaveExcel(SavePath);
        }

        private void ExportIssue(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_IssueReport] where [Issue Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Issue Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\material issue report.xlsx");
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                excel.InsertDataTable("Sheet1", 8, dt, false);
                excel.SaveExcel(SavePath);
            }
        }

        private void ExportReceiving(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_ReceiveReport] where [Recv Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Recv Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\material Receive report.xlsx");
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                excel.InsertDataTable("Sheet1", 8, dt, false);
                excel.SaveExcel(SavePath);
            }
        }

        private void ExportTransferBack(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_TransferBackReport] where [Transfer Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Transfer Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\material TransferBack report.xlsx");
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                excel.InsertDataTable("Sheet1", 8, dt, false);
                excel.SaveExcel(SavePath);
            }
        }

        private void ExportTransferWellToWell(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_WellToWellReport] where [Transfer Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Transfer Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\material WellToWell report.xlsx");
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                excel.InsertDataTable("Sheet1", 8, dt, false);
                excel.SaveExcel(SavePath);
            }
        }

        private void SDTWH()
        {
            string sql = "";
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\SDT-WH.xlsx");
            excel.InsertCell("Sheet1", 10, 0, StartDateTime.Value.ToString("yyyy-MM-dd"));
            excel.InsertCell("Sheet1", 10, 4, EndDateTime.Value.ToString("yyyy-MM-dd"));

            sql = @"SELECT dbo.WH_Receive.ReceiveNO AS Expr1 FROM
                        dbo.WH_Receive LEFT OUTER JOIN
                        dbo.WH_ReceiveMaterial ON dbo.WH_Receive.ReceiveID = dbo.WH_ReceiveMaterial.ReceiveID
                        WHERE(dbo.WH_ReceiveMaterial.MaterialName <> N'Diesel Fuel')
                        GROUP BY dbo.WH_Receive.ReceiveNO
                        HAVING (MAX(dbo.WH_Receive.ReceiveDate) >= '" + StartDateTime.Value.ToString("yyyy-MM-dd") + @"') AND
                        (MAX(dbo.WH_Receive.ReceiveDate) <= '" + EndDateTime.Value.ToString("yyyy-MM-dd") + @"')
                        ORDER BY MAX(dbo.WH_Receive.ReceiveDate), dbo.WH_Receive.ReceiveNO";
            NewMethod(excel, 14, 3, sql);

            sql = @"SELECT dbo.WH_Receive.ReceiveNO AS Expr1 FROM
                        dbo.WH_Receive LEFT OUTER JOIN
                        dbo.WH_ReceiveMaterial ON dbo.WH_Receive.ReceiveID = dbo.WH_ReceiveMaterial.ReceiveID
                        WHERE(dbo.WH_ReceiveMaterial.MaterialName = N'Diesel Fuel')
                        GROUP BY dbo.WH_Receive.ReceiveNO
                        HAVING (MAX(dbo.WH_Receive.ReceiveDate) >= '" + StartDateTime.Value.ToString("yyyy-MM-dd") + @"') AND
                       (MAX(dbo.WH_Receive.ReceiveDate) <= '" + EndDateTime.Value.ToString("yyyy-MM-dd") + @"')
                        ORDER BY MAX(dbo.WH_Receive.ReceiveDate), dbo.WH_Receive.ReceiveNO";
            NewMethod(excel, 15, 3, sql);

            sql = @"SELECT dbo.WH_Issue.IssueNo
                        FROM      dbo.WH_Issue INNER JOIN
                        dbo.WH_IssueMaterial ON dbo.WH_Issue.IssueID = dbo.WH_IssueMaterial.IssueID
                        WHERE   (dbo.WH_IssueMaterial.MaterialName <> N'Diesel Fuel')
                        GROUP BY dbo.WH_Issue.IssueNo
                        HAVING   (MAX(dbo.WH_Issue.IssueDate) >=  '" + StartDateTime.Value.ToString("yyyy-MM-dd") + @"') AND
                       (MAX(dbo.WH_Issue.IssueDate) <= '" + EndDateTime.Value.ToString("yyyy-MM-dd") + @"')
                        ORDER BY MAX(dbo.WH_Issue.IssueDate),dbo.WH_Issue.IssueNo";
            NewMethod(excel, 18, 3, sql);

            sql = @"SELECT dbo.WH_Issue.IssueNo
                        FROM      dbo.WH_Issue INNER JOIN
                        dbo.WH_IssueMaterial ON dbo.WH_Issue.IssueID = dbo.WH_IssueMaterial.IssueID
                        WHERE   (dbo.WH_IssueMaterial.MaterialName = N'Diesel Fuel')
                        GROUP BY dbo.WH_Issue.IssueNo
                        HAVING   (MAX(dbo.WH_Issue.IssueDate) >=  '" + StartDateTime.Value.ToString("yyyy-MM-dd") + @"') AND
                       (MAX(dbo.WH_Issue.IssueDate) <= '" + EndDateTime.Value.ToString("yyyy-MM-dd") + @"')
                        ORDER BY MAX(dbo.WH_Issue.IssueDate),dbo.WH_Issue.IssueNo";
            NewMethod(excel, 19, 3, sql);

            sql = @"SELECT   TOP 100 PERCENT ReturnNo, ReturnDate
                        FROM      dbo.WH_Return
                        WHERE   (ReturnDate >= '" + StartDateTime.Value.ToString("yyyy-MM-dd") + @"') AND (ReturnDate <= '" + EndDateTime.Value.ToString("yyyy-MM-dd") + @"') 
                        ORDER BY ReturnDate,ReturnNo";
            NewMethod(excel, 20, 3, sql);

            sql = @"SELECT   TOP 100 PERCENT TransferWell2WellNo, TransferDate
                        FROM      dbo.WH_TransferWell2Well
                        WHERE   (TransferDate >= '" + StartDateTime.Value.ToString("yyyy-MM-dd") + @"') AND (TransferDate <= '" + EndDateTime.Value.ToString("yyyy-MM-dd") + @"')
                        ORDER BY TransferDate,TransferWell2WellNo";
            NewMethod(excel, 21, 3, sql);

            excel.InsertCell("Sheet1", 26, 0, EndDateTime.Value.ToString("yyyy-MM-dd"));

            excel.SaveExcel(folderBrowserDialog1.SelectedPath + @"SDT-WH " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
        }
        private void NewMethod(IExcelHelper excel, int row, int column, string sql)
        {
            string temp = "";
            DataTable dt = odbc.ExecuteDataTable(sql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (i == 0)
                {
                    excel.InsertCell("Sheet1", row, column, dt.Rows.Count.ToString());
                }
                temp += dt.Rows[i][0].ToString() + @" \ ";
                if (i == dt.Rows.Count - 1)
                {
                    excel.InsertCell("Sheet1", row, column + 1, temp);
                }
            }
        }
        private void ExportInventoryInfor(string SavePath)
        {
            IExcelHelper excel = new AsposeCellsExcelHelper(System.Windows.Forms.Application.StartupPath + @"\Template\各部门库存情况.xlsx");
            excel.InsertCell("Sheet1", 2, 1, EndDateTime.Value.ToString("yyyy-MM-dd") + "总库存");
            excel.InsertCell("Sheet1", 2, 2, "入库≥9个月库存(" + EndDateTime.Value.AddMonths(-9).ToString("yyyy-MM-dd") + ")前入库");
            excel.InsertCell("Sheet1", 2, 3, "入库≥12个月库存(" + EndDateTime.Value.AddMonths(-12).ToString("yyyy-MM-dd") + ")前入库");
            excel.InsertCell("Sheet1", 2, 4, "入库≥18个月库存(" + EndDateTime.Value.AddMonths(-18).ToString("yyyy-MM-dd") + ")前入库");
            excel.InsertCell("Sheet1", 2, 5, "入库≥24个月库存(" + EndDateTime.Value.AddMonths(-24).ToString("yyyy-MM-dd") + ")前入库");
            excel.SaveExcel(SavePath);
        }
        #endregion 私有方法


    }
}