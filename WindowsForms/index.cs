﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Infrastructure;

namespace WindowsForms
{
    public partial class index : Form
    {
        ODBC odbc = new ODBC();
        public index()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ODBC odbc = new ODBC();
            DataTable dt = odbc.ExecuteDataTable("select * from [INVENTORY] where [Receive Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Receive Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'");
            dataGridView1.DataSource = dt;
        }

        private void 设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BaseOperation.operForm(new Settings());
        }

        private void index_Load(object sender, EventArgs e)
        {
            StartDateTime.Value = DateTime.Now.AddDays(-7);
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void materialInventryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportInventory(folderBrowserDialog1.SelectedPath + @"material inventry report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
            }
        }

        private void materialReceiveReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportReceiving(folderBrowserDialog1.SelectedPath + @"material Receive report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
            }
        }

        private void materialIssueReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportIssue(folderBrowserDialog1.SelectedPath + @"material issue report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
            }
        }

        private void materialWellToWellReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportTransferWellToWell(folderBrowserDialog1.SelectedPath + @"material WellToWell report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
            }
        }
        private void materialTransferBackReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExportTransferBack(folderBrowserDialog1.SelectedPath + @"material TransferBack report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xls");
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
            }
        }
        private void ExportInventory(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [INVENTORY]"; //where [Receive Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Receive Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            MyxlsExcelHelper excel = new MyxlsExcelHelper();
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            excel.ImportDataTable("ALL", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Production Dept.'");
            excel.ImportDataTable("Production Dept.", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Field Operation Dept.' or [Depart Desc]='Base camp'");
            excel.ImportDataTable("Field Operation Dept.", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Pipeline'");
            excel.ImportDataTable("Pipeline", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Development Dep.'");
            excel.ImportDataTable("Development Dep.", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Operation Dept.'");
            excel.ImportDataTable("Operation Dept.", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Engineering & Construction'");
            excel.ImportDataTable("Engineering & Construction", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='HSE Dept.'");
            excel.ImportDataTable("HSE Dept.", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='Administration Dept.'");
            excel.ImportDataTable("Administration Dept.", 8, dt);
            dt = odbc.ExecuteDataTable(sql + " where [Depart Desc]='P&L Dept.'");
            excel.ImportDataTable("P&L Dept.", 8, dt);
            excel.SaveExcel(SavePath);
        }
        private void ExportReceiving(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_ReceiveReport] where [Recv Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Recv Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            MyxlsExcelHelper excel = new MyxlsExcelHelper();
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            excel.ImportDataTable("sheet1", 8, dt);
            excel.SaveExcel(SavePath);
        }
        private void ExportIssue(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_IssueReport] where [Issue Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Issue Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            MyxlsExcelHelper excel = new MyxlsExcelHelper();
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            excel.ImportDataTable("sheet1", 8, dt);
            excel.SaveExcel(SavePath);
        }
        private void ExportTransferWellToWell(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_WellToWellReport] where [Transfer Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Transfer Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            MyxlsExcelHelper excel = new MyxlsExcelHelper();
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            excel.ImportDataTable("sheet1", 8, dt);
            excel.SaveExcel(SavePath);
        }
        private void ExportTransferBack(string SavePath)
        {
            //初始化数据库操作类
            DataTable dt;
            string sql = @"select * from [FLF_TransferBackReport] where [Transfer Date]>='" + StartDateTime.Value.ToString("yyyy-MM-dd") + "' and [Transfer Date]<='" + EndDateTime.Value.ToString("yyyy-MM-dd") + "'";
            //初始化excel操作类，新建Excel
            MyxlsExcelHelper excel = new MyxlsExcelHelper();
            //读取库存表
            dt = odbc.ExecuteDataTable(sql);
            excel.ImportDataTable("sheet1", 8, dt);
            excel.SaveExcel(SavePath);
        }

       
    }
}
