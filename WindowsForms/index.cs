using System;
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
            System.Environment.Exit(0);
        }

        private void materialInventryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                ExcelHelper excel = new ExcelHelper(Application.StartupPath + @"\Template\material inventry report.xlsx", folderBrowserDialog1.SelectedPath + @"material inventry report " + StartDateTime.Value.ToString("yyyy-MM-dd") + " to " + EndDateTime.Value.ToString("yyyy-MM-dd") + ".xlsx");
                DataTable dt = (DataTable)dataGridView1.DataSource;
                InsertExcel(dt, excel);
                excel.AddSheet("Production Dept.");
                int r = 1, c = 1;
                foreach (var rows in dt.AsEnumerable().Where(x => x.Field<string>("Depart Desc") == "Production Dept."))
                {
                    foreach (var columns in rows.Table.Columns)
                    {
                        excel.SetCells(r, c, columns.ToString());
                        c++;
                    }
                    r++;
                }
                excel.SaveAsFile();
                excel.Dispose();
            }
        }
        private void InsertExcel(DataTable dt, ExcelHelper excel)
        {
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    excel.SetCells(r + 8, 1, (r + 1).ToString());
                    excel.SetCells(r + 8, c + 2, dt.Rows[r][c].ToString());
                    excel.SetCells(r + 8, 12, "=TODAY() - J" + (r + 8) + "");
                }
            }
        }
    }
}
