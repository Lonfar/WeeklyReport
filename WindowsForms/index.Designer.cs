﻿namespace WindowsForms
{
    partial class index
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.StartDateTime = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.EndDateTime = new System.Windows.Forms.DateTimePicker();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.分开导出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.materialInventryReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.materialIssueReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.materialReceiveReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.materialTransferBackReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.materialWellToWellReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sDTWHToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.各部门库存情况ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.全部导出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label2 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.StartDateTime);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.EndDateTime);
            this.splitContainer1.Panel1.Controls.Add(this.menuStrip1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.label2);
            this.splitContainer1.Size = new System.Drawing.Size(324, 125);
            this.splitContainer1.SplitterDistance = 25;
            this.splitContainer1.TabIndex = 0;
            // 
            // StartDateTime
            // 
            this.StartDateTime.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.StartDateTime.Location = new System.Drawing.Point(88, 4);
            this.StartDateTime.Name = "StartDateTime";
            this.StartDateTime.Size = new System.Drawing.Size(98, 21);
            this.StartDateTime.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(192, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "To";
            // 
            // EndDateTime
            // 
            this.EndDateTime.CustomFormat = "";
            this.EndDateTime.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.EndDateTime.Location = new System.Drawing.Point(215, 4);
            this.EndDateTime.Name = "EndDateTime";
            this.EndDateTime.Size = new System.Drawing.Size(97, 21);
            this.EndDateTime.TabIndex = 1;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(324, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.分开导出ToolStripMenuItem,
            this.全部导出ToolStripMenuItem,
            this.toolStripSeparator2,
            this.退出ToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(44, 21);
            this.toolStripMenuItem1.Text = "选项";
            // 
            // 分开导出ToolStripMenuItem
            // 
            this.分开导出ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.materialInventryReportToolStripMenuItem,
            this.materialIssueReportToolStripMenuItem,
            this.materialReceiveReportToolStripMenuItem,
            this.materialTransferBackReportToolStripMenuItem,
            this.materialWellToWellReportToolStripMenuItem,
            this.sDTWHToolStripMenuItem,
            this.各部门库存情况ToolStripMenuItem});
            this.分开导出ToolStripMenuItem.Name = "分开导出ToolStripMenuItem";
            this.分开导出ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.分开导出ToolStripMenuItem.Text = "分开导出";
            // 
            // materialInventryReportToolStripMenuItem
            // 
            this.materialInventryReportToolStripMenuItem.Name = "materialInventryReportToolStripMenuItem";
            this.materialInventryReportToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.materialInventryReportToolStripMenuItem.Text = "material inventry report";
            this.materialInventryReportToolStripMenuItem.Click += new System.EventHandler(this.materialInventryReportToolStripMenuItem_Click);
            // 
            // materialIssueReportToolStripMenuItem
            // 
            this.materialIssueReportToolStripMenuItem.Name = "materialIssueReportToolStripMenuItem";
            this.materialIssueReportToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.materialIssueReportToolStripMenuItem.Text = "material issue report";
            this.materialIssueReportToolStripMenuItem.Click += new System.EventHandler(this.materialIssueReportToolStripMenuItem_Click);
            // 
            // materialReceiveReportToolStripMenuItem
            // 
            this.materialReceiveReportToolStripMenuItem.Name = "materialReceiveReportToolStripMenuItem";
            this.materialReceiveReportToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.materialReceiveReportToolStripMenuItem.Text = "material Receive report";
            this.materialReceiveReportToolStripMenuItem.Click += new System.EventHandler(this.materialReceiveReportToolStripMenuItem_Click);
            // 
            // materialTransferBackReportToolStripMenuItem
            // 
            this.materialTransferBackReportToolStripMenuItem.Name = "materialTransferBackReportToolStripMenuItem";
            this.materialTransferBackReportToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.materialTransferBackReportToolStripMenuItem.Text = "material TransferBack report";
            this.materialTransferBackReportToolStripMenuItem.Click += new System.EventHandler(this.materialTransferBackReportToolStripMenuItem_Click);
            // 
            // materialWellToWellReportToolStripMenuItem
            // 
            this.materialWellToWellReportToolStripMenuItem.Name = "materialWellToWellReportToolStripMenuItem";
            this.materialWellToWellReportToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.materialWellToWellReportToolStripMenuItem.Text = "material WellToWell report";
            this.materialWellToWellReportToolStripMenuItem.Click += new System.EventHandler(this.materialWellToWellReportToolStripMenuItem_Click);
            // 
            // sDTWHToolStripMenuItem
            // 
            this.sDTWHToolStripMenuItem.Name = "sDTWHToolStripMenuItem";
            this.sDTWHToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.sDTWHToolStripMenuItem.Text = "SDT-WH";
            this.sDTWHToolStripMenuItem.Click += new System.EventHandler(this.sDTWHToolStripMenuItem_Click);
            // 
            // 各部门库存情况ToolStripMenuItem
            // 
            this.各部门库存情况ToolStripMenuItem.Name = "各部门库存情况ToolStripMenuItem";
            this.各部门库存情况ToolStripMenuItem.Size = new System.Drawing.Size(244, 22);
            this.各部门库存情况ToolStripMenuItem.Text = "各部门库存情况";
            this.各部门库存情况ToolStripMenuItem.Click += new System.EventHandler(this.各部门库存情况ToolStripMenuItem_Click);
            // 
            // 全部导出ToolStripMenuItem
            // 
            this.全部导出ToolStripMenuItem.Name = "全部导出ToolStripMenuItem";
            this.全部导出ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.全部导出ToolStripMenuItem.Text = "全部导出";
            this.全部导出ToolStripMenuItem.Click += new System.EventHandler(this.全部导出ToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(121, 6);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(239, 60);
            this.label2.TabIndex = 0;
            this.label2.Text = "周报生成工具：\r\n1、先选定“起始日期”\r\n2、在选定”终止日期“\r\n3、“选项”中有需要生成的Excel\r\n4、选择盘符的根目录即可，不要选择文件夹";
            // 
            // index
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(324, 125);
            this.Controls.Add(this.splitContainer1);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "index";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "周报生成";
            this.Load += new System.EventHandler(this.index_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem 分开导出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem materialInventryReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem materialIssueReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem materialReceiveReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem materialTransferBackReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem materialWellToWellReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sDTWHToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 各部门库存情况ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 全部导出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.DateTimePicker StartDateTime;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker EndDateTime;
        private System.Windows.Forms.Label label2;
    }
}

