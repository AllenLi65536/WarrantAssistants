namespace WarrantDataManager2._0
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.IssueCreditRefresh = new System.Windows.Forms.Button();
            this.IssueCheckRefresh = new System.Windows.Forms.Button();
            this.WarrantDataRefresh = new System.Windows.Forms.Button();
            this.UnderlyingDataRefresh = new System.Windows.Forms.Button();
            this.SummaryRefresh = new System.Windows.Forms.Button();
            this.PricesRefresh = new System.Windows.Forms.Button();
            this.UpdateAll = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.CleanApplyList = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // IssueCreditRefresh
            // 
            this.IssueCreditRefresh.Location = new System.Drawing.Point(19, 13);
            this.IssueCreditRefresh.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.IssueCreditRefresh.Name = "IssueCreditRefresh";
            this.IssueCreditRefresh.Size = new System.Drawing.Size(128, 45);
            this.IssueCreditRefresh.TabIndex = 1;
            this.IssueCreditRefresh.Text = "權證額度更新";
            this.IssueCreditRefresh.UseVisualStyleBackColor = true;
            this.IssueCreditRefresh.Click += new System.EventHandler(this.IssueCreditRefresh_Click);
            // 
            // IssueCheckRefresh
            // 
            this.IssueCheckRefresh.Location = new System.Drawing.Point(171, 13);
            this.IssueCheckRefresh.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.IssueCheckRefresh.Name = "IssueCheckRefresh";
            this.IssueCheckRefresh.Size = new System.Drawing.Size(128, 45);
            this.IssueCheckRefresh.TabIndex = 2;
            this.IssueCheckRefresh.Text = "發行檢查更新";
            this.IssueCheckRefresh.UseVisualStyleBackColor = true;
            this.IssueCheckRefresh.Click += new System.EventHandler(this.IssueCheckRefresh_Click);
            // 
            // WarrantDataRefresh
            // 
            this.WarrantDataRefresh.Location = new System.Drawing.Point(19, 76);
            this.WarrantDataRefresh.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.WarrantDataRefresh.Name = "WarrantDataRefresh";
            this.WarrantDataRefresh.Size = new System.Drawing.Size(128, 45);
            this.WarrantDataRefresh.TabIndex = 3;
            this.WarrantDataRefresh.Text = "權證資料更新";
            this.WarrantDataRefresh.UseVisualStyleBackColor = true;
            this.WarrantDataRefresh.Click += new System.EventHandler(this.WarrantDataRefresh_Click);
            // 
            // UnderlyingDataRefresh
            // 
            this.UnderlyingDataRefresh.Location = new System.Drawing.Point(171, 76);
            this.UnderlyingDataRefresh.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.UnderlyingDataRefresh.Name = "UnderlyingDataRefresh";
            this.UnderlyingDataRefresh.Size = new System.Drawing.Size(128, 45);
            this.UnderlyingDataRefresh.TabIndex = 4;
            this.UnderlyingDataRefresh.TabStop = false;
            this.UnderlyingDataRefresh.Text = "標的資料更新";
            this.UnderlyingDataRefresh.UseVisualStyleBackColor = true;
            this.UnderlyingDataRefresh.Click += new System.EventHandler(this.UnderlyingDataRefresh_Click);
            // 
            // SummaryRefresh
            // 
            this.SummaryRefresh.Location = new System.Drawing.Point(323, 13);
            this.SummaryRefresh.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.SummaryRefresh.Name = "SummaryRefresh";
            this.SummaryRefresh.Size = new System.Drawing.Size(128, 45);
            this.SummaryRefresh.TabIndex = 5;
            this.SummaryRefresh.Text = "Summary更新";
            this.SummaryRefresh.UseVisualStyleBackColor = true;
            this.SummaryRefresh.Click += new System.EventHandler(this.SummaryRefresh_Click);
            // 
            // PricesRefresh
            // 
            this.PricesRefresh.Location = new System.Drawing.Point(323, 76);
            this.PricesRefresh.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.PricesRefresh.Name = "PricesRefresh";
            this.PricesRefresh.Size = new System.Drawing.Size(128, 45);
            this.PricesRefresh.TabIndex = 6;
            this.PricesRefresh.Text = "價格更新";
            this.PricesRefresh.UseVisualStyleBackColor = true;
            this.PricesRefresh.Click += new System.EventHandler(this.PricesRefresh_Click);
            // 
            // UpdateAll
            // 
            this.UpdateAll.Location = new System.Drawing.Point(475, 13);
            this.UpdateAll.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.UpdateAll.Name = "UpdateAll";
            this.UpdateAll.Size = new System.Drawing.Size(128, 45);
            this.UpdateAll.TabIndex = 7;
            this.UpdateAll.Text = "更新全部";
            this.UpdateAll.UseVisualStyleBackColor = true;
            this.UpdateAll.Click += new System.EventHandler(this.UpdateAll_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.CleanApplyList);
            this.panel1.Controls.Add(this.UpdateAll);
            this.panel1.Controls.Add(this.PricesRefresh);
            this.panel1.Controls.Add(this.SummaryRefresh);
            this.panel1.Controls.Add(this.UnderlyingDataRefresh);
            this.panel1.Controls.Add(this.WarrantDataRefresh);
            this.panel1.Controls.Add(this.IssueCheckRefresh);
            this.panel1.Controls.Add(this.IssueCreditRefresh);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(625, 137);
            this.panel1.TabIndex = 8;
            // 
            // listBox1
            // 
            this.listBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 17;
            this.listBox1.Location = new System.Drawing.Point(0, 137);
            this.listBox1.Name = "listBox1";
            this.listBox1.ScrollAlwaysVisible = true;
            this.listBox1.Size = new System.Drawing.Size(625, 257);
            this.listBox1.TabIndex = 9;
            // 
            // CleanApplyList
            // 
            this.CleanApplyList.Location = new System.Drawing.Point(475, 76);
            this.CleanApplyList.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CleanApplyList.Name = "CleanApplyList";
            this.CleanApplyList.Size = new System.Drawing.Size(128, 45);
            this.CleanApplyList.TabIndex = 8;
            this.CleanApplyList.Text = "申請表清空";
            this.CleanApplyList.UseVisualStyleBackColor = true;
            this.CleanApplyList.Click += new System.EventHandler(this.CleanApplyList_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(625, 394);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "權證轉檔中心";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button IssueCreditRefresh;
        private System.Windows.Forms.Button IssueCheckRefresh;
        private System.Windows.Forms.Button WarrantDataRefresh;
        private System.Windows.Forms.Button UnderlyingDataRefresh;
        private System.Windows.Forms.Button SummaryRefresh;
        private System.Windows.Forms.Button PricesRefresh;
        private System.Windows.Forms.Button UpdateAll;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button CleanApplyList;

    }
}

