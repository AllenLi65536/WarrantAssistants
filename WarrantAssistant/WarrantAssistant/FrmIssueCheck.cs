﻿using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmIssueCheck:Form
    {
        private DataTable dataTable;
        private string enteredKey = "";

        public FrmIssueCheck() {
            InitializeComponent();
        }

        private void InitialGrid() {
            dataGridView1.Columns[0].HeaderText = "標的代號";
            dataGridView1.Columns[1].HeaderText = "標的名稱";
            dataGridView1.Columns[2].HeaderText = "現金股利";
            dataGridView1.Columns[3].HeaderText = "股票股利";
            dataGridView1.Columns[4].HeaderText = "現金股利日期";
            dataGridView1.Columns[5].HeaderText = "股票股利日期";
            dataGridView1.Columns[6].HeaderText = "現增日期";
            dataGridView1.Columns[7].HeaderText = "處置結束日";
            dataGridView1.Columns[8].HeaderText = "注意次數";
            dataGridView1.Columns[9].HeaderText = "警示分數";
            dataGridView1.Columns[10].HeaderText = "損益(累計)";

            dataGridView1.Columns[10].DefaultCellStyle.Format = "N0";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }

        private void LoadData() {
            string sql = @"SELECT [UnderlyingID]
                                 ,[UnderlyingName]
                                 ,[CashDividend] as CashDividend
                                 ,[StockDividend] as StockDividend
                                 ,IsNull([CashDividendDate],'2030-12-31') CashDividendDate
                                 ,IsNull([StockDividendDate],'2030-12-31') StockDividendDate
                                 ,IsNull([PublicOfferingDate],'2030-12-31') PublicOfferingDate
                                 ,IsNull([DisposeEndDate],'1990-12-31') DisposeEndDate
                                 ,[WatchCount]
                                 ,[WarningScore]
                                 ,[AccNetIncome]
                              FROM [EDIS].[dbo].[WarrantIssueCheck]";
            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows) {
                row["CashDividend"] = Math.Round((double) row["CashDividend"], 3);
                row["StockDividend"] = Math.Round((double) row["StockDividend"], 3);
            }
        }

        private void FrmIssueCheck_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "CashDividendDate":
                case "StockDividendDate":
                case "PublicOfferingDate":
                    DateTime cellValue = (DateTime) e.Value;
                    if (cellValue != new DateTime(2030, 12, 31))
                        e.CellStyle.BackColor = Color.LightYellow;
                    if (cellValue == GlobalVar.globalParameter.nextTradeDate1)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "DisposeEndDate":
                    cellValue = (DateTime) e.Value;
                    if (cellValue != new DateTime(1990, 12, 31))
                        e.CellStyle.BackColor = Color.LightYellow;
                    if (cellValue.AddMonths(3) > DateTime.Today)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "WatchCount":
                    int cellValueInt = (int) e.Value;
                    if (cellValueInt == 1)
                        e.CellStyle.BackColor = Color.LightYellow;
                    else if (cellValueInt > 1)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "WarningScore":
                    if ((int) e.Value > 0)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "AccNetIncome":
                    if ((double) e.Value < 0)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
            }
        }

        public void SelectUnderlying(string underlyingID) {
            GlobalUtility.SelectUnderlying(underlyingID, dataGridView1);
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e) {
            try {
                if (e.KeyCode == Keys.Enter) {
                    SelectUnderlying(enteredKey);
                    e.Handled = true;                  
                    enteredKey = "";
                } else
                    GlobalUtility.KeyDecoder(e, ref enteredKey);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
