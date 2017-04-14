using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmIssueCheck : Form
    {
        private DataTable dataTable = new DataTable();
        private string enteredKey = "";

        public FrmIssueCheck()
        {
            InitializeComponent();
        }
        
        private void InitialGrid()
        {
            dataTable.Columns.Add("UnderlyingID", typeof(string));
            dataTable.Columns.Add("UnderlyingName", typeof(string));
            dataTable.Columns.Add("CashDividend", typeof(double));
            dataTable.Columns.Add("StockDividend", typeof(double));
            dataTable.Columns.Add("CashDividendDate", typeof(DateTime));
            dataTable.Columns.Add("StockDividendDate", typeof(DateTime));
            dataTable.Columns.Add("PublicOfferingDate", typeof(DateTime));
            dataTable.Columns.Add("DisposeEndDate", typeof(DateTime));
            dataTable.Columns.Add("WatchCount", typeof(int));
            dataTable.Columns.Add("WarningScore", typeof(int));
            dataTable.Columns.Add("AccNetIncome", typeof(double));

            dataTable.PrimaryKey = new DataColumn[] { dataTable.Columns["UnderlyingID"] };

            dataGridView1.DataSource = dataTable;

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

            dataGridView1.Columns[10].DefaultCellStyle.Format = "###,###";
            
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void loadData()
        {
            dataTable.Rows.Clear();

            string sql = @"SELECT [UnderlyingID]
                                 ,[UnderlyingName]
                                 ,[CashDividend]
                                 ,[StockDividend]
                                 ,IsNull([CashDividendDate],'2030-12-31') CashDividendDate
                                 ,IsNull([StockDividendDate],'2030-12-31') StockDividendDate
                                 ,IsNull([PublicOfferingDate],'2030-12-31') PublicOfferingDate
                                 ,IsNull([DisposeEndDate],'1990-12-31') DisposeEndDate
                                 ,[WatchCount]
                                 ,[WarningScore]
                                 ,[AccNetIncome]
                              FROM [EDIS].[dbo].[WarrantIssueCheck]";

            DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

            foreach (DataRowView drv in dv)
            {
                try
                {
                    DataRow dr = dataTable.NewRow();
                    dr["UnderlyingID"] = drv["UnderlyingID"].ToString();
                    dr["UnderlyingName"] = drv["UnderlyingName"].ToString();
                    dr["CashDividend"] = Math.Round(Convert.ToDouble(drv["CashDividend"]), 3);
                    dr["StockDividend"] = Math.Round(Convert.ToDouble(drv["StockDividend"]), 3);
                    dr["CashDividendDate"] = Convert.ToDateTime(drv["CashDividendDate"]);
                    dr["StockDividendDate"] = Convert.ToDateTime(drv["StockDividendDate"]);
                    dr["PublicOfferingDate"] = Convert.ToDateTime(drv["PublicOfferingDate"]);
                    dr["DisposeEndDate"] = Convert.ToDateTime(drv["DisposeEndDate"]);
                    dr["WatchCount"]=Convert.ToInt32(drv["WatchCount"]);
                    dr["WarningScore"]=Convert.ToInt32(drv["WarningScore"]);
                    dr["AccNetIncome"] = Convert.ToDouble(drv["AccNetIncome"]);
                    dataTable.Rows.Add(dr);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void FrmIssueCheck_Load(object sender, EventArgs e)
        {
            InitialGrid();
            loadData();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "CashDividendDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                DateTime notTheDate = new DateTime(2030, 12, 31);
                if (cellValue != notTheDate)
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "StockDividendDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                DateTime notTheDate = new DateTime(2030, 12, 31);
                if (cellValue != notTheDate)
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "PublicOfferingDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                DateTime notTheDate = new DateTime(2030, 12, 31);
                if (cellValue != notTheDate)
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "DisposeEndDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                DateTime notTheDate = new DateTime(1990, 12, 31);
                if (cellValue != notTheDate)
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "CashDividendDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                if (cellValue == GlobalVar.globalParameter.nextTradeDate1)
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "StockDividendDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                if (cellValue == GlobalVar.globalParameter.nextTradeDate1)
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "PublicOfferingDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                if (cellValue == GlobalVar.globalParameter.nextTradeDate1)
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "DisposeEndDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                if (cellValue.AddMonths(3) > DateTime.Today)
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "WatchCount")
            {
                int cellValue = (int)e.Value;
                if (cellValue > 0)
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "WatchCount")
            {
                int cellValue = (int)e.Value;
                if (cellValue > 1)
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "WarningScore")
            {
                int cellValue = (int)e.Value;
                if (cellValue > 0)
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "AccNetIncome")
            {
                double cellValue = (double)e.Value;
                if (cellValue < 0)
                    e.CellStyle.BackColor = Color.Coral;
            }


        }

        public void selectUnderlying(string UnderlyingID)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string uID = (string)dataGridView1.Rows[i].Cells[0].Value;
                if (uID == UnderlyingID)
                    dataGridView1.CurrentCell = dataGridView1.Rows[i ].Cells[0];
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    selectUnderlying(enteredKey);
                    enteredKey = "";
                }
                else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
                {
                    if (enteredKey.Length > 0)
                        enteredKey = enteredKey.Substring(0, enteredKey.Length - 1);
                }
                else if (e.KeyCode == Keys.Escape)
                    enteredKey = "";
                else
                {
                    if (e.KeyCode == Keys.NumPad0 || e.KeyCode == Keys.D0)
                        enteredKey += "0";
                    else if (e.KeyCode == Keys.NumPad1 || e.KeyCode == Keys.D1)
                        enteredKey += "1";
                    else if (e.KeyCode == Keys.NumPad2 || e.KeyCode == Keys.D2)
                        enteredKey += "2";
                    else if (e.KeyCode == Keys.NumPad3 || e.KeyCode == Keys.D3)
                        enteredKey += "3";
                    else if (e.KeyCode == Keys.NumPad4 || e.KeyCode == Keys.D4)
                        enteredKey += "4";
                    else if (e.KeyCode == Keys.NumPad5 || e.KeyCode == Keys.D5)
                        enteredKey += "5";
                    else if (e.KeyCode == Keys.NumPad6 || e.KeyCode == Keys.D6)
                        enteredKey += "6";
                    else if (e.KeyCode == Keys.NumPad7 || e.KeyCode == Keys.D7)
                        enteredKey += "7";
                    else if (e.KeyCode == Keys.NumPad8 || e.KeyCode == Keys.D8)
                        enteredKey += "8";
                    else if (e.KeyCode == Keys.NumPad9 || e.KeyCode == Keys.D9)
                        enteredKey += "9";
                    else
                        enteredKey += e.KeyCode.ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
