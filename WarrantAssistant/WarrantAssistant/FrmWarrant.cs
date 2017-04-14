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
    public partial class FrmWarrant : Form
    {
        private DataTable dataTable = new DataTable();
        private string enteredKey = "";

        public FrmWarrant()
        {
            InitializeComponent();
        }

        private void FrmWarrant_Load(object sender, EventArgs e)
        {
            InitialGrid();
            loadData();
            toolStripComboBox1.Items.Add("0005986");
            toolStripComboBox1.Items.Add("0007643");
            toolStripComboBox1.Items.Add("0008570");
            toolStripComboBox1.Items.Add("0008629");
            toolStripComboBox1.Items.Add("0008730");
            
        }

        private void InitialGrid()
        {
            dataTable.Columns.Add("WarrantID", typeof(string));
            dataTable.Columns.Add("WarrantName", typeof(string));
            dataTable.Columns.Add("UnderlyingID", typeof(string));
            dataTable.Columns.Add("UnderlyingName", typeof(string));
            dataTable.Columns.Add("Market", typeof(string));
            dataTable.Columns.Add("TraderID", typeof(string));
            dataTable.Columns.Add("WarrantType", typeof(string));
            dataTable.Columns.Add("K", typeof(double));
            dataTable.Columns.Add("T", typeof(int));
            dataTable.Columns.Add("exeRatio", typeof(double));
            dataTable.Columns.Add("HV", typeof(double));
            dataTable.Columns.Add("IV", typeof(double));
            dataTable.Columns.Add("IssuePrice", typeof(double));
            dataTable.Columns.Add("isReward", typeof(string));
            dataTable.Columns.Add("ExpiryDate", typeof(DateTime));
            dataTable.Columns.Add("IssueNum", typeof(double));
            dataTable.Columns.Add("FurthurIssueNum", typeof(double));

            dataTable.PrimaryKey = new DataColumn[] { dataTable.Columns["WarrantID"] };

            dataGridView1.DataSource = dataTable;

            dataGridView1.Columns[0].HeaderText = "權證代號";
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].HeaderText = "權證名稱";
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].HeaderText = "標的代號";
            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[3].HeaderText = "標的名稱";
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].HeaderText = "市場";
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].HeaderText = "交易員";
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[6].HeaderText = "權證型態";
            dataGridView1.Columns[6].Width = 110;
            dataGridView1.Columns[7].HeaderText = "履約價";
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].HeaderText = "存續期間";
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].HeaderText = "行使比例";
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].HeaderText = "避險Vol";
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].HeaderText = "發行Vol";
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].HeaderText = "發行價格";
            dataGridView1.Columns[12].Width = 80;
            dataGridView1.Columns[13].HeaderText = "獎勵額度";
            dataGridView1.Columns[13].Width = 80;
            dataGridView1.Columns[14].HeaderText = "到期日";
            dataGridView1.Columns[15].HeaderText = "發行張數";
            dataGridView1.Columns[16].HeaderText = "增額張數";

            dataGridView1.Columns[15].DefaultCellStyle.Format = "###,###";
            dataGridView1.Columns[16].DefaultCellStyle.Format = "###,###";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void loadData()
        {
            dataTable.Rows.Clear();

            string sql = @"SELECT [WarrantID]
                                 ,[WarrantName]
                                 ,[UnderlyingID]
                                 ,[UnderlyingName]
                                 ,[Market]
                                 ,[TraderID]
                                 ,[WarrantType]
                                 ,[K]
                                 ,[T]
                                 ,IsNull([exeRatio],0) [exeRatio]
                                 ,[HV]
                                 ,[IV]
                                 ,[IssuePrice]
                                 ,CASE WHEN [isReward]=0 THEN 'N' ELSE 'Y' END [isReward]
                                 ,[ExpiryDate]
                                 ,[IssueNum]/1000 [IssueNum]
                                 ,[FurthurIssueNum]/1000 [FurthurIssueNum]
                             FROM [EDIS].[dbo].[WarrantBasic]
                             ORDER BY Market desc, UnderlyingID, ExpiryDate";

            DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

            foreach (DataRowView drv in dv)
            {
                try
                {
                    DataRow dr = dataTable.NewRow();

                    dr["WarrantID"] = drv["WarrantID"].ToString();
                    dr["WarrantName"] = drv["WarrantName"].ToString();
                    dr["UnderlyingID"] = drv["UnderlyingID"].ToString();
                    dr["UnderlyingName"] = drv["UnderlyingName"].ToString();
                    dr["Market"] = drv["Market"].ToString();
                    dr["TraderID"] = drv["TraderID"].ToString();
                    dr["WarrantType"] = drv["WarrantType"].ToString();
                    dr["K"] = Convert.ToDouble(drv["K"]);
                    dr["T"] = Convert.ToInt32(drv["T"]);
                    dr["exeRatio"] = Convert.ToDouble(drv["exeRatio"]);
                    dr["HV"] = Convert.ToDouble(drv["HV"]);
                    dr["IV"] = Convert.ToDouble(drv["IV"]);
                    dr["IssuePrice"] = Convert.ToDouble(drv["IssuePrice"]);
                    dr["isReward"] = drv["isReward"].ToString();
                    dr["ExpiryDate"] = Convert.ToDateTime(drv["ExpiryDate"]);
                    dr["IssueNum"] = Convert.ToDouble(drv["IssueNum"]);
                    dr["FurthurIssueNum"] = Convert.ToDouble(drv["FurthurIssueNum"]);

                    dataTable.Rows.Add(dr);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "WarrantType")
            {
                string cellValue = (string)e.Value;
                if (cellValue != "一般型認購權證" && cellValue !="一般型認售權證")
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "isReward")
            {
                string cellValue = (string)e.Value;
                if (cellValue == "Y")
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "ExpiryDate")
            {
                DateTime cellValue = (DateTime)e.Value;
                if (cellValue < DateTime.Today.AddDays(3))
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "FurthurIssueNum")
            {
                double cellValue = (double)e.Value;
                if (cellValue > 0)
                    e.CellStyle.BackColor = Color.LightYellow;
            }
        }

        public void selectUnderlying(string UnderlyingID)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string uID = (string)dataGridView1.Rows[i].Cells[0].Value;
                if (uID == UnderlyingID)
                    dataGridView1.CurrentCell = dataGridView1.Rows[i - 1].Cells[0];
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
                    else if (e.KeyCode == Keys.B)
                        enteredKey += "B";
                    else if (e.KeyCode == Keys.C)
                        enteredKey += "C";
                    else if (e.KeyCode == Keys.P)
                        enteredKey += "P";
                    else
                        enteredKey += e.KeyCode.ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            loadDataByUnderlying();
        }

        private void loadDataByUnderlying()
        {
            string textBoxContent = toolStripTextBox1.Text;
            if (textBoxContent !="")
            {
                dataTable.Rows.Clear();

                string sql = @"SELECT [WarrantID]
                                     ,[WarrantName]
                                     ,[UnderlyingID]
                                     ,[UnderlyingName]
                                     ,[Market]
                                     ,[TraderID]
                                     ,[WarrantType]
                                     ,[K]
                                     ,[T]
                                     ,IsNull([exeRatio],0) [exeRatio]
                                     ,[HV]
                                     ,[IV]
                                     ,[IssuePrice]
                                     ,CASE WHEN [isReward]=0 THEN 'N' ELSE 'Y' END [isReward]
                                     ,[ExpiryDate]
                                     ,[IssueNum]/1000 [IssueNum]
                                     ,[FurthurIssueNum]/1000 [FurthurIssueNum]
                                 FROM [EDIS].[dbo].[WarrantBasic] ";

                sql+="WHERE [UnderlyingID]='"+ toolStripTextBox1.Text +"' ORDER BY ExpiryDate";
                  

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv)
                {
                    try
                    {
                        DataRow dr = dataTable.NewRow();

                        dr["WarrantID"] = drv["WarrantID"].ToString();
                        dr["WarrantName"] = drv["WarrantName"].ToString();
                        dr["UnderlyingID"] = drv["UnderlyingID"].ToString();
                        dr["UnderlyingName"] = drv["UnderlyingName"].ToString();
                        dr["Market"] = drv["Market"].ToString();
                        dr["TraderID"] = drv["TraderID"].ToString();
                        dr["WarrantType"] = drv["WarrantType"].ToString();
                        dr["K"] = Convert.ToDouble(drv["K"]);
                        dr["T"] = Convert.ToInt32(drv["T"]);
                        dr["exeRatio"] = Convert.ToDouble(drv["exeRatio"]);
                        dr["HV"] = Convert.ToDouble(drv["HV"]);
                        dr["IV"] = Convert.ToDouble(drv["IV"]);
                        dr["IssuePrice"] = Convert.ToDouble(drv["IssuePrice"]);
                        dr["isReward"] = drv["isReward"].ToString();
                        dr["ExpiryDate"] = Convert.ToDateTime(drv["ExpiryDate"]);
                        dr["IssueNum"] = Convert.ToDouble(drv["IssueNum"]);
                        dr["FurthurIssueNum"] = Convert.ToDouble(drv["FurthurIssueNum"]);

                        dataTable.Rows.Add(dr);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                toolStripTextBox1.Text = "";
            }
            else
                loadData();
        }

        private void loadDataByTrader()
        {
            string comboBoxContent = toolStripComboBox1.Text;
            if (comboBoxContent != "")
            {
                dataTable.Rows.Clear();

                string sql = @"SELECT [WarrantID]
                                     ,[WarrantName]
                                     ,[UnderlyingID]
                                     ,[UnderlyingName]
                                     ,[Market]
                                     ,[TraderID]
                                     ,[WarrantType]
                                     ,[K]
                                     ,[T]
                                     ,IsNull([exeRatio],0) [exeRatio]
                                     ,[HV]
                                     ,[IV]
                                     ,[IssuePrice]
                                     ,CASE WHEN [isReward]=0 THEN 'N' ELSE 'Y' END [isReward]
                                     ,[ExpiryDate]
                                     ,[IssueNum]/1000 [IssueNum]
                                     ,[FurthurIssueNum]/1000 [FurthurIssueNum]
                                 FROM [EDIS].[dbo].[WarrantBasic] ";

                sql += "WHERE [TraderID]='" + toolStripComboBox1.Text + "' ORDER BY UnderlyingID, ExpiryDate";


                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv)
                {
                    try
                    {
                        DataRow dr = dataTable.NewRow();

                        dr["WarrantID"] = drv["WarrantID"].ToString();
                        dr["WarrantName"] = drv["WarrantName"].ToString();
                        dr["UnderlyingID"] = drv["UnderlyingID"].ToString();
                        dr["UnderlyingName"] = drv["UnderlyingName"].ToString();
                        dr["Market"] = drv["Market"].ToString();
                        dr["TraderID"] = drv["TraderID"].ToString();
                        dr["WarrantType"] = drv["WarrantType"].ToString();
                        dr["K"] = Convert.ToDouble(drv["K"]);
                        dr["T"] = Convert.ToInt32(drv["T"]);
                        dr["exeRatio"] = Convert.ToDouble(drv["exeRatio"]);
                        dr["HV"] = Convert.ToDouble(drv["HV"]);
                        dr["IV"] = Convert.ToDouble(drv["IV"]);
                        dr["IssuePrice"] = Convert.ToDouble(drv["IssuePrice"]);
                        dr["isReward"] = drv["isReward"].ToString();
                        dr["ExpiryDate"] = Convert.ToDateTime(drv["ExpiryDate"]);
                        dr["IssueNum"] = Convert.ToDouble(drv["IssueNum"]);
                        dr["FurthurIssueNum"] = Convert.ToDouble(drv["FurthurIssueNum"]);

                        dataTable.Rows.Add(dr);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                toolStripComboBox1.Text = "";

            }
            else
                loadData();
        }

        private void toolStripTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                    loadDataByUnderlying();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        private void toolStripComboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                    loadDataByTrader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

 
    }
}
