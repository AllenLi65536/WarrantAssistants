using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmWarrant:Form
    {
        private DataTable dataTable;
        private string enteredKey = "";

        public FrmWarrant() {
            InitializeComponent();
        }

        private void FrmWarrant_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
            foreach (var item in GlobalVar.globalParameter.traders)
                toolStripComboBox1.Items.Add(item);

            /*toolStripComboBox1.Items.Add("0005986");
            toolStripComboBox1.Items.Add("0007643");
            toolStripComboBox1.Items.Add("0008570");
            //toolStripComboBox1.Items.Add("0008629");
            toolStripComboBox1.Items.Add("0008730");
            toolStripComboBox1.Items.Add("0010120");*/

        }

        private void InitialGrid() {
            dataGridView1.Columns[0].HeaderText = "權證代號";
            dataGridView1.Columns[1].HeaderText = "權證名稱";
            dataGridView1.Columns[2].HeaderText = "標的代號";
            dataGridView1.Columns[3].HeaderText = "標的名稱";
            dataGridView1.Columns[4].HeaderText = "市場";
            dataGridView1.Columns[5].HeaderText = "交易員";
            dataGridView1.Columns[6].HeaderText = "權證型態";
            dataGridView1.Columns[7].HeaderText = "履約價";
            dataGridView1.Columns[8].HeaderText = "存續期間";
            dataGridView1.Columns[9].HeaderText = "行使比例";
            dataGridView1.Columns[10].HeaderText = "避險Vol";
            dataGridView1.Columns[11].HeaderText = "發行Vol";
            dataGridView1.Columns[12].HeaderText = "發行價格";
            dataGridView1.Columns[13].HeaderText = "獎勵額度";
            dataGridView1.Columns[14].HeaderText = "到期日";
            dataGridView1.Columns[15].HeaderText = "發行張數";
            dataGridView1.Columns[16].HeaderText = "增額張數";

            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[6].Width = 110;
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].Width = 80;
            dataGridView1.Columns[13].Width = 80;

            dataGridView1.Columns[15].DefaultCellStyle.Format = "###,###";
            dataGridView1.Columns[16].DefaultCellStyle.Format = "###,###";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }

        private void LoadData() {
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

            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            dataGridView1.DataSource = dataTable;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "WarrantType":
                    string cellValue = (string) e.Value;
                    if (cellValue != "一般型認購權證" && cellValue != "一般型認售權證")
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
                case "isReward":
                    if ((string) e.Value == "Y")
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
                case "ExpiryDate":
                    if ((DateTime) e.Value < DateTime.Today.AddDays(3))
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
                    /*case "FurthurIssueNum":
                        if ((double) e.Value > 0)
                            e.CellStyle.BackColor = Color.LightYellow;
                        break;*/
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

        private void toolStripButton1_Click(object sender, EventArgs e) {
            loadDataByUnderlying();
        }

        private void loadDataByUnderlying() {
            string textBoxContent = toolStripTextBox1.Text;
            if (textBoxContent != "") {
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

                sql += "WHERE [UnderlyingID]='" + toolStripTextBox1.Text + "' ORDER BY ExpiryDate";

                dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                dataGridView1.DataSource = dataTable;                
                toolStripTextBox1.Text = "";
            } else
                LoadData();
        }

        private void loadDataByTrader() {
            string comboBoxContent = toolStripComboBox1.Text;
            if (comboBoxContent != "") {
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

                dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                dataGridView1.DataSource = dataTable;
                toolStripComboBox1.Text = "";
            } else
                LoadData();
        }

        private void toolStripTextBox1_KeyDown(object sender, KeyEventArgs e) {
            try {
                if (e.KeyCode == Keys.Enter)
                    loadDataByUnderlying();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e) {
            foreach (DataGridViewRow row in dataGridView1.Rows) {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        private void toolStripComboBox1_KeyDown(object sender, KeyEventArgs e) {
            try {
                if (e.KeyCode == Keys.Enter)
                    loadDataByTrader();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
