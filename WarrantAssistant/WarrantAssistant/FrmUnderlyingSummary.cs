using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WarrantAssistant
{
    public partial class FrmUnderlyingSummary:Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dataTable = new DataTable();
        private string enteredKey = "";

        public FrmUnderlyingSummary() {
            InitializeComponent();
        }

        private void InitialGrid() {
            dataTable.Columns.Add("UnderlyingID", typeof(string));
            dataTable.Columns.Add("UnderlyingName", typeof(string));
            dataTable.Columns.Add("TraderID", typeof(string));
            dataTable.Columns.Add("Market", typeof(string));
            dataTable.Columns.Add("Issuable", typeof(string));
            dataTable.Columns.Add("PutIssuable", typeof(string));
            dataTable.Columns.Add("IssuedPercent", typeof(double));
            dataTable.Columns.Add("IssueCredit", typeof(double));
            dataTable.Columns.Add("RewardIssueCredit", typeof(double));
            dataTable.Columns.Add("AccNetIncome", typeof(string));

            dataTable.PrimaryKey = new DataColumn[] { dataTable.Columns["UnderlyingID"] };

            dataGridView1.DataSource = dataTable;

            dataGridView1.Columns[0].HeaderText = "標的代號";
            dataGridView1.Columns[1].HeaderText = "標的名稱";
            dataGridView1.Columns[2].HeaderText = "交易員";
            dataGridView1.Columns[3].HeaderText = "市場";
            dataGridView1.Columns[4].HeaderText = "是否可發";
            dataGridView1.Columns[5].HeaderText = "Put發行檢查";
            dataGridView1.Columns[7].HeaderText = "今日額度";
            dataGridView1.Columns[6].HeaderText = "已發行(%)";
            dataGridView1.Columns[8].HeaderText = "獎勵額度";
            dataGridView1.Columns[9].HeaderText = "是否虧損";

            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].Width = 80;

            dataGridView1.Columns[7].DefaultCellStyle.Format = "###,###";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "###,###";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void loadData() {
            string sql = "SELECT [UnderlyingID], [UnderlyingName], [TraderID], [Market], [Issuable], [PutIssuable], IsNull([IssueCredit],0) [IssueCredit], IsNull([IssuedPercent],0) [IssuedPercent], IsNull([RewardIssueCredit],0) [RewardIssueCredit], CASE WHEN [AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome FROM [EDIS].[dbo].[WarrantUnderlyingSummary] ORDER BY Market desc, UnderlyingID";

            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows) {
                row["IssuedPercent"] = Math.Round((double) row["IssuedPercent"], 2);
            }
            /*DataTable dv = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            foreach (DataRow drv in dv.Rows) {
                try {
                    DataRow dr = dataTable.NewRow();
                    dr["UnderlyingID"] = drv["UnderlyingID"].ToString();
                    dr["UnderlyingName"] = drv["UnderlyingName"].ToString();
                    dr["TraderID"] = drv["TraderID"].ToString();
                    dr["Market"] = drv["Market"].ToString();
                    dr["Issuable"] = drv["Issuable"].ToString();
                    dr["PutIssuable"] = drv["PutIssuable"].ToString();
                    dr["IssueCredit"] = Math.Floor(Convert.ToDouble(drv["IssueCredit"]));
                    dr["IssuedPercent"] = Math.Round(Convert.ToDouble(drv["IssuedPercent"]), 2);
                    dr["RewardIssueCredit"] = Math.Floor(Convert.ToDouble(drv["RewardIssueCredit"]));
                    dr["AccNetIncome"] = drv["AccNetIncome"].ToString();
                    dataTable.Rows.Add(dr);

                } catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            }*/
        }

        private void FrmUnderlyingSummary_Load(object sender, EventArgs e) {
            InitialGrid();
            loadData();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "Issuable":
                case "PutIssuable":
                    if ((string) e.Value == "N")
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "IssueCredit":
                    if ((double) e.Value < 0)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "AccNetIncome":
                    if ((string) e.Value == "Y")
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
                    enteredKey = "";
                } else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) {
                    if (enteredKey.Length > 0)
                        enteredKey = enteredKey.Substring(0, enteredKey.Length - 1);
                } else if (e.KeyCode == Keys.Escape)
                    enteredKey = "";
                else {
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

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e) {
            string target = (string) dataGridView1.Rows[e.RowIndex].Cells[0].Value;
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "Issuable":
                    GlobalUtility.MenuItemClick<FrmIssueCheck>().SelectUnderlying(target);
                    break;
                case "PutIssuable":
                    GlobalUtility.MenuItemClick<FrmIssueCheckPut>().SelectUnderlying(target);
                    break;
            }
        }
    }
}
