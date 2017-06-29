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

        private void LoadData() {
            string sql = "SELECT [UnderlyingID], [UnderlyingName], [TraderID], [Market], [Issuable], [PutIssuable], IsNull([IssueCredit],0) [IssueCredit], IsNull([IssuedPercent],0) [IssuedPercent], IsNull([RewardIssueCredit],0) [RewardIssueCredit], CASE WHEN [AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome FROM [EDIS].[dbo].[WarrantUnderlyingSummary] ORDER BY Market desc, UnderlyingID";

            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows) {
                row["IssuedPercent"] = Math.Round((double) row["IssuedPercent"], 2);
            }
        }

        private void FrmUnderlyingSummary_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
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
                    e.Handled = true;
                    enteredKey = "";
                } else
                    GlobalUtility.KeyDecoder(e, ref enteredKey);
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
