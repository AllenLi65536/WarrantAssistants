using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmIssueCheckPut:Form
    {
        private DataTable dataTable;
        private string enteredKey = "";
        public FrmIssueCheckPut() {
            InitializeComponent();
        }

        private void FrmIssueCheckPut_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {
            dataGridView1.Columns[0].HeaderText = "標的代號";
            dataGridView1.Columns[1].HeaderText = "標的名稱";
            dataGridView1.Columns[2].HeaderText = "台灣50成分股";
            dataGridView1.Columns[3].HeaderText = "本益比";
            dataGridView1.Columns[4].HeaderText = "過去一年損益";
            dataGridView1.Columns[5].HeaderText = "股價";
            dataGridView1.Columns[6].HeaderText = "前一季股價";
            dataGridView1.Columns[7].HeaderText = "前一年股價";
            dataGridView1.Columns[8].HeaderText = "季報酬";
            dataGridView1.Columns[9].HeaderText = "年報酬";

            dataGridView1.Columns[4].DefaultCellStyle.Format = "N0";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }

        private void LoadData() {
            string sql = @"SELECT [UnderlyingID]
                                 ,[UnderlyingName]
                                 ,[IsTW50Stocks]
                                 ,[PERatio]
                                 ,[SumEarning]
                                 ,[Price]
                                 ,[PriceQuarter]
                                 ,[PriceYear]
                                 ,[ReturnQuarter]
                                 ,[ReturnYear]
                             FROM [EDIS].[dbo].[WarrantIssueCheckPut]";
            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows) {
                row["ReturnQuarter"] = Math.Round((double) row["ReturnQuarter"], 2);
                row["ReturnYear"] = Math.Round((double) row["ReturnYear"], 2);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "PERatio":
                    if ((double) e.Value > 40)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "SumEarning":
                    if ((double) e.Value < 0)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "ReturnQuarter":
                    if ((double) e.Value > 0.5)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "ReturnYear":
                    if ((double) e.Value > 1)
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
