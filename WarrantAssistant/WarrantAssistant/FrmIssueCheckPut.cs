using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmIssueCheckPut:Form
    {
        private DataTable dataTable = new DataTable();
        private string enteredKey = "";

        public FrmIssueCheckPut() {
            InitializeComponent();
        }

        private void FrmIssueCheckPut_Load(object sender, EventArgs e) {
            InitialGrid();
            loadData();
        }

        private void InitialGrid() {
            dataTable.Columns.Add("UnderlyingID", typeof(string));
            dataTable.Columns.Add("UnderlyingName", typeof(string));
            dataTable.Columns.Add("IsTW50Stocks", typeof(string));
            dataTable.Columns.Add("PERatio", typeof(double));
            dataTable.Columns.Add("SumEarning", typeof(double));
            dataTable.Columns.Add("Price", typeof(double));
            dataTable.Columns.Add("PriceQuarter", typeof(double));
            dataTable.Columns.Add("PriceYear", typeof(double));
            dataTable.Columns.Add("ReturnQuarter", typeof(double));
            dataTable.Columns.Add("ReturnYear", typeof(double));

            dataTable.PrimaryKey = new DataColumn[] { dataTable.Columns["UnderlyingID"] };

            dataGridView1.DataSource = dataTable;

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

            dataGridView1.Columns[4].DefaultCellStyle.Format = "###,###";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void loadData() {
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

            /*DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

            foreach (DataRowView drv in dv) {
                try {
                    DataRow dr = dataTable.NewRow();

                    dr["UnderlyingID"] = drv["UnderlyingID"].ToString();
                    dr["UnderlyingName"] = drv["UnderlyingName"].ToString();
                    dr["IsTW50Stocks"] = drv["IsTW50Stocks"].ToString();
                    dr["PERatio"] = Convert.ToDouble(drv["PERatio"]);
                    dr["SumEarning"] = Convert.ToDouble(drv["SumEarning"]);
                    dr["Price"] = Convert.ToDouble(drv["Price"]);
                    dr["PriceQuarter"] = Convert.ToDouble(drv["PriceQuarter"]);
                    dr["PriceYear"] = Convert.ToDouble(drv["PriceYear"]);
                    dr["ReturnQuarter"] = Math.Round(Convert.ToDouble(drv["ReturnQuarter"]), 2);
                    dr["ReturnYear"] = Math.Round(Convert.ToDouble(drv["ReturnYear"]), 2);

                    dataTable.Rows.Add(dr);

                } catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            }*/
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
            /*for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                string uID = (string) dataGridView1.Rows[i].Cells[0].Value;
                if (uID == underlyingID)
                    dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
            }*/
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
    }
}
