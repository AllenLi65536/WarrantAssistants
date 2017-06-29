using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmReIssuable:Form
    {
        private DataTable dt = new DataTable();

        public FrmReIssuable() {
            InitializeComponent();
        }

        private void FrmReIssuable_Load(object sender, EventArgs e) {
            InitialGrid();
            LoadData();
        }

        private void InitialGrid() {
            dt.Columns.Add("WarrantID", typeof(string));
            dt.Columns.Add("WarrantName", typeof(string));
            dt.Columns.Add("UnderlyingID", typeof(string));
            dt.Columns.Add("UnderlyingName", typeof(string));
            dt.Columns.Add("Issuable", typeof(string));
            dt.Columns.Add("isReward", typeof(string));
            dt.Columns.Add("TraderAccount", typeof(string));
            dt.Columns.Add("IssueNum", typeof(double));
            dt.Columns.Add("SoldNum", typeof(double));
            dt.Columns.Add("Last1Sold", typeof(double));
            dt.Columns.Add("Last2Sold", typeof(double));
            dt.Columns.Add("Last3Sold", typeof(double));
            dt.Columns.Add("ExpiryDate", typeof(DateTime));

            dt.PrimaryKey = new DataColumn[] { dt.Columns["WarrantID"] };
            dataGridView1.DataSource = dt;

            dataGridView1.Columns[0].HeaderText = "權證代號";
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].HeaderText = "權證名稱";
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].HeaderText = "標的代號";
            dataGridView1.Columns[3].HeaderText = "標的名稱";

            dataGridView1.Columns[4].HeaderText = "是否可發";
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].HeaderText = "獎勵額度";
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[6].HeaderText = "交易員";
            dataGridView1.Columns[6].Width = 80;
            dataGridView1.Columns[7].HeaderText = "發行張數";
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].HeaderText = "流通在外";
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].HeaderText = "前1日(%)";
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].HeaderText = "前2日(%)";
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].HeaderText = "前3日(%)";
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].HeaderText = "到期日";

            dataGridView1.Columns[7].DefaultCellStyle.Format = "###,###";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "###,###";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void LoadData() {
            try {
                string sql = @"SELECT a.WarrantID
                                  ,a.WarrantName
                                  ,IsNull(b.UnderlyingID,'NA') UnderlyingID
                                  ,IsNull(b.UnderlyingName,'NA') UnderlyingName
                                  ,IsNull(c.Issuable,'NA') Issuable
                                  ,CASE WHEN b.isReward=1 THEN 'Y' ELSE 'N' END isReward
                                  ,IsNull(d.TraderAccount,'NA') TraderAccount
                                  ,a.IssueNum/1000 as IssueNum
                                  ,a.SoldNum/1000 as SoldNum
                                  ,a.Last1Sold
                                  ,a.Last2Sold
                                  ,a.Last3Sold
                                  ,IsNull(b.ExpiryDate,'') ExpiryDate
                              FROM [EDIS].[dbo].[WarrantReIssuable] a
                              LEFT JOIN [EDIS].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID
                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] c ON c.UnderlyingID=b.UnderlyingID
                              LEFT JOIN [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] d ON d.UID=b.UnderlyingID
                              ORDER BY b.UnderlyingID, a.WarrantID";

                dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                dataGridView1.DataSource = dt;

                /*DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    dr["權證代號"] = drv["WarrantID"].ToString();
                    dr["權證名稱"] = drv["WarrantName"].ToString();
                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    dr["標的名稱"] = drv["UnderlyingName"].ToString();
                    dr["是否可發"] = drv["Issuable"].ToString();
                    dr["獎勵額度"] = drv["isReward"].ToString();
                    dr["交易員"] = drv["TraderAccount"].ToString();
                    dr["發行張數"] = Convert.ToDouble(drv["IssueNum"]) / 1000;
                    dr["流通在外"] = Convert.ToDouble(drv["SoldNum"]) / 1000;
                    dr["前1日(%)"] = Convert.ToDouble(drv["Last1Sold"]);
                    dr["前2日(%)"] = Convert.ToDouble(drv["Last2Sold"]);
                    dr["前3日(%)"] = Convert.ToDouble(drv["Last3Sold"]);
                    dr["到期日"] = drv["ExpiryDate"];

                    dt.Rows.Add(dr);
                }*/
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "Issuable":
                    string cellValue = (string) e.Value;
                    if (cellValue == "N")
                        e.CellStyle.BackColor = Color.Coral;
                    if (cellValue == "NA")
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
                case "isReward":
                    if ((string) e.Value == "Y")
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "ExpiryDate":
                    if ((DateTime) e.Value < DateTime.Today.AddDays(7))
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
            }

            /*if (this.dataGridView1.Columns[e.ColumnIndex].Name == "是否可發") {
                string cellValue = (string) e.Value;
                if (cellValue == "N")
                    e.CellStyle.BackColor = Color.Coral;
                if (cellValue == "NA")
                    e.CellStyle.BackColor = Color.LightYellow;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "獎勵額度") {
                string cellValue = (string) e.Value;
                if ((string) e.Value == "Y")
                    e.CellStyle.BackColor = Color.Coral;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].Name == "到期日") {
                DateTime cellValue = (DateTime) e.Value;
                if ((DateTime) e.Value < DateTime.Today.AddDays(7))
                    e.CellStyle.BackColor = Color.LightYellow;
            }*/
        }
    }
}
