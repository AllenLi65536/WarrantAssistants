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
    public partial class FrmReIssuable : Form
    {
        private DataTable dt = new DataTable();

        public FrmReIssuable()
        {
            InitializeComponent();
        }

        private void FrmReIssuable_Load(object sender, EventArgs e)
        {
            InitialGrid();
            LoadData();
        }

        private void InitialGrid()
        {
            dt.Columns.Add("權證代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("標的名稱", typeof(string));
            dt.Columns.Add("是否可發", typeof(string));
            dt.Columns.Add("獎勵額度", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("發行張數", typeof(double));
            dt.Columns.Add("流通在外", typeof(double));
            dt.Columns.Add("前1日(%)", typeof(double));
            dt.Columns.Add("前2日(%)", typeof(double));
            dt.Columns.Add("前3日(%)", typeof(double));
            dt.Columns.Add("到期日", typeof(DateTime));

            dt.PrimaryKey = new DataColumn[] { dt.Columns["權證代號"] };
            dataGridView1.DataSource = dt;

            dataGridView1.Columns["權證代號"].Width = 80;
            dataGridView1.Columns["權證名稱"].Width = 120;
            dataGridView1.Columns["是否可發"].Width = 80;
            dataGridView1.Columns["獎勵額度"].Width = 80;
            dataGridView1.Columns["交易員"].Width = 80;
            dataGridView1.Columns["發行張數"].Width = 80;
            dataGridView1.Columns["流通在外"].Width = 80;
            dataGridView1.Columns["前1日(%)"].Width = 80;
            dataGridView1.Columns["前2日(%)"].Width = 80;
            dataGridView1.Columns["前3日(%)"].Width = 80;

            dataGridView1.Columns["發行張數"].DefaultCellStyle.Format = "###,###";
            dataGridView1.Columns["流通在外"].DefaultCellStyle.Format = "###,###";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void LoadData()
        {
            try
            {
                dt.Rows.Clear();
                string sql = @"SELECT a.WarrantID
                                  ,a.WarrantName
                                  ,IsNull(b.UnderlyingID,'NA') UnderlyingID
                                  ,IsNull(b.UnderlyingName,'NA') UnderlyingName
                                  ,IsNull(c.Issuable,'NA') Issuable
                                  ,CASE WHEN b.isReward=1 THEN 'Y' ELSE 'N' END isReward
                                  ,IsNull(d.TraderAccount,'NA') TraderAccount
                                  ,a.IssueNum
                                  ,a.SoldNum
                                  ,a.Last1Sold
                                  ,a.Last2Sold
                                  ,a.Last3Sold
                                  ,IsNull(b.ExpiryDate,'') ExpiryDate
                              FROM [EDIS].[dbo].[WarrantReIssuable] a
                              LEFT JOIN [EDIS].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID
                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] c ON c.UnderlyingID=b.UnderlyingID
                              LEFT JOIN [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] d ON d.UID=b.UnderlyingID
                              ORDER BY b.UnderlyingID, a.WarrantID";
                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv)
                {
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
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
                if (this.dataGridView1.Columns[e.ColumnIndex].Name == "是否可發")
                {
                    string cellValue = (string)e.Value;
                    if (cellValue == "N")
                        e.CellStyle.BackColor = Color.Coral;
                    if (cellValue == "NA")
                        e.CellStyle.BackColor = Color.LightYellow;
                }

                if (this.dataGridView1.Columns[e.ColumnIndex].Name == "獎勵額度")
                {
                    string cellValue = (string)e.Value;
                    if (cellValue == "Y")
                        e.CellStyle.BackColor = Color.Coral;
                }

                if (this.dataGridView1.Columns[e.ColumnIndex].Name == "到期日")
                {
                    DateTime cellValue = (DateTime)e.Value;
                    if (cellValue < DateTime.Today.AddDays(7))
                        e.CellStyle.BackColor = Color.LightYellow;
                }
        }
    }
}
