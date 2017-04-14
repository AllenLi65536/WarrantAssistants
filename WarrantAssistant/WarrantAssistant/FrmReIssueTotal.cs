using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;

namespace WarrantAssistant
{
    public partial class FrmReIssueTotal : Form
    {
        private DataTable dt = new DataTable();
        private string userID = GlobalVar.globalParameter.userID;

        public FrmReIssueTotal()
        {
            InitializeComponent();
        }

        private void FrmReIssueTotal_Load(object sender, EventArgs e)
        {
            InitialGrid();
            LoadData();
        }

        private void InitialGrid()
        {
            dt.Columns.Add("序號", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("權證代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("張數", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("額度結果", typeof(double));
            dt.Columns.Add("今日額度(%)", typeof(double));
            dt.Columns.Add("獎勵額度", typeof(double));
            dt.Columns.Add("使用獎勵", typeof(string));
            dt.Columns.Add("明日上市", typeof(string));

            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Format = "###,###";

            ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證代號"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].Width = 150;
            ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].Width = 80;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].Width = 100;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Width = 100;
            ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].Width = 70;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證代號"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].CellActivation = Activation.NoEdit;

            ultraGrid1.DisplayLayout.Bands[0].Columns["序號"].Hidden = true;

        }

        private void LoadData()
        {
            try
            {
                dt.Rows.Clear();
                string sql = @"SELECT a.SerialNum
                                      ,SUBSTRING(a.TraderID,4,4) TraderID
                                      ,a.UnderlyingID
                                      ,a.WarrantID
                                      ,a.WarrantName
                                      ,a.exeRatio
                                      ,a.ReIssueNum
                                      ,c.EquivalentNum
                                      ,c.Result
                                      ,IsNull(b.IssuedPercent,0) IssuedPercent
                                      ,IsNull(b.RewardIssueCredit,0) RewardIssueCredit
                                      ,a.UseReward
                                      ,a.MarketTmr
                                  FROM [EDIS].[dbo].[ReIssueOfficial] a
                                  LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                                  LEFT JOIN [EDIS].[dbo].[ApplyTotalList] c ON a.SerialNum=c.SerialNum";

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                if (dv.Count > 0)
                {
                    foreach (DataRowView drv in dv)
                    {
                        DataRow dr = dt.NewRow();

                        dr["序號"] = drv["SerialNum"].ToString();
                        dr["交易員"] = drv["TraderID"].ToString();
                        dr["標的代號"] = drv["UnderlyingID"].ToString();
                        dr["權證代號"] = drv["WarrantID"].ToString();
                        dr["權證名稱"] = drv["WarrantName"].ToString();
                        dr["行使比例"] = Convert.ToDouble(drv["exeRatio"]);
                        dr["張數"] = Convert.ToDouble(drv["ReIssueNum"]);
                        dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                        dr["額度結果"] = drv["Result"];
                        dr["今日額度(%)"] = Math.Round(Convert.ToDouble(drv["IssuedPercent"]), 2);
                        double rewardCredit = (double)drv["RewardIssueCredit"];
                        rewardCredit = Math.Floor(rewardCredit);
                        dr["獎勵額度"] = rewardCredit;
                        dr["使用獎勵"] = drv["UseReward"].ToString();
                        dr["明日上市"] = drv["MarketTmr"].ToString();

                        dt.Rows.Add(dr);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            if (DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime)
            {
                double equivalentNum = Convert.ToDouble(e.Row.Cells["約當張數"].Value);
                double result = e.Row.Cells["額度結果"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["額度結果"].Value);

                if (result >= equivalentNum)
                {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.PaleGreen;
                }


            }
            string warrantID = "";
            string underlyingID = "";
            warrantID = e.Row.Cells["權證代號"].Value.ToString();
            underlyingID = e.Row.Cells["標的代號"].Value.ToString();

            string issuable = "NA";
            string reissuable = "NA";

            string toolTip1 = "今日未達增額標準";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的發行檢查=N";

            string sqlTemp = "SELECT IsNull(Issuable,'NA') Issuable FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp.Count > 0)
            {
                foreach (DataRowView drTemp in dvTemp)
                {
                    issuable = drTemp["Issuable"].ToString();
                }
            }
            string sqlTemp2 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [EDIS].[dbo].[WarrantReIssuable] WHERE WarrantID = '" + warrantID + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp2.Count > 0)
            {
                foreach (DataRowView drTemp2 in dvTemp2)
                {
                    reissuable = drTemp2["ReIssuable"].ToString();
                }
            }

            if (issuable == "NA")
            {
                e.Row.ToolTipText = toolTip2;
                e.Row.Appearance.ForeColor = Color.Red;
            }
            else if (issuable == "N")
            {
                e.Row.Cells["標的代號"].ToolTipText = toolTip3;
                e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
            }

            if (reissuable == "NA")
            {
                e.Row.Cells["權證代號"].ToolTipText = toolTip1;
                e.Row.Cells["權證代號"].Appearance.ForeColor = Color.Red;
            }

        }
    }
}
