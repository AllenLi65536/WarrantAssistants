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
    public partial class FrmReIssue : Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt = new DataTable();
        private bool isEdit = false;
        public string userID;
        public string userName;
        private int applyCount = 0;

        public FrmReIssue()
        {
            InitializeComponent();
        }

        private void FrmReIssue_Load(object sender, EventArgs e)
        {
            toolStripLabel1.Text = "使用者: " + userName;
            toolStripLabel2.Text = "";
            InitialGrid();
            LoadData();
        }

        private void InitialGrid()
        {
            dt.Columns.Add("權證代號", typeof(string));
            dt.Columns.Add("增額張數", typeof(double));
            dt.Columns.Add("明日上市", typeof(string));
            dt.Columns.Add("確認", typeof(bool));
            dt.Columns["確認"].ReadOnly = false;
            dt.Columns.Add("獎勵", typeof(string));
            dt.Columns.Add("權證價格", typeof(double));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("今日額度(%)", typeof(double));
            dt.Columns.Add("獎勵額度", typeof(double));
            dt.Columns.Add("交易員", typeof(string));

            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].Hidden = true;

            ultraGrid1.DisplayLayout.Bands[0].Columns["增額張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Format = "###,###";

            ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;

            ultraGrid1.DisplayLayout.Bands[0].Columns["權證代號"].Width = 100;
            ultraGrid1.DisplayLayout.Bands[0].Columns["增額張數"].Width = 100;
            ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].Width = 100;
            ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Width = 70; 
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證價格"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].Width = 150;
            ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].Width = 73;

            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證價格"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].CellAppearance.BackColor = Color.LightGray;
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellAppearance.BackColor = Color.LightGray;

            ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            
            SetButton();
        }

        private void LoadData()
        {
            try
            {
                dt.Rows.Clear();
                string sql = @"SELECT a.WarrantID
                                  ,a.ReIssueNum
                                  ,a.MarketTmr
                                  ,CASE WHEN a.ConfirmChecked='Y' THEN 1 ELSE 0 END ConfirmChecked
                                  ,IsNull(c.MPrice,ISNull(c.BPrice,IsNull(c.APrice,0))) MPrice
                                  ,CASE WHEN b.isReward='1' THEN 'Y' ELSE 'N' END isReward
                                  ,b.UnderlyingID
                                  ,b.WarrantName
                                  ,IsNull(b.exeRatio,0) exeRatio
                                  ,(a.ReIssueNum*IsNull(b.exeRatio,0)) as EquivalentNum
                                  ,IsNull(d.IssuedPercent,0) IssuedPercent
                                  ,IsNull(d.RewardIssueCredit,0) RewardIssueCredit
                                  ,a.TraderID
                              FROM [EDIS].[dbo].[ReIssueTempList] a
                              LEFT JOIN [EDIS].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID
                              LEFT JOIN [EDIS].[dbo].[WarrantPrices] c ON a.WarrantID=c.CommodityID
                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] d ON b.UnderlyingID=d.UnderlyingID ";
                sql += "WHERE a.UserID='" + userID + "' ";
                sql += "ORDER BY a.MDate";

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv)
                {
                    DataRow dr = dt.NewRow();

                    string warrantID = drv["WarrantID"].ToString();
                    dr["權證代號"] = warrantID;
                    double reIssueNum = Convert.ToDouble(drv["ReIssueNum"]);
                    dr["增額張數"] = reIssueNum;
                    string marketTmr = drv["MarketTmr"].ToString();
                    dr["明日上市"] = marketTmr;
                    dr["確認"] = drv["ConfirmChecked"];
                    dr["獎勵"] = drv["isReward"].ToString();
                    double warrantPrice = 0.0;
                    warrantPrice = Convert.ToDouble(drv["MPrice"]);
                    dr["權證價格"] = warrantPrice;
                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    dr["權證名稱"] = drv["WarrantName"].ToString();
                    double cr = Convert.ToDouble(drv["exeRatio"]);
                    dr["行使比例"] = cr;
                    dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                    dr["今日額度(%)"] = Math.Round(Convert.ToDouble(drv["IssuedPercent"]), 2);
                    double rewardCredit = (double)drv["RewardIssueCredit"];
                    rewardCredit = Math.Floor(rewardCredit);
                    dr["獎勵額度"] = rewardCredit;
                    dr["交易員"] = drv["TraderID"].ToString();

                    dt.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateData()
        {
            try
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                conn.Close();

                string sql = @"INSERT INTO [ReIssueTempList] (SerialNum, WarrantID, ReIssueNum, MarketTmr, ConfirmChecked, TraderID, MDate, UserID) ";
                sql += "VALUES(@SerialNum, @WarrantID, @ReIssueNum, @MarketTmr, @ConfirmChecked, @TraderID, @MDate, @UserID)";
                List<SqlParameter> ps = new List<SqlParameter>();
                ps.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@WarrantID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ReIssueNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@MarketTmr", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ConfirmChecked", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@TraderID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                ps.Add(new SqlParameter("@UserID", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sql, ps);

                int i = 1;
                applyCount = 0;
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                {
                    string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "02" + i.ToString("0#");
                    string warrantID = r.Cells["權證代號"].Value.ToString();
                    double reIssueNum = Convert.ToDouble(r.Cells["增額張數"].Value);
                    string marketTmr = r.Cells["明日上市"].Value==DBNull.Value? "Y": r.Cells["明日上市"].Value.ToString();
                    string traderID = r.Cells["交易員"].Value == DBNull.Value ? userID : r.Cells["交易員"].Value.ToString();
                    bool confirmed = false;
                    confirmed = r.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["確認"].Value);
                    string confirmChecked = "N";
                    if (confirmed == true)
                    {
                        confirmChecked = "Y";
                        applyCount++;
                    }
                    else
                        confirmChecked = "N";

                    h.SetParameterValue("@SerialNum", serialNumber);
                    h.SetParameterValue("@WarrantID", warrantID);
                    h.SetParameterValue("@ReIssueNum", reIssueNum);
                    h.SetParameterValue("@MarketTmr", marketTmr);
                    h.SetParameterValue("@ConfirmChecked", confirmChecked);
                    h.SetParameterValue("@TraderID", traderID);
                    h.SetParameterValue("@MDate", DateTime.Now);
                    h.SetParameterValue("@UserID", userID);

                    h.ExecuteCommand();
                    i++;
                }

                h.Dispose();
                GlobalUtility.logInfo("Log", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔增額");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Log");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔增額");
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OfficiallyApply()
        {
            try
            {
                UpdateData();
                string sql1 = "DELETE FROM [EDIS].[dbo].[ReIssueOfficial] WHERE [UserID]='" + userID + "'";
                string sql2 = @"INSERT INTO [EDIS].[dbo].[ReIssueOfficial] ([SerialNum],[UnderlyingID],[WarrantID],[WarrantName],[exeRatio],[ReIssueNum],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT a.SerialNum, b.UnderlyingID, a.WarrantID, b.WarrantName, b.exeRatio, a.ReIssueNum, CASE WHEN b.isReward='1' THEN 'Y' ELSE 'N' END isReward, a.MarketTmr, a.TraderID, a.MDate, a.UserID
                                  FROM [EDIS].[dbo].[ReIssueTempList] a
                                  LEFT JOIN [EDIS].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID";
                sql2 += " WHERE a.[UserID]='" + userID + "' AND a.[ConfirmChecked]='Y'";
                string sql3 = "DELETE FROM [EDIS].[dbo].[ApplyTotalList] WHERE [UserID]='" + userID + "' AND [ApplyKind]='2'";
                string sql4 = @"INSERT INTO [EDIS].[dbo].[ApplyTotalList] ([ApplyKind],[SerialNum],[Market],[UnderlyingID],[WarrantName],[CR] ,[IssueNum],[EquivalentNum],[Credit],[RewardCredit],[Type],[CP],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT '2',a.SerialNum, b.Market, a.UnderlyingID, a.WarrantName, a.exeRatio, a.ReIssueNum, (a.exeRatio*a.ReIssueNum), b.IssueCredit, b.RewardIssueCredit, CASE WHEN SUBSTRING(c.WarrantType,1,2)='浮動' THEN '重設型' ELSE (CASE WHEN c.WarrantType='重設' THEN '牛熊證' ELSE '一般型' END) END, CASE WHEN SUBSTRING(c.WarrantType,LEN(c.WarrantType)-3,4)='熊證認售' OR SUBSTRING(c.WarrantType,LEN(c.WarrantType)-3,4)='認售權證' THEN 'P' ELSE 'C' END, a.UseReward,a.MarketTmr, a.TraderID, GETDATE(), a.UserID
                                FROM [EDIS].[dbo].[ReIssueOfficial] a
                                LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                                LEFT JOIN [EDIS].[dbo].[WarrantBasic] c ON a.WarrantID=c.WarrantID";
                sql4 += " WHERE a.[UserID]='" + userID + "'";

                SqlCommand cmd1 = new SqlCommand(sql1, conn);
                SqlCommand cmd2 = new SqlCommand(sql2, conn);
                SqlCommand cmd3 = new SqlCommand(sql3, conn);
                SqlCommand cmd4 = new SqlCommand(sql4, conn);

                conn.Open();
                cmd1.ExecuteNonQuery();
                cmd1.Dispose();
                cmd2.ExecuteNonQuery();
                cmd2.Dispose();
                cmd3.ExecuteNonQuery();
                cmd3.Dispose();
                cmd4.ExecuteNonQuery();
                cmd4.Dispose();
                conn.Close();

                toolStripLabel2.Text = DateTime.Now + "申請成功";
                GlobalUtility.logInfo("Info", GlobalVar.globalParameter.userID + " 申請" + applyCount + "檔權證增額");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Info");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 申請" + applyCount + "檔權證增額");
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton()
        {
            if (isEdit)
            {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                ultraGrid1.DisplayLayout.Bands[0].Columns["權證代號"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["增額張數"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證價格"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.AllowEdit;

                buttonEdit.Visible = false;
                buttonConfirm.Visible = true;
                buttonDelete.Visible = true;
                buttonCancel.Visible = true;
                toolStripButton1.Visible = false;
                toolStripSeparator2.Visible = false;

                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = true;

            }
            else
            {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].CellActivation = Activation.AllowEdit;

                ultraGrid1.DisplayLayout.Bands[0].Columns["權證代號"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["增額張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證價格"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellActivation = Activation.NoEdit;

                buttonEdit.Visible = true;
                buttonConfirm.Visible = false;
                buttonDelete.Visible = false;
                buttonCancel.Visible = false;
                toolStripButton1.Visible = true;
                toolStripSeparator2.Visible = true;

                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度(%)"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = false;
            }
        }

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;

            ValueList v;
            if (!e.Layout.ValueLists.Exists("MyValueList"))
            {
                v = e.Layout.ValueLists.Add("MyValueList");
                v.ValueListItems.Add("Y", "Y");
                v.ValueListItems.Add("N", "N");
            }
            e.Layout.Bands[0].Columns["明日上市"].ValueList = e.Layout.ValueLists["MyValueList"];
        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            isEdit = true;
            SetButton();
        }

        private void buttonConfirm_Click(object sender, EventArgs e)
        {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            isEdit = true;

            DialogResult result = MessageBox.Show("將全部刪除，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                conn.Close();
            }
            LoadData();
            SetButton();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (GlobalVar.globalParameter.userGroup == "FE")
            {
                OfficiallyApply();
                LoadData();
            }
            else
            {
                if (DateTime.Now.TimeOfDay.TotalMinutes > 630)
                    MessageBox.Show("超過交易所申報時間，欲改條件請洽行政組");
                else if (DateTime.Now.TimeOfDay.TotalMinutes > 570)
                {
                    DialogResult result = MessageBox.Show("超過約定的9:30了，已經告知組長及行政?", "逾時申請", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes)
                    {
                        OfficiallyApply();
                        LoadData();
                        GlobalUtility.logInfo("Announce", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證增額");
                        /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                        List<SqlParameter> psInfo = new List<SqlParameter>();
                        psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                        psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                        psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                        psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                        SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                        hInfo.SetParameterValue("@MDate", DateTime.Now);
                        hInfo.SetParameterValue("@InformationType", "Announce");
                        hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證增額");
                        hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                        hInfo.ExecuteCommand();
                        hInfo.Dispose();*/
                    }
                    else
                        LoadData();
                }
                else
                {
                    OfficiallyApply();
                    LoadData();
                }
            }
        }

        private void ultraGrid1_CellChange(object sender, CellEventArgs e)
        {
            if (e.Cell.Column.Key == "確認")
                ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e)
        {
            if (e.Cell.Column.Key == "權證代號")
            {
                string warrantID = e.Cell.Row.Cells["權證代號"].Value.ToString();
                string useReward = "";
                double warrantPrice = 0.0;
                string warrantName = "";
                string traderID = "";
                double cr = 0.0;

                string sqlTemp = @"SELECT CASE WHEN a.isReward='1' THEN 'Y' ELSE 'N' END isReward
		                                ,IsNull(b.MPrice,ISNull(b.BPrice,IsNull(b.APrice,0))) MPrice
		                                ,a.WarrantName
		                                ,a.exeRatio
                                        ,a.TraderID
                                  FROM [EDIS].[dbo].[WarrantBasic] a
                                  LEFT JOIN [EDIS].[dbo].[WarrantPrices] b ON a.WarrantID=b.CommodityID";
                sqlTemp += " WHERE a.WarrantID='" + warrantID + "'";
                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drTemp in dvTemp)
                {
                    useReward = drTemp["isReward"].ToString();
                    warrantPrice = Convert.ToDouble(drTemp["MPrice"]);
                    warrantName = drTemp["WarrantName"].ToString();
                    traderID = drTemp["TraderID"].ToString();
                    cr = Convert.ToDouble(drTemp["exeRatio"]);
                }
                e.Cell.Row.Cells["獎勵"].Value = useReward;
                e.Cell.Row.Cells["權證價格"].Value = warrantPrice;
                e.Cell.Row.Cells["權證名稱"].Value = warrantName;
                e.Cell.Row.Cells["行使比例"].Value = cr;
                e.Cell.Row.Cells["交易員"].Value = traderID;
            }
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show();
            }
        }

        private void ultraGrid1_BeforeRowsDeleted(object sender, BeforeRowsDeletedEventArgs e)
        {
            e.DisplayPromptMsg = false;
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("刪除此檔，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                ultraGrid1.ActiveRow.Delete();
                UpdateData();
            }
            LoadData();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string warrantID = "";
            string underlyingID = "";
            warrantID = e.Row.Cells["權證代號"].Value.ToString();
            underlyingID = e.Row.Cells["標的代號"].Value.ToString();

            string traderID = "NA";
            string issuable = "NA";
            string reissuable = "NA";

            string toolTip1 = "今日未達增額標準";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的發行檢查=N";
            string toolTip4 = "非此使用者所屬標的";

            string sqlTemp = "SELECT TraderID, IsNull(Issuable,'NA') Issuable FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp.Count > 0)
            {
                foreach (DataRowView drTemp in dvTemp)
                {
                    traderID = "000" + drTemp["TraderID"].ToString();
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
            if (!isEdit)
            {
                
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

                if (issuable != "NA" && traderID != userID)
                {
                    e.Row.Appearance.BackColor = Color.LightYellow;
                    e.Row.Cells["權證代號"].ToolTipText = toolTip4;
                }
            }
            
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e)
        {
            if (e.Cell.Column.Key == "權證代號")
            {
                FrmReIssuable frmReIssuable = null;

                foreach (Form iForm in Application.OpenForms)
                {
                    if (iForm.GetType() == typeof(FrmReIssuable))
                    {
                        frmReIssuable = (FrmReIssuable)iForm;
                        break;
                    }
                }

                if (frmReIssuable != null)
                    frmReIssuable.BringToFront();
                else
                {
                    frmReIssuable = new FrmReIssuable();
                    frmReIssuable.StartPosition = FormStartPosition.CenterScreen;
                    frmReIssuable.Show();
                }
            }
        }

        private void ultraGrid1_DoubleClickHeader(object sender, DoubleClickHeaderEventArgs e)
        {
            if (e.Header.Column.Key == "確認")
            {
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                {
                    r.Cells["確認"].Value = true;
                    UpdateData();
                    LoadData();
                }
            }
        }

        private void FrmReIssue_FormClosed(object sender, FormClosedEventArgs e)
        {
            UpdateData();
        }

    }
}
