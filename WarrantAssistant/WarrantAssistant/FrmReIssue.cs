using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using EDLib.SQL;

namespace WarrantAssistant
{
    public partial class FrmReIssue:Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt = new DataTable();
        private bool isEdit = false;
        public string userID = GlobalVar.globalParameter.userID;
        public string userName = GlobalVar.globalParameter.userName;
        private int applyCount = 0;

        public FrmReIssue() {
            InitializeComponent();
        }

        private void FrmReIssue_Load(object sender, EventArgs e) {
            toolStripLabel1.Text = "使用者: " + userName;
            toolStripLabel2.Text = "";
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {
            /*dt.Columns.Add("WarrantID", typeof(string));
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
            ultraGrid1.DataSource = dt;*/

            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            band0.Columns["TraderID"].Hidden = true;

            band0.Columns["ReIssueNum"].Format = "###,###";
            band0.Columns["EquivalentNum"].Format = "###,###";
            band0.Columns["RewardIssueCredit"].Format = "###,###";

            band0.Columns["MarketTmr"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;

            band0.Columns["WarrantID"].Width = 100;
            band0.Columns["ReIssueNum"].Width = 100;
            band0.Columns["MarketTmr"].Width = 100;
            band0.Columns["ConfirmChecked"].Width = 70;
            band0.Columns["isReward"].Width = 70;
            band0.Columns["MPrice"].Width = 70;
            band0.Columns["UnderlyingID"].Width = 70;
            band0.Columns["WarrantName"].Width = 150;
            band0.Columns["exeRatio"].Width = 73;

            band0.Columns["isReward"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["MPrice"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["UnderlyingID"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["WarrantName"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["exeRatio"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["EquivalentNum"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["IssuedPercent"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["RewardIssueCredit"].CellAppearance.BackColor = Color.LightGray;

            band0.Columns["MarketTmr"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            SetButton();
        }

        private void LoadData() {
            try {                
                string sql = @"SELECT a.WarrantID
                                  ,a.ReIssueNum
                                  ,a.MarketTmr
                                  ,CASE WHEN a.ConfirmChecked='Y' THEN 1 ELSE 0 END ConfirmChecked                                  
                                  ,CASE WHEN b.isReward='1' THEN 'Y' ELSE 'N' END isReward
                                  ,IsNull(c.MPrice,ISNull(c.BPrice,IsNull(c.APrice,0))) MPrice
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

                dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                ultraGrid1.DataSource = dt;

                dt.Columns[0].Caption = "權證代號";
                dt.Columns[1].Caption = "增額張數";
                dt.Columns[2].Caption = "明日上市";
                dt.Columns[3].Caption = "確認";
                dt.Columns[4].Caption = "獎勵";
                dt.Columns[5].Caption = "權證價格";
                dt.Columns[6].Caption = "標的代號";
                dt.Columns[7].Caption = "權證名稱";
                dt.Columns[8].Caption = "行使比例";
                dt.Columns[9].Caption = "約當張數";
                dt.Columns[10].Caption = "今日額度(%)";
                dt.Columns[11].Caption = "獎勵額度";
                dt.Columns[12].Caption = "交易員";

                /*DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    string warrantID = drv["WarrantID"].ToString();
                    dr["WarrantID"] = warrantID;
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
                    dr["exeRatio"] = cr;
                    dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                    dr["IssuedPercent"] = Math.Round(Convert.ToDouble(drv["IssuedPercent"]), 2);
                    double rewardCredit = (double) drv["RewardIssueCredit"];
                    rewardCredit = Math.Floor(rewardCredit);
                    dr["獎勵額度"] = rewardCredit;
                    dr["交易員"] = drv["TraderID"].ToString();

                    dt.Rows.Add(dr);
                }*/
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateData() {
            try {
                MSSQL.ExecSqlCmd("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);
                /*SqlCommand cmd = new SqlCommand("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                conn.Close();*/

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
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "02" + i.ToString("0#");
                    string warrantID = r.Cells["WarrantID"].Value.ToString();
                    double reIssueNum = Convert.ToDouble(r.Cells["ReIssueNum"].Value);
                    string marketTmr = r.Cells["MarketTmr"].Value == DBNull.Value ? "Y" : r.Cells["MarketTmr"].Value.ToString();
                    string traderID = r.Cells["TraderID"].Value == DBNull.Value ? userID : r.Cells["TraderID"].Value.ToString();
                    bool confirmed = false;
                    confirmed = r.Cells["ConfirmChecked"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["ConfirmChecked"].Value);
                    string confirmChecked = "N";
                    if (confirmed) {
                        confirmChecked = "Y";
                        applyCount++;
                    } else
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
                GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔增額");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void OfficiallyApply() {
            try {
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

                /*SqlCommand cmd1 = new SqlCommand(sql1, conn);
                SqlCommand cmd2 = new SqlCommand(sql2, conn);
                SqlCommand cmd3 = new SqlCommand(sql3, conn);
                SqlCommand cmd4 = new SqlCommand(sql4, conn);*/

                conn.Open();
                MSSQL.ExecSqlCmd(sql1, conn);
                MSSQL.ExecSqlCmd(sql2, conn);
                MSSQL.ExecSqlCmd(sql3, conn);
                MSSQL.ExecSqlCmd(sql4, conn);

                /*cmd1.ExecuteNonQuery();
                cmd1.Dispose();
                cmd2.ExecuteNonQuery();
                cmd2.Dispose();
                cmd3.ExecuteNonQuery();
                cmd3.Dispose();
                cmd4.ExecuteNonQuery();
                cmd4.Dispose();*/
                conn.Close();

                toolStripLabel2.Text = DateTime.Now + "申請成功";
                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 申請" + applyCount + "檔權證增額");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton() {
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            if (isEdit) {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                band0.Columns["WarrantID"].CellActivation = Activation.AllowEdit;
                band0.Columns["ReIssueNum"].CellActivation = Activation.AllowEdit;
                band0.Columns["MarketTmr"].CellActivation = Activation.AllowEdit;
                band0.Columns["isReward"].CellActivation = Activation.AllowEdit;
                band0.Columns["MPrice"].CellActivation = Activation.AllowEdit;
                band0.Columns["WarrantName"].CellActivation = Activation.AllowEdit;
                band0.Columns["exeRatio"].CellActivation = Activation.AllowEdit;

                buttonEdit.Visible = false;
                buttonConfirm.Visible = true;
                buttonDelete.Visible = true;
                buttonCancel.Visible = true;
                toolStripButton1.Visible = false;
                toolStripSeparator2.Visible = false;

                band0.Columns["ConfirmChecked"].Hidden = true;
                band0.Columns["UnderlyingID"].Hidden = true;
                band0.Columns["EquivalentNum"].Hidden = true;
                band0.Columns["IssuedPercent"].Hidden = true;
                band0.Columns["RewardIssueCredit"].Hidden = true;

            } else {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                band0.Columns["ConfirmChecked"].CellActivation = Activation.AllowEdit;

                band0.Columns["WarrantID"].CellActivation = Activation.NoEdit;
                band0.Columns["ReIssueNum"].CellActivation = Activation.NoEdit;
                band0.Columns["MarketTmr"].CellActivation = Activation.NoEdit;
                band0.Columns["isReward"].CellActivation = Activation.NoEdit;
                band0.Columns["MPrice"].CellActivation = Activation.NoEdit;
                band0.Columns["UnderlyingID"].CellActivation = Activation.NoEdit;
                band0.Columns["WarrantName"].CellActivation = Activation.NoEdit;
                band0.Columns["exeRatio"].CellActivation = Activation.NoEdit;
                band0.Columns["EquivalentNum"].CellActivation = Activation.NoEdit;
                band0.Columns["IssuedPercent"].CellActivation = Activation.NoEdit;
                band0.Columns["RewardIssueCredit"].CellActivation = Activation.NoEdit;

                buttonEdit.Visible = true;
                buttonConfirm.Visible = false;
                buttonDelete.Visible = false;
                buttonCancel.Visible = false;
                toolStripButton1.Visible = true;
                toolStripSeparator2.Visible = true;

                band0.Columns["ConfirmChecked"].Hidden = false;
                band0.Columns["UnderlyingID"].Hidden = false;
                band0.Columns["EquivalentNum"].Hidden = false;
                band0.Columns["IssuedPercent"].Hidden = false;
                band0.Columns["RewardIssueCredit"].Hidden = false;
            }
        }

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;

            ValueList v;
            if (!e.Layout.ValueLists.Exists("MyValueList")) {
                v = e.Layout.ValueLists.Add("MyValueList");
                v.ValueListItems.Add("Y", "Y");
                v.ValueListItems.Add("N", "N");
            }
            e.Layout.Bands[0].Columns["MarketTmr"].ValueList = e.Layout.ValueLists["MyValueList"];
        }

        private void buttonEdit_Click(object sender, EventArgs e) {
            isEdit = true;
            SetButton();
        }

        private void buttonConfirm_Click(object sender, EventArgs e) {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void buttonDelete_Click(object sender, EventArgs e) {
            isEdit = true;

            DialogResult result = MessageBox.Show("將全部刪除，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                MSSQL.ExecSqlCmd("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);
                /*SqlCommand cmd = new SqlCommand("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                conn.Close();*/
            }
            LoadData();
            SetButton();
        }

        private void buttonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void toolStripButton1_Click(object sender, EventArgs e) {
            if (GlobalVar.globalParameter.userGroup == "FE") {
                OfficiallyApply();
                LoadData();
            } else {
                if (DateTime.Now.TimeOfDay.TotalMinutes > 630)
                    MessageBox.Show("超過交易所申報時間，欲改條件請洽行政組");
                else if (DateTime.Now.TimeOfDay.TotalMinutes > 570) {
                    DialogResult result = MessageBox.Show("超過約定的9:30了，已經告知組長及行政?", "逾時申請", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes) {
                        OfficiallyApply();
                        LoadData();
                        GlobalUtility.LogInfo("Announce", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證增額");

                    } else
                        LoadData();
                } else {
                    OfficiallyApply();
                    LoadData();
                }
            }
        }

        private void ultraGrid1_CellChange(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "ConfirmChecked")
                ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "WarrantID") {
                string warrantID = e.Cell.Row.Cells["WarrantID"].Value.ToString();
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
                //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                DataTable dtTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRow drTemp in dtTemp.Rows) {
                    useReward = drTemp["isReward"].ToString();
                    warrantPrice = Convert.ToDouble(drTemp["MPrice"]);
                    warrantName = drTemp["WarrantName"].ToString();
                    traderID = drTemp["TraderID"].ToString();
                    cr = Convert.ToDouble(drTemp["exeRatio"]);
                }
                e.Cell.Row.Cells["isReward"].Value = useReward;
                e.Cell.Row.Cells["MPrice"].Value = warrantPrice;
                e.Cell.Row.Cells["WarrantName"].Value = warrantName;
                e.Cell.Row.Cells["exeRatio"].Value = cr;
                e.Cell.Row.Cells["TraderID"].Value = traderID;
            }
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right) {
                contextMenuStrip1.Show();
            }
        }

        private void ultraGrid1_BeforeRowsDeleted(object sender, BeforeRowsDeletedEventArgs e) {
            e.DisplayPromptMsg = false;
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e) {
            DialogResult result = MessageBox.Show("刪除此檔，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                ultraGrid1.ActiveRow.Delete();
                UpdateData();
            }
            LoadData();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            string warrantID = "";
            string underlyingID = "";
            warrantID = e.Row.Cells["WarrantID"].Value.ToString();
            underlyingID = e.Row.Cells["UnderlyingID"].Value.ToString();

            string traderID = "NA";
            string issuable = "NA";
            string reissuable = "NA";

            string toolTip1 = "今日未達增額標準";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的發行檢查=N";
            string toolTip4 = "非此使用者所屬標的";

            string sqlTemp = "SELECT TraderID, IsNull(Issuable,'NA') Issuable FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp.Count > 0) {
                foreach (DataRowView drTemp in dvTemp) {
                    traderID = "000" + drTemp["TraderID"].ToString();
                    issuable = drTemp["Issuable"].ToString();
                }
            }
            string sqlTemp2 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [EDIS].[dbo].[WarrantReIssuable] WHERE WarrantID = '" + warrantID + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp2.Count > 0) {
                foreach (DataRowView drTemp2 in dvTemp2) {
                    reissuable = drTemp2["ReIssuable"].ToString();
                }
            }
            if (!isEdit) {

                if (issuable == "NA") {
                    e.Row.ToolTipText = toolTip2;
                    e.Row.Appearance.ForeColor = Color.Red;
                } else if (issuable == "N") {
                    e.Row.Cells["UnderlyingID"].ToolTipText = toolTip3;
                    e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Red;
                }

                if (reissuable == "NA") {
                    e.Row.Cells["WarrantID"].ToolTipText = toolTip1;
                    e.Row.Cells["WarrantID"].Appearance.ForeColor = Color.Red;
                }

                if (issuable != "NA" && traderID != userID) {
                    e.Row.Appearance.BackColor = Color.LightYellow;
                    e.Row.Cells["WarrantID"].ToolTipText = toolTip4;
                }
            }

        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "WarrantID")
                GlobalUtility.MenuItemClick<FrmReIssuable>();
        }

        private void ultraGrid1_DoubleClickHeader(object sender, DoubleClickHeaderEventArgs e) {
            if (e.Header.Column.Key == "ConfirmChecked") {
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    r.Cells["ConfirmChecked"].Value = true;
                    UpdateData();
                    LoadData();
                }
            }
        }

        private void FrmReIssue_FormClosed(object sender, FormClosedEventArgs e) {
            UpdateData();
        }
    }
}
