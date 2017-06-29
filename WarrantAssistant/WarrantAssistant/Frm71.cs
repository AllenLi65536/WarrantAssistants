using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.IO;
using System.Net;
using HtmlAgilityPack;

namespace WarrantAssistant
{
    public partial class Frm71:Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt = new DataTable();
        private bool isEdit = false;

        public Frm71() {
            InitializeComponent();
        }

        private void Frm71_Load(object sender, EventArgs e) {
            toolStripLabel1.Text = "";
            InitialGrid();
            LoadData();
        }

        private void InitialGrid() {
            dt.Columns.Add("發行人", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("發行張數", typeof(double));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("申報時間", typeof(string));
            dt.Columns.Add("可發行股數", typeof(double));
            dt.Columns.Add("截至前一日", typeof(double));
            dt.Columns.Add("本日累積發行", typeof(double));
            dt.Columns.Add("累計%", typeof(string));
            dt.Columns.Add("同標的2檔", typeof(string));
            dt.Columns.Add("原始申報時間", typeof(string));

            //dt.PrimaryKey = new DataColumn[] { dt.Columns["權證名稱"] };
            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["發行張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["可發行股數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["截至前一日"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["本日累積發行"].Format = "###,###";

            ultraGrid1.DisplayLayout.Bands[0].Columns["發行人"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].Width = 130;
            ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["發行張數"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["申報時間"].Width = 110;
            ultraGrid1.DisplayLayout.Bands[0].Columns["可發行股數"].Width = 120;
            ultraGrid1.DisplayLayout.Bands[0].Columns["截至前一日"].Width = 120;
            ultraGrid1.DisplayLayout.Bands[0].Columns["本日累積發行"].Width = 120;
            ultraGrid1.DisplayLayout.Bands[0].Columns["累計%"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["同標的2檔"].Width = 80;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["原始申報時間"].Width = 110;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            ultraGrid1.DisplayLayout.Bands[0].Columns["可發行股數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["截至前一日"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["本日累積發行"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["累計%"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;

            SetButton();
        }

        private void SetButton() {
            if (isEdit) {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
                toolStripButtonEdit.Visible = false;
                toolStripButtonConfirm.Visible = true;
                toolStripButtonCancel.Visible = true;
                toolStripButtonGetData.Visible = true;
                                
                for (int x = 0; x < 100; x++) {
                    ultraGrid1.DisplayLayout.Bands[0].AddNew();
                    //ultraGrid1.Rows[x].Cells[1].Value = (x + 1).ToString();
                    //ultraGrid1.Rows[x].Cells[9].Value = DateTime.Now;
                }

                ultraGrid1.ActiveRowScrollRegion.ScrollRowIntoView(ultraGrid1.Rows[0]);
                //ultraGrid1.Rows[0].Selected = true;
            } else {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
                toolStripButtonEdit.Visible = true;
                toolStripButtonConfirm.Visible = false;
                toolStripButtonCancel.Visible = false;
                toolStripButtonGetData.Visible = false;
            }
        }

        private void LoadData() {
            dt.Rows.Clear();
            string sql = @"SELECT [Issuer]
                                  ,[WarrantName]
                                  ,[UnderlyingID]
                                  ,[IssueNum]
                                  ,[exeRatio]
                                  ,[ApplyTime]
                                  ,[AvailableShares]
                                  ,[LastDayUsedShares]
                                  ,[TodayApplyShares]
                                  ,[AccUsedShares]
                                  ,[SameUnderlying]
                                  ,[OriApplyTime]
                              FROM [EDIS].[dbo].[Apply_71]";
            DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

            foreach (DataRowView drv in dv) {
                DataRow dr = dt.NewRow();

                dr["發行人"] = drv["Issuer"].ToString();
                dr["權證名稱"] = drv["WarrantName"].ToString();
                dr["標的代號"] = drv["UnderlyingID"].ToString();
                dr["發行張數"] = Convert.ToDouble(drv["IssueNum"]);
                dr["行使比例"] = Convert.ToDouble(drv["exeRatio"]);
                dr["申報時間"] = drv["ApplyTime"].ToString();
                dr["可發行股數"] = Convert.ToDouble(drv["AvailableShares"]);
                dr["截至前一日"] = Convert.ToDouble(drv["LastDayUsedShares"]);
                dr["本日累積發行"] = Convert.ToDouble(drv["TodayApplyShares"]);
                dr["累計%"] = drv["AccUsedShares"].ToString();
                dr["同標的2檔"] = drv["SameUnderlying"].ToString();
                dr["原始申報時間"] = drv["OriApplyTime"].ToString();

                dt.Rows.Add(dr);
            }
        }

        private void UpdateDB() {
            for (int x = ultraGrid1.Rows.Count - 1; x >= 0; x--) {
                try {
                    if (ultraGrid1.Rows[x].Cells[0].Value.ToString() == "") {
                        ultraGrid1.Rows[x].Delete(false);
                    }
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            }

            SqlCommand cmd = new SqlCommand("DELETE FROM [Apply_71]", conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            conn.Close();

            try {
                string sql = "INSERT INTO [Apply_71] values(@Issuer,@WarrantName,@UnderlyingID,@IssueNum,@exeRatio,@ApplyTime,@AvailableShares,@LastDayUsedShares,@TodayApplyShares,@AccUsedShares,@SameUnderlying,@OriApplyTime,@Result, @ApplyStatus, @ReIssueResult, @SerialNum)";
                List<SqlParameter> ps = new List<SqlParameter>();
                ps.Add(new SqlParameter("@Issuer", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@exeRatio", SqlDbType.Float));
                ps.Add(new SqlParameter("@ApplyTime", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@AvailableShares", SqlDbType.Float));
                ps.Add(new SqlParameter("@LastDayUsedShares", SqlDbType.Float));
                ps.Add(new SqlParameter("@TodayApplyShares", SqlDbType.Float));
                ps.Add(new SqlParameter("@AccUsedShares", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@SameUnderlying", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@OriApplyTime", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@Result", SqlDbType.Float));
                ps.Add(new SqlParameter("@ApplyStatus", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ReIssueResult", SqlDbType.Float));
                ps.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sql, ps);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string issuer = r.Cells["發行人"].Value.ToString();
                    string warrantName = r.Cells["權證名稱"].Value.ToString();
                    string underlyingID = r.Cells["標的代號"].Value.ToString();
                    double issueNum = Convert.ToDouble(r.Cells["發行張數"].Value);
                    double exeRatio = Convert.ToDouble(r.Cells["行使比例"].Value);
                    string applyTime = r.Cells["申報時間"].Value.ToString();
                    double availableShares = 0.0;
                    double lastDayUsedShares = 0.0;
                    double todayApplyShares = 0.0;
                    if (underlyingID == "IX0001") {
                        availableShares = 0.0;
                        lastDayUsedShares = 0.0;
                        todayApplyShares = 0.0;
                    } else {
                        availableShares = Convert.ToDouble(r.Cells["可發行股數"].Value);
                        lastDayUsedShares = Convert.ToDouble(r.Cells["截至前一日"].Value);
                        todayApplyShares = Convert.ToDouble(r.Cells["本日累積發行"].Value);
                    }
                    string accUsedShares = r.Cells["累計%"].Value.ToString();
                    string sameUnderlying = r.Cells["同標的2檔"].Value.ToString();
                    string oriApplyTime = r.Cells["原始申報時間"].Value.ToString();

                    //string underlyingName = warrantName.Substring(1, warrantName.Length - 7);//需考慮以前的短權證名稱

                    double multiplier = 1.0;
                    string sqlTemp = "SELECT CASE WHEN [StockType]='DS' OR [StockType]='DR' THEN 0.22 ELSE 1 END AS Multiplier FROM [EDIS].[dbo].[WarrantUnderlying] WHERE UnderlyingID = '" + underlyingID + "'";
                    DataView dv = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                    foreach (DataRowView dr in dv) {
                        multiplier = Convert.ToDouble(dr["Multiplier"]);
                    }

                    double todayAvailable = Math.Round(((availableShares * multiplier - lastDayUsedShares) / 1000), 1);
                    double attempShares = issueNum * exeRatio;
                    double result = 0.0;
                    double tempAvailable = 0.0;
                    string applyStatus = "";
                    tempAvailable = todayAvailable - todayApplyShares / 1000 + attempShares;

                    if (underlyingID == "IX0001") {
                        result = attempShares;
                        applyStatus = "Y";
                    } else if (applyTime.Substring(0, 2) == "09") {
                        if (tempAvailable >= attempShares) {
                            result = attempShares;
                            applyStatus = "Y";
                        } else if (tempAvailable > 0) {
                            result = tempAvailable;
                            applyStatus = "排隊中";
                        } else {
                            result = 0;
                            if (todayAvailable >= 0.6 * attempShares)
                                applyStatus = "排隊中";
                            else
                                applyStatus = "X 沒額度";
                        }
                    } else if (applyTime.Substring(0, 2) == "22") {
                        result = 0;
                        applyStatus = "X 沒額度";
                    } else if (applyTime.Substring(0, 2) == "10") {
                        if (tempAvailable >= attempShares) {
                            result = attempShares;
                            applyStatus = "Y";
                        } else if (tempAvailable > 0) {
                            result = tempAvailable;
                            applyStatus = "排隊中";
                        } else {
                            result = 0;
                            if (todayAvailable >= 0.6 * attempShares)
                                applyStatus = "排隊中";
                            else
                                applyStatus = "X 沒額度";
                        }
                    }

                    double accUsed = 0.0;
                    accUsed = (lastDayUsedShares + todayApplyShares) / availableShares;
                    double reIssueResult = 0.0;
                    if (accUsed <= 0.3)
                        reIssueResult = attempShares;
                    else
                        reIssueResult = 0.0;

                    h.SetParameterValue("@Issuer", issuer);
                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@UnderlyingID", underlyingID);
                    h.SetParameterValue("@IssueNum", issueNum);
                    h.SetParameterValue("@exeRatio", exeRatio);
                    h.SetParameterValue("@ApplyTime", applyTime);
                    h.SetParameterValue("@AvailableShares", availableShares);
                    h.SetParameterValue("@LastDayUsedShares", lastDayUsedShares);
                    h.SetParameterValue("@TodayApplyShares", todayApplyShares);
                    h.SetParameterValue("@AccUsedShares", accUsedShares);
                    h.SetParameterValue("@SameUnderlying", sameUnderlying);
                    h.SetParameterValue("@OriApplyTime", oriApplyTime);
                    h.SetParameterValue("@Result", result);
                    h.SetParameterValue("@ApplyStatus", applyStatus);
                    h.SetParameterValue("@ReIssueResult", reIssueResult);
                    h.SetParameterValue("@SerialNum", "0");

                    h.ExecuteCommand();
                }

                h.Dispose();

                string sql5 = "UPDATE [EDIS].[dbo].[Apply_71] SET SerialNum = B.SerialNum FROM [EDIS].[dbo].[ApplyTotalList] B WHERE [Apply_71].[WarrantName]=B.WarrantName";
                string sql2 = "UPDATE [EDIS].[dbo].[ApplyTotalList] SET Result=0";
                string sql3 = @"UPDATE [EDIS].[dbo].[ApplyTotalList] 
                                SET Result= CASE WHEN ApplyKind='1' THEN B.Result ELSE B.ReIssueResult END
                                FROM [EDIS].[dbo].[Apply_71] B
                                WHERE [ApplyTotalList].[WarrantName]=B.WarrantName";
                string sql4 = @"UPDATE [EDIS].[dbo].[ApplyTotalList]
                                SET Result= CASE WHEN [RewardCredit]>=[EquivalentNum] THEN [EquivalentNum] ELSE [RewardCredit] END
                                WHERE [UseReward]='Y'";

                SqlCommand cmd5 = new SqlCommand(sql5, conn);
                SqlCommand cmd2 = new SqlCommand(sql2, conn);
                SqlCommand cmd3 = new SqlCommand(sql3, conn);
                SqlCommand cmd4 = new SqlCommand(sql4, conn);

                conn.Open();
                cmd5.ExecuteNonQuery();
                cmd5.Dispose();
                cmd2.ExecuteNonQuery();
                cmd2.Dispose();
                cmd3.ExecuteNonQuery();
                cmd3.Dispose();
                cmd4.ExecuteNonQuery();
                cmd4.Dispose();
                conn.Close();

                toolStripLabel1.Text = DateTime.Now + "更新成功";

                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 更新7-1試算表");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));
                
                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Info");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID+" 更新7-1試算表");
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/


            } catch (Exception ex) {
                GlobalUtility.LogInfo("Exception", GlobalVar.globalParameter.userID + "7-1試算表" + ex.Message);
                /*string sqlInfo = "INSERT INTO [InformationLog] values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Exception");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + "7-1試算表"+ex.Message);
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/

                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e) {
            toolStripLabel1.Text = "";
            dt.Rows.Clear();
            isEdit = true;
            SetButton();
        }

        private void toolStripButtonConfirm_Click(object sender, EventArgs e) {
            isEdit = false;
            UpdateDB();
            LoadData();
            SetButton();
        }

        private void toolStripButtonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e) {
            string applyTime = e.Row.Cells["申報時間"].Value == DBNull.Value ? "" : e.Row.Cells["申報時間"].Value.ToString();
            double issueNum = e.Row.Cells["發行張數"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["發行張數"].Value);
            string underlyingID = e.Row.Cells["標的代號"].Value.ToString();


            if (applyTime.Length > 0) {
                applyTime = applyTime.Substring(0, 2);

                if (applyTime == "22")
                    e.Row.Cells["申報時間"].Appearance.ForeColor = Color.Red;

                if (applyTime == "09")
                    e.Row.Cells["申報時間"].Appearance.ForeColor = Color.Green;

                if (applyTime == "10" && issueNum != 10000)
                    e.Row.Cells["發行張數"].Appearance.ForeColor = Color.Red;
            }


        }

        private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e) {
            if (isEdit)
                e.Cancel = true;
        }

        private void toolStripLabel1_Click(object sender, EventArgs e) {

        }

        private void toolStripButtonGetData_Click(object sender, EventArgs e) {
            //Get key and id
            DataView dv = DeriLib.Util.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='10'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string key = dv[0]["FLGDAT_FLGDTA"].ToString();

            dv = DeriLib.Util.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='20'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string id = dv[0]["FLGDAT_FLGDTA"].ToString();

            DateTime lastTrade = TradeDate.LastNTradeDateDT(1);
            string aday = (lastTrade.Year - 1911) + lastTrade.ToString("MMdd");
            string twseUrl = "http://siis.twse.com.tw/server-java/t150sa10?step=0&id=9200pd" + id + "&TYPEK=sii&key=" + key;

            dt.Rows.Clear();

            //parse TWSE 7-1 html
            parsehtml(twseUrl);

            //parse OTC 7-1 html
            twseUrl = "http://siis.twse.com.tw/server-java/o_t150sa10?step=0&id=9200pd"+id+"&TYPEK=otc&key="+key;
            parsehtml(twseUrl);

            GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 下載7-1試算表");
        }

        private void parsehtml(string url) {

            string FirstResponse = GlobalUtility.GetHtml(url);

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(FirstResponse);            
            HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]").ChildNodes; // /td[1]/table[1]/tr[2]
     
            // MessageBox.Show(navNodeChild.Count.ToString());
            // for (int i = 5; i < navNodeChild.Count; i +=2)
            //   MessageBox.Show(navNodeChild[i].InnerText);

            int loopend = navNodeChild.Count;
            /*if (twse)
                loopend = navNodeChild.Count - 4;
            else
                loopend = navNodeChild.Count - 8;*/
            
            for (int i = 5; i < loopend; i += 2) {
                //MessageBox.Show(navNodeChild[i].InnerText);

                string[] split = navNodeChild[i].InnerText.Split(new string[] {"\n"}, StringSplitOptions.RemoveEmptyEntries); //" ", "\t", "&nbsp;",
                if (split.Length != 12)
                    continue;
                DataRow dr = dt.NewRow();

                dr["發行人"] = split[0];
                dr["權證名稱"] = split[1];
                dr["標的代號"] = split[2];
                dr["發行張數"] = split[3];
                dr["行使比例"] = split[4];
                dr["申報時間"] = split[5];
                dr["可發行股數"] = split[6];
                dr["截至前一日"] = split[7];
                dr["本日累積發行"] = split[8];
                dr["累計%"] = split[9];

                if (!split[10].StartsWith("&nbsp"))
                    dr["同標的2檔"] = split[10];
                if (!split[11].StartsWith("&nbsp"))
                    dr["原始申報時間"] = split[11];

                dt.Rows.Add(dr);
            }
        }

        /*
private void ultraGrid1_CellDataError(object sender, Infragistics.Win.UltraWinGrid.CellDataErrorEventArgs e)
{
if (isEdit)
{
e.RaiseErrorEvent = false;
e.StayInEditMode = false;
}
}
*/

    }
}
