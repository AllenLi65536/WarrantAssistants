using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using System.Net;
using System.IO;
using System.Threading;
using HtmlAgilityPack;

namespace WarrantAssistant
{
    public partial class FrmReIssueInput:Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt = new DataTable();
        private bool isEdit = false;

        public FrmReIssueInput() {
            InitializeComponent();
        }

        private void FrmReIssueInput_Load(object sender, EventArgs e) {
            InitialGrid();
            LoadData();
        }

        private void InitialGrid() {
            dt.Columns.Add("權證代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("發行張數", typeof(double));
            dt.Columns.Add("流通在外", typeof(double));
            dt.Columns.Add("前1日", typeof(double));
            dt.Columns.Add("前2日", typeof(double));
            dt.Columns.Add("前3日", typeof(double));
            dt.Columns.Add("最後交易日", typeof(string));
            dt.Columns.Add("符合增額條件", typeof(string));
            //dt.Columns.Add("維護時間", typeof(DateTime));

            //dt.PrimaryKey = new DataColumn[] { dt.Columns["權證代號"] };
            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["發行張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["流通在外"].Format = "###,###";

            ultraGrid1.DisplayLayout.Bands[0].Columns["權證代號"].Width = 90;
            ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].Width = 135;
            ultraGrid1.DisplayLayout.Bands[0].Columns["發行張數"].Width = 90;
            ultraGrid1.DisplayLayout.Bands[0].Columns["流通在外"].Width = 90;
            ultraGrid1.DisplayLayout.Bands[0].Columns["前1日"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["前2日"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["前3日"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["最後交易日"].Width = 90;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["符合增額條件"].Width = 90;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["維護時間"].Width = 120;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

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

                for (int x = 0; x < 30; x++) {
                    ultraGrid1.DisplayLayout.Bands[0].AddNew();
                    //ultraGrid1.Rows[x].Cells[0].Value = (x+1).ToString();
                    //ultraGrid1.Rows[x].Cells[9].Value = DateTime.Now;
                }
                ultraGrid1.ActiveRowScrollRegion.ScrollRowIntoView(ultraGrid1.Rows[0]);
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
            string sql = @"SELECT [WarrantID]
                                  ,[WarrantName]
                                  ,[IssueNum]
                                  ,[SoldNum]
                                  ,[Last1Sold]
                                  ,[Last2Sold]
                                  ,[Last3Sold]
                                  ,[LastTradingDate]
                                  ,[ReIssuable]
                              FROM [EDIS].[dbo].[WarrantReIssuable]";
            DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

            foreach (DataRowView drv in dv) {
                DataRow dr = dt.NewRow();

                dr["權證代號"] = drv["WarrantID"].ToString();
                dr["權證名稱"] = drv["WarrantName"].ToString();
                dr["發行張數"] = Convert.ToDouble(drv["IssueNum"]) / 1000;
                dr["流通在外"] = Convert.ToDouble(drv["SoldNum"]) / 1000;
                dr["前1日"] = Convert.ToDouble(drv["Last1Sold"]);
                dr["前2日"] = Convert.ToDouble(drv["Last2Sold"]);
                dr["前3日"] = Convert.ToDouble(drv["Last3Sold"]);
                dr["最後交易日"] = drv["LastTradingDate"].ToString();
                dr["符合增額條件"] = drv["ReIssuable"].ToString();
                //dr["維護時間"] = Convert.ToDateTime(drv["MDate"]);

                dt.Rows.Add(dr);
            }
        }

        private void UpdateDB() {
            for (int x = ultraGrid1.Rows.Count - 1; x >= 0; x--) {
                try {
                    if (ultraGrid1.Rows[x].Cells[1].Value.ToString() == "") {
                        ultraGrid1.Rows[x].Delete(false);
                    }
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            }

            SqlCommand cmd = new SqlCommand("DELETE FROM [WarrantReIssuable]", conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            conn.Close();

            try {
                string sql = "INSERT INTO [WarrantReIssuable] values(@WarrantID,@WarrantName,@IssueNum,@SoldNum,@Last1Sold,@Last2Sold,@Last3Sold,@LastTradingDate,@ReIssuable,@MDate)";
                List<SqlParameter> ps = new List<SqlParameter>();
                ps.Add(new SqlParameter("@WarrantID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@SoldNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@Last1Sold", SqlDbType.Float));
                ps.Add(new SqlParameter("@Last2Sold", SqlDbType.Float));
                ps.Add(new SqlParameter("@Last3Sold", SqlDbType.Float));
                ps.Add(new SqlParameter("@LastTradingDate", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ReIssuable", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@MDate", SqlDbType.DateTime));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sql, ps);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string warrantID = r.Cells["權證代號"].Value.ToString();
                    string warrantName = r.Cells["權證名稱"].Value.ToString();
                    double issueNum = Convert.ToDouble(r.Cells["發行張數"].Value);
                    double soldNum = Convert.ToDouble(r.Cells["流通在外"].Value);
                    double last1Sold = Convert.ToDouble(r.Cells["前1日"].Value);
                    double last2Sold = Convert.ToDouble(r.Cells["前2日"].Value);
                    double last3Sold = Convert.ToDouble(r.Cells["前3日"].Value);
                    string lastTradingDate = r.Cells["最後交易日"].Value.ToString();
                    string reIssuable = r.Cells["符合增額條件"].Value.ToString();


                    h.SetParameterValue("@WarrantID", warrantID);
                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@IssueNum", issueNum);
                    h.SetParameterValue("@SoldNum", soldNum);
                    h.SetParameterValue("@Last1Sold", last1Sold);
                    h.SetParameterValue("@Last2Sold", last2Sold);
                    h.SetParameterValue("@Last3Sold", last3Sold);
                    h.SetParameterValue("@LastTradingDate", lastTradingDate);
                    h.SetParameterValue("@ReIssuable", reIssuable);
                    h.SetParameterValue("@MDate", DateTime.Now);

                    h.ExecuteCommand();
                }

                h.Dispose();
                toolStripLabel1.Text = DateTime.Now + "更新成功";

                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 更新可增額列表");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Info");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 更新可增額列表");
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/
            } catch (Exception ex) {
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
            string twseUrl = "http://siis.twse.com.tw/server-java/t159sa04?step=1&id=9200pd" + id + "&TYPEK=sii&key=" + key + "&cDATE=" + aday + "&co_id=9200";
            
            dt.Rows.Clear();
            
            //parse TWSE Incr html
            parsehtml(twseUrl);

            //parse OTC Incr html
            twseUrl = "http://siis.twse.com.tw/server-java/t159sa04?step=1&id=9200pd" + id + "&TYPEK=otc&key=" + key + "&cDATE=" + aday + "&co_id=9200";
            parsehtml(twseUrl);
            //LoadData();    
            GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 下載可增額列表");

        }
        private void parsehtml(string url) {

            string FirstResponse = GlobalUtility.GetHtml(url);

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(FirstResponse);
            HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]").ChildNodes;
            // /html[1]/body[1]/center[1]/table

            for (int i = 5; i < navNodeChild.Count; i += 2) {
                //MessageBox.Show(navNodeChild[i].InnerText);
                                
                string[] split = navNodeChild[i].InnerText.Split(new string[] { " ", "\t", "&nbsp;", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                DataRow dr = dt.NewRow();

                dr["權證代號"] = split[0];
                dr["權證名稱"] = split[1];
                dr["發行張數"] = split[2];
                dr["流通在外"] = split[3];
                dr["前1日"] = split[4];
                dr["前2日"] = split[5];
                dr["前3日"] = split[6];
                dr["最後交易日"] = split[7];
                dr["符合增額條件"] = split[8];

                dt.Rows.Add(dr);
            }
        }
    }
}
