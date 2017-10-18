﻿using System;
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
        private DataTable dt;// = new DataTable();
        private bool isEdit = false;

        public FrmReIssueInput() {
            InitializeComponent();
        }

        private void FrmReIssueInput_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {

            //dt.PrimaryKey = new DataColumn[] { dt.Columns["權證代號"] };
            //ultraGrid1.DataSource = dt;
            UltraGridBand bands0 = ultraGrid1.DisplayLayout.Bands[0];
            bands0.Columns["IssueNum"].Format = "N0";
            bands0.Columns["SoldNum"].Format = "N0";

            bands0.Columns["WarrantID"].Width = 90;
            bands0.Columns["WarrantName"].Width = 135;
            bands0.Columns["IssueNum"].Width = 90;
            bands0.Columns["SoldNum"].Width = 90;
            bands0.Columns["Last1Sold"].Width = 70;
            bands0.Columns["Last2Sold"].Width = 70;
            bands0.Columns["Last3Sold"].Width = 70;
            bands0.Columns["LastTradingDate"].Width = 90;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["ReIssuable"].Width = 90;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["維護時間"].Width = 120;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            bands0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;

            SetButton();
        }

        private void SetButton() {
            if (isEdit) {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;                               
            } else {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;                
            }
            toolStripButtonEdit.Visible = !isEdit;
            toolStripButtonConfirm.Visible = isEdit;
            toolStripButtonCancel.Visible = isEdit;
            Edit2.Visible = isEdit;
        }

        private void LoadData() {
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
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            ultraGrid1.DataSource = dt;

            dt.Columns[0].Caption = "權證代號";
            dt.Columns[1].Caption = "權證名稱";
            dt.Columns[2].Caption = "發行張數";
            dt.Columns[3].Caption = "流通在外";
            dt.Columns[4].Caption = "前1日";
            dt.Columns[5].Caption = "前2日";
            dt.Columns[6].Caption = "前3日";
            dt.Columns[7].Caption = "最後交易日";
            dt.Columns[8].Caption = "符合增額條件";

            foreach (DataRow row in dt.Rows) {
                row["IssueNum"] = (double) row["IssueNum"] / 1000;
                row["SoldNum"] = (double) row["SoldNum"] / 1000;
            }
        }

        private void UpdateDB() {
            for (int x = ultraGrid1.Rows.Count - 1; x >= 0; x--) {
                try {
                    if (ultraGrid1.Rows[x].Cells[1].Value.ToString() == "")
                        ultraGrid1.Rows[x].Delete(false);

                } catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            }

            EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [WarrantReIssuable]", conn);

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
                    string warrantID = r.Cells["WarrantID"].Value.ToString();
                    string warrantName = r.Cells["WarrantName"].Value.ToString();
                    double issueNum = Convert.ToDouble(r.Cells["IssueNum"].Value);
                    double soldNum = Convert.ToDouble(r.Cells["SoldNum"].Value);
                    double last1Sold = Convert.ToDouble(r.Cells["Last1Sold"].Value);
                    double last2Sold = Convert.ToDouble(r.Cells["Last2Sold"].Value);
                    double last3Sold = Convert.ToDouble(r.Cells["Last3Sold"].Value);
                    string lastTradingDate = r.Cells["LastTradingDate"].Value.ToString();
                    string reIssuable = r.Cells["ReIssuable"].Value.ToString();


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
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e) {
            toolStripLabel1.Text = "";
            dt.Rows.Clear();
            isEdit = true;
            SetButton();

            //Get key and id
            DataTable dv = EDLib.SQL.MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='10'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string key = dv.Rows[0]["FLGDAT_FLGDTA"].ToString();

            dv = EDLib.SQL.MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='20'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string id = dv.Rows[0]["FLGDAT_FLGDTA"].ToString();

            DateTime lastTrade = TradeDate.LastNTradeDateDT(1);
            string aday = (lastTrade.Year - 1911) + lastTrade.ToString("MMdd");
            string twseUrl = "http://siis.twse.com.tw/server-java/t159sa04?step=1&id=9200pd" + id + "&TYPEK=sii&key=" + key + "&cDATE=" + aday + "&co_id=9200";

            //dt.Rows.Clear();

            //parse TWSE Incr html
            if (!ParseHtml(twseUrl))
                return;

            //parse OTC Incr html
            twseUrl = "http://siis.twse.com.tw/server-java/t159sa04?step=1&id=9200pd" + id + "&TYPEK=otc&key=" + key + "&cDATE=" + aday + "&co_id=9200";
            if (!ParseHtml(twseUrl))
                return;

            //LoadData();    
            GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 下載可增額列表");
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

        private bool ParseHtml(string url) {
            try {
                string firstResponse = EDLib.Utility.GetHtml(url, System.Text.Encoding.Default);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(firstResponse);
                HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]").ChildNodes;
                // /html[1]/body[1]/center[1]/table

                for (int i = 5; i < navNodeChild.Count; i += 2) {

                    string[] split = navNodeChild[i].InnerText.Split(new string[] { " ", "\t", "&nbsp;", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    DataRow dr = dt.NewRow();

                    dr["WarrantID"] = split[0];
                    dr["WarrantName"] = split[1];
                    dr["IssueNum"] = split[2];
                    dr["SoldNum"] = split[3];
                    dr["Last1Sold"] = split[4];
                    dr["Last2Sold"] = split[5];
                    dr["Last3Sold"] = split[6];
                    dr["LastTradingDate"] = split[7];
                    dr["ReIssuable"] = split[8];

                    dt.Rows.Add(dr);
                }
                return true;
            } catch (Exception e) {
                MessageBox.Show("可能要更新Key或是還沒有資料");
                return false;
            }
        }

        private void Edit2_Click(object sender, EventArgs e) {
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
            dt.Rows.Clear();
            for (int x = 0; x < 30; x++) 
                ultraGrid1.DisplayLayout.Bands[0].AddNew();                
            
            ultraGrid1.ActiveRowScrollRegion.ScrollRowIntoView(ultraGrid1.Rows[0]);
        }
    }
}
