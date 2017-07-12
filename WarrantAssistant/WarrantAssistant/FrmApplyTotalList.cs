using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;

namespace WarrantAssistant
{
    public partial class FrmApplyTotalList:Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt;// = new DataTable();
        //private string userID = GlobalVar.globalParameter.userID;
        private bool isEdit = false;

        public FrmApplyTotalList() {
            InitializeComponent();
        }

        private void FrmApplyTotalList_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {
            
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            band0.Columns["IssueNum"].Format = "###,###";
            band0.Columns["EquivalentNum"].Format = "###,###";
            band0.Columns["Result"].Format = "###,###";
            band0.Columns["Credit"].Format = "###,###";
            band0.Columns["RewardCredit"].Format = "###,###";

            band0.Columns["IssueNum"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["EquivalentNum"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["Result"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["Credit"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["RewardCredit"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["UseReward"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            ultraGrid1.DisplayLayout.Bands[0].Columns["SerialNum"].Hidden = true;

            SetButton();
        }

        private void LoadData() {
            try {
                //dt.Rows.Clear();
                string sql = @"SELECT a.[SerialNum]
                              ,CASE WHEN a.[ApplyKind]='1' THEN '新發' ELSE '增額' END ApplyKind
                              ,a.[TraderID]
                              ,a.[Market]
                              ,a.[UnderlyingID]
                              ,a.[WarrantName]
                              ,a.[CR]
                              ,a.[IssueNum]
                              ,a.[EquivalentNum]
                              ,IsNull(a.[Result], 0) Result
                              ,IsNull(b.[IssueCredit],0) Credit
                              ,IsNull(b.[RewardIssueCredit],0) RewardCredit
                              ,a.[UseReward]                              
                          FROM [EDIS].[dbo].[ApplyTotalList] a 
                          LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                          left join Underlying_TraderIssue c on a.UnderlyingID=c.UID 
                          ORDER BY  a.Market desc, a.ApplyKind, a.SerialNum";// or (a.UnderlyingID = 'IX0001' and c.UID ='TWA00')
                dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                ultraGrid1.DataSource = dt;

                dt.Columns[0].Caption = "序號";
                dt.Columns[1].Caption = "類型";
                dt.Columns[2].Caption = "交易員";
                dt.Columns[3].Caption = "市場";
                dt.Columns[4].Caption = "標的代號";
                dt.Columns[5].Caption = "權證名稱";
                dt.Columns[6].Caption = "行使比例";
                dt.Columns[7].Caption = "張數";
                dt.Columns[8].Caption = "約當張數";
                dt.Columns[9].Caption = "額度結果";
                dt.Columns[10].Caption = "今日額度";
                dt.Columns[11].Caption = "獎勵額度";
                dt.Columns[12].Caption = "使用獎勵";
                foreach (DataRow row in dt.Rows) {
                    row["Result"] = Math.Round((double) row["Result"]);//Math.Floor((double) row["Result"]);
                    row["Credit"] = Math.Round((double) row["Credit"]);
                    row["RewardCredit"] = Math.Round((double) row["RewardCredit"]);
                    row["TraderID"] = row["TraderID"].ToString().TrimStart('0');
                }

                /*DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    dr["序號"] = drv["SerialNum"].ToString();
                    dr["類型"] = drv["ApplyKind"].ToString();
                    dr["交易員"] = drv["TraderID"].ToString();
                    dr["市場"] = drv["Market"].ToString();
                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    dr["權證名稱"] = drv["WarrantName"].ToString();
                    dr["行使比例"] = Convert.ToDouble(drv["CR"]);
                    dr["張數"] = Convert.ToDouble(drv["IssueNum"]);
                    dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                    //double result = drv["Result"] == DBNull.Value ? 0.0 : Convert.ToDouble(drv["Result"]);
                    //result = Math.Floor(drv["Result"] == DBNull.Value ? 0.0 : Convert.ToDouble(drv["Result"]));
                    double credit = (double) drv["Credit"];
                    credit = Math.Floor(credit);
                    double rewardCredit = (double) drv["RewardCredit"];
                    rewardCredit = Math.Floor(rewardCredit);
                    dr["Result"] = Math.Floor(drv["Result"] == DBNull.Value ? 0.0 : Convert.ToDouble(drv["Result"]));
                    dr["今日額度"] = credit;
                    dr["獎勵額度"] = rewardCredit;
                    dr["使用獎勵"] = drv["UseReward"].ToString();
                    // dr["發行原因"] = drv["Reason"] == DBNull.Value ? " " : reasonString[Convert.ToInt32(drv["Reason"])];
                    dt.Rows.Add(dr);
                }*/
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateData() {
            try {
                string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName, CR=@CR, IssueNum=@IssueNum, EquivalentNum=@EquivalentNum, Result=@Result, UseReward=@UseReward WHERE SerialNum=@SerialNum";
                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                pars.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@CR", SqlDbType.Float));
                pars.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                pars.Add(new SqlParameter("@EquivalentNum", SqlDbType.Float));
                pars.Add(new SqlParameter("@Result", SqlDbType.Float));
                pars.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@UseReward", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string warrantName = r.Cells["WarrantName"].Value.ToString();
                    double cr = r.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["CR"].Value);
                    double issueNum = r.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IssueNum"].Value);
                    double equivalentNum = cr * issueNum;
                    double result = r.Cells["Result"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["Result"].Value);
                    string useReward = r.Cells["UseReward"].Value.ToString();
                    string serialNum = r.Cells["SerialNum"].Value.ToString();

                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@CR", cr);
                    h.SetParameterValue("@IssueNum", issueNum);
                    h.SetParameterValue("@EquivalentNum", equivalentNum);
                    h.SetParameterValue("@Result", result);
                    h.SetParameterValue("@UseReward", useReward);
                    h.SetParameterValue("@SerialNum", serialNum);
                    h.ExecuteCommand();
                }
                h.Dispose();

                string cmdText2 = "UPDATE [ApplyOfficial] SET R=@R, IssueNum=@IssueNum, UseReward=@UseReward WHERE SerialNumber=@SerialNumber";
                List<System.Data.SqlClient.SqlParameter> pars2 = new List<System.Data.SqlClient.SqlParameter>();

                pars2.Add(new SqlParameter("@R", SqlDbType.Float));
                pars2.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                pars2.Add(new SqlParameter("@SerialNumber", SqlDbType.VarChar));
                pars2.Add(new SqlParameter("@UseReward", SqlDbType.VarChar));

                SQLCommandHelper h2 = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText2, pars2);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    double cr = r.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["CR"].Value);
                    double issueNum = r.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IssueNum"].Value);
                    string serialNumber = r.Cells["SerialNum"].Value.ToString();
                    string useReward = r.Cells["UseReward"].Value.ToString();

                    h2.SetParameterValue("@R", cr);
                    h2.SetParameterValue("@IssueNum", issueNum);
                    h2.SetParameterValue("@SerialNumber", serialNumber);
                    h2.SetParameterValue("@UseReward", useReward);
                    h2.ExecuteCommand();
                }
                h2.Dispose();

                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 更新搶額度總表");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton() {
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            if (isEdit) {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Default;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                band0.Columns["ApplyKind"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["Market"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["UnderlyingID"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["EquivalentNum"].CellAppearance.BackColor = Color.LightGray;

                band0.Columns["WarrantName"].CellActivation = Activation.AllowEdit;
                band0.Columns["CR"].CellActivation = Activation.AllowEdit;
                band0.Columns["IssueNum"].CellActivation = Activation.AllowEdit;
                band0.Columns["Result"].CellActivation = Activation.AllowEdit;
                band0.Columns["UseReward"].CellActivation = Activation.AllowEdit;

                toolStripButtonReload.Visible = false;
                toolStripButtonEdit.Visible = false;
                toolStripButtonConfirm.Visible = true;
                toolStripButtonCancel.Visible = true;

                //ultraGrid1.DisplayLayout.Bands[0].Columns["ApplyKind"].Hidden = true;
                band0.Columns["TraderID"].Hidden = true;
                band0.Columns["Credit"].Hidden = true;
                band0.Columns["RewardCredit"].Hidden = true;

                ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            } else {
                //ultraGrid1.DisplayLayout.Bands[0].Columns["ApplyKind"].Hidden = false;
                band0.Columns["TraderID"].Hidden = false;
                band0.Columns["Credit"].Hidden = false;
                band0.Columns["RewardCredit"].Hidden = false;

                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                band0.Columns["ApplyKind"].CellAppearance.BackColor = Color.White;
                band0.Columns["Market"].CellAppearance.BackColor = Color.White;
                band0.Columns["UnderlyingID"].CellAppearance.BackColor = Color.White;
                band0.Columns["EquivalentNum"].CellAppearance.BackColor = Color.White;

                band0.Columns["ApplyKind"].CellActivation = Activation.NoEdit;
                band0.Columns["Market"].CellActivation = Activation.NoEdit;
                band0.Columns["TraderID"].CellActivation = Activation.NoEdit;
                band0.Columns["UnderlyingID"].CellActivation = Activation.NoEdit;
                band0.Columns["WarrantName"].CellActivation = Activation.NoEdit;
                band0.Columns["CR"].CellActivation = Activation.NoEdit;
                band0.Columns["IssueNum"].CellActivation = Activation.NoEdit;
                band0.Columns["EquivalentNum"].CellActivation = Activation.NoEdit;
                band0.Columns["Result"].CellActivation = Activation.NoEdit;
                band0.Columns["Credit"].CellActivation = Activation.NoEdit;
                band0.Columns["RewardCredit"].CellActivation = Activation.NoEdit;
                band0.Columns["UseReward"].CellActivation = Activation.NoEdit;

                band0.Columns["ApplyKind"].Width = 70;
                band0.Columns["TraderID"].Width = 70;
                band0.Columns["Market"].Width = 70;
                band0.Columns["UnderlyingID"].Width = 70;
                band0.Columns["WarrantName"].Width = 150;
                band0.Columns["CR"].Width = 70;
                band0.Columns["IssueNum"].Width = 80;
                band0.Columns["EquivalentNum"].Width = 80;
                band0.Columns["Result"].Width = 80;
                band0.Columns["UseReward"].Width = 70;
                ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

                toolStripButtonReload.Visible = true;
                toolStripButtonEdit.Visible = true;
                toolStripButtonConfirm.Visible = false;
                toolStripButtonCancel.Visible = false;

                if (GlobalVar.globalParameter.userGroup == "TR") {
                    toolStripButtonEdit.Visible = false;
                    刪除ToolStripMenuItem.Visible = false;
                }
            }
        }


        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e) {
            isEdit = true;
            SetButton();
        }

        private void toolStripButtonConfirm_Click(object sender, EventArgs e) {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void toolStripButtonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            string applyKind = e.Row.Cells["ApplyKind"].Value.ToString();
            string warrantName = e.Row.Cells["WarrantName"].Value.ToString();
            string serialNum = e.Row.Cells["SerialNum"].Value.ToString();
            string applyStatus = "";
            string applyTime = "";

            double issueNum = Convert.ToDouble(e.Row.Cells["IssueNum"].Value);            

            double equivalentNum = Convert.ToDouble(e.Row.Cells["EquivalentNum"].Value);
            double result = e.Row.Cells["Result"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["Result"].Value);
            //MessageBox.Show()

            string isReward = e.Row.Cells["UseReward"].Value.ToString();            

            if (isReward == "Y")
                e.Row.Cells["UseReward"].Appearance.ForeColor = Color.Blue;

            if (!isEdit && DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime) {

                string sqlTemp = "SELECT [ApplyStatus],[ApplyTime] FROM [EDIS].[dbo].[Apply_71] WHERE SerialNum = '" + serialNum + "'";
                //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                DataTable dtTemp = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                foreach (DataRow drTemp in dtTemp.Rows) {
                    applyStatus = drTemp["ApplyStatus"].ToString();
                    applyTime = drTemp["ApplyTime"].ToString().Substring(0, 2);                    
                }
                if (applyTime == "10" && applyStatus != "X 沒額度" && issueNum != 10000) {
                    e.Row.Cells["IssueNum"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["IssueNum"].ToolTipText = "applyTime == 10 && applyStatus != 沒額度 && issueNum != 10000";
                }

                if (applyStatus == "X 沒額度") {
                    e.Row.Cells["EquivalentNum"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["Result"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["EquivalentNum"].ToolTipText = "沒額度";
                    e.Row.Cells["Result"].ToolTipText = "沒額度";
                }

                //double precision issue
                if (result + 0.00001 >= equivalentNum) {
                    e.Row.Cells["EquivalentNum"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["Result"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["EquivalentNum"].ToolTipText = "額度OK";
                    e.Row.Cells["Result"].ToolTipText = "額度OK";
                }

                if (result + 0.00001 < equivalentNum && result > 0) {
                    e.Row.Cells["EquivalentNum"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["Result"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["EquivalentNum"].ToolTipText = "部分額度";
                    e.Row.Cells["Result"].ToolTipText = "部分額度";
                }
            } else
                e.Row.Appearance.BackColor = Color.White;

            string underlyingID = e.Row.Cells["UnderlyingID"].Value.ToString();           
            string issuable = "NA";
            string accNI = "N";
            string reissuable = "NA";

            string toolTip1 = "標的發行檢查=N";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的虧損";
            string toolTip4 = "今日未達增額標準";
            string sqlTemp2 = "SELECT IsNull(Issuable,'NA') Issuable, CASE WHEN [AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            //DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            DataTable dtTemp2 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            if (dtTemp2.Rows.Count > 0) {
                issuable = dtTemp2.Rows[0]["Issuable"].ToString();
                accNI = dtTemp2.Rows[0]["AccNetIncome"].ToString();
                /*foreach (DataRow drTemp2 in dtTemp2.Rows) {
                    issuable = drTemp2["Issuable"].ToString();
                    accNI = drTemp2["AccNetIncome"].ToString();
                }*/
            }


            if (!isEdit) {
                if (accNI == "Y" && result >= equivalentNum) {
                    e.Row.Cells["UnderlyingID"].ToolTipText = toolTip3;
                    e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Blue;
                }

                if (issuable == "NA") {
                    e.Row.ToolTipText = toolTip2;
                    e.Row.Appearance.ForeColor = Color.Red;
                } else if (issuable == "N") {
                    e.Row.ToolTipText = toolTip1;
                    e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Red;
                }
            }

            if (applyKind == "增額") {
                string sqlTemp3 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [EDIS].[dbo].[WarrantReIssuable] WHERE WarrantName = '" + warrantName + "'";
                //DataView dvTemp3 = DeriLib.Util.ExecSqlQry(sqlTemp3, GlobalVar.loginSet.edisSqlConnString);
                DataTable dtTemp3 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp3, GlobalVar.loginSet.edisSqlConnString);
                if (dtTemp3.Rows.Count > 0)
                    reissuable = dtTemp3.Rows[0]["ReIssuable"].ToString();
                //foreach (DataRow drTemp3 in dtTemp3.Rows)
                //    reissuable = drTemp3["ReIssuable"].ToString();

                if (!isEdit && reissuable == "NA") {
                    e.Row.Cells["WarrantName"].ToolTipText = toolTip4;
                    e.Row.Cells["WarrantName"].Appearance.ForeColor = Color.Red;
                }
            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e) {
            LoadData();
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "CR" || e.Cell.Column.Key == "IssueNum") {
                double cr = e.Cell.Row.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["CR"].Value);
                double issueNum = e.Cell.Row.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["IssueNum"].Value);
                double equivalentNum = cr * issueNum;
                e.Cell.Row.Cells["EquivalentNum"].Value = equivalentNum;
            }
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "UnderlyingID")
                GlobalUtility.MenuItemClick<FrmIssueCheck>().SelectUnderlying((string) e.Cell.Value);
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right)
                contextMenuStrip1.Show();
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e) {
            string applyKind = ultraGrid1.ActiveRow.Cells["ApplyKind"].Value.ToString();
            string warrantName = ultraGrid1.ActiveRow.Cells["WarrantName"].Value.ToString();

            DialogResult result = MessageBox.Show("刪除此檔 " + applyKind + warrantName + "，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                string serialNum = ultraGrid1.ActiveRow.Cells["SerialNum"].Value.ToString();               

                conn.Open();
                EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [ApplyTotalList] WHERE SerialNum='" + serialNum + "'", conn);
                if (applyKind == "新發")
                    EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [ApplyOfficial] WHERE SerialNumber='" + serialNum + "'", conn);
                else if (applyKind == "增額")
                    EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [ReIssueOfficial] WHERE SerialNum='" + serialNum + "'", conn);
                conn.Close();

                LoadData();

                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 刪除一檔" + applyKind + ": " + warrantName);
            }
        }
    }
}
