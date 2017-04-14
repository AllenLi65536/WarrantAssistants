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
    public partial class FrmApplyTotalList : Form
    {
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt = new DataTable();
        private string userID = GlobalVar.globalParameter.userID;
        private bool isEdit = false;
        /*private static Dictionary<int, string> reasonString = new Dictionary<int, string> {
            { 0," "},
            { 1,"技術面偏多，股價持續看好，因此發行認購權證吸引投資人。" },
            { 2,"基本面良好，股價具有漲升的條件，因此發行認購權證吸引投資人。"},
            { 3, "營運動能具提升潛力，因此發行認購權證吸引投資人。"},
            { 4, "提供投資人槓桿避險工具"},
            { 5, "持續針對不同的履約條件、存續期間及認購認售等發行新條件，提供投資人更多元投資選擇"}
        };*/

        public FrmApplyTotalList()
        {
            InitializeComponent();
        }

        private void FrmApplyTotalList_Load(object sender, EventArgs e)
        {
            
            InitialGrid();
            LoadData();
        }

        private void InitialGrid()
        {
            dt.Columns.Add("序號", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("市場", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("張數", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("額度結果", typeof(double));
            dt.Columns.Add("今日額度", typeof(double));
            dt.Columns.Add("獎勵額度", typeof(double));
            dt.Columns.Add("使用獎勵", typeof(string));
            // dt.Columns.Add("發行原因", typeof(string));

            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].Format = "###,###";
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Format = "###,###";

            ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["發行原因"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            // ultraGrid1.DisplayLayout.Bands[0].Columns["發行原因"].Width = 300;
            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            ultraGrid1.DisplayLayout.Bands[0].Columns["序號"].Hidden = true;

            SetButton();
        }

        private void LoadData()
        {
            try
            {
                dt.Rows.Clear();
                string sql = @"SELECT a.[SerialNum]
                              ,CASE WHEN a.[ApplyKind]='1' THEN '新發' ELSE '增額' END ApplyKind
                              ,SUBSTRING(a.[TraderID],4,4) TraderID
                              ,a.[Market]
                              ,a.[UnderlyingID]
                              ,a.[WarrantName]
                              ,a.[CR]
                              ,a.[IssueNum]
                              ,a.[EquivalentNum]
                              ,a.[Result]
                              ,IsNull(b.[IssueCredit],0) Credit
                              ,IsNull(b.[RewardIssueCredit],0) RewardCredit
                              ,a.[UseReward]
                              ,CASE WHEN a.CP='C' THEN c.Reason ELSE c.ReasonP END Reason
                          FROM [EDIS].[dbo].[ApplyTotalList] a 
                          LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                          left join Underlying_TraderIssue c on a.UnderlyingID=c.UID 
                          ORDER BY  a.Market desc, a.ApplyKind, a.SerialNum";// or (a.UnderlyingID = 'IX0001' and c.UID ='TWA00')

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv)
                {
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
                    double result = drv["Result"] == DBNull.Value ? 0.0 : Convert.ToDouble(drv["Result"]);
                    result = Math.Floor(result);
                    double credit = (double) drv["Credit"];
                    credit = Math.Floor(credit);
                    double rewardCredit = (double)drv["RewardCredit"];
                    rewardCredit = Math.Floor(rewardCredit);
                    dr["額度結果"] = result;
                    dr["今日額度"] = credit;
                    dr["獎勵額度"] = rewardCredit;
                    dr["使用獎勵"] = drv["UseReward"].ToString();
                    // dr["發行原因"] = drv["Reason"] == DBNull.Value ? " " : reasonString[Convert.ToInt32(drv["Reason"])];
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

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                {
                    string warrantName = r.Cells["權證名稱"].Value.ToString();
                    double cr = r.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["行使比例"].Value);
                    double issueNum = r.Cells["張數"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["張數"].Value);
                    double equivalentNum = cr*issueNum;
                    double result = r.Cells["額度結果"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["額度結果"].Value);
                    string useReward = r.Cells["使用獎勵"].Value.ToString();
                    string serialNum = r.Cells["序號"].Value.ToString();

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

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                {
                    double cr = r.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["行使比例"].Value);
                    double issueNum = r.Cells["張數"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["張數"].Value);
                    string serialNumber = r.Cells["序號"].Value.ToString();
                    string useReward = r.Cells["使用獎勵"].Value.ToString();

                    h2.SetParameterValue("@R", cr);
                    h2.SetParameterValue("@IssueNum", issueNum);
                    h2.SetParameterValue("@SerialNumber", serialNumber);
                    h2.SetParameterValue("@UseReward", useReward);
                    h2.ExecuteCommand();
                }
                h2.Dispose();

                GlobalUtility.logInfo("Info", GlobalVar.globalParameter.userID + " 更新搶額度總表");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Info");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 更新搶額度總表");
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
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Default;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellAppearance.BackColor = Color.LightGray;

                ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].CellActivation = Activation.AllowEdit;

                toolStripButtonReload.Visible = false;
                toolStripButtonEdit.Visible = false;
                toolStripButtonConfirm.Visible = true;
                toolStripButtonCancel.Visible = true;

                //ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = true;

                ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            }
            else
            {
                //ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = false;

                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellAppearance.BackColor = Color.White;

                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].CellActivation = Activation.NoEdit;
 
                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].Width = 70;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].Width = 70;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].Width = 70;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].Width = 70;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].Width = 150;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].Width = 70;
                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].Width = 80;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Width = 80;
                ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].Width = 80;
                ultraGrid1.DisplayLayout.Bands[0].Columns["使用獎勵"].Width = 70;
                ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

                toolStripButtonReload.Visible = true;
                toolStripButtonEdit.Visible = true;
                toolStripButtonConfirm.Visible = false;
                toolStripButtonCancel.Visible = false;

                if (GlobalVar.globalParameter.userGroup == "TR")
                {
                    toolStripButtonEdit.Visible = false;
                    刪除ToolStripMenuItem.Visible = false;
                }


            }
        }


        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e)
        {
            isEdit = true;
            SetButton();
        }

        private void toolStripButtonConfirm_Click(object sender, EventArgs e)
        {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void toolStripButtonCancel_Click(object sender, EventArgs e)
        {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string applyKind = e.Row.Cells["類型"].Value.ToString();
            string warrantName = e.Row.Cells["權證名稱"].Value.ToString();
            string serialNum = e.Row.Cells["序號"].Value.ToString();
            string applyStatus = "";
            string applyTime = "";

            double issueNum = 0.0;
            issueNum = Convert.ToDouble(e.Row.Cells["張數"].Value);

            double equivalentNum = Convert.ToDouble(e.Row.Cells["約當張數"].Value);
            double result = e.Row.Cells["額度結果"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["額度結果"].Value);
            //MessageBox.Show()

            string isReward = "N";
            isReward = e.Row.Cells["使用獎勵"].Value.ToString();

            if (isReward == "Y")
                e.Row.Cells["使用獎勵"].Appearance.ForeColor = Color.Blue;

            if (!isEdit && DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime)
            {

                string sqlTemp = "SELECT [ApplyStatus],[ApplyTime] FROM [EDIS].[dbo].[Apply_71] WHERE SerialNum = '" + serialNum + "'";
                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                foreach (DataRowView drTemp in dvTemp)
                {
                    applyStatus = drTemp["ApplyStatus"].ToString();
                    applyTime = drTemp["ApplyTime"].ToString();
                    applyTime = applyTime.Substring(0, 2);
                }
                if (applyTime == "10" && applyStatus != "X 沒額度" && issueNum != 10000) {
                    e.Row.Cells["張數"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["張數"].ToolTipText = "applyTime == 10 && applyStatus != 沒額度 && issueNum != 10000";
                }

                if (applyStatus == "X 沒額度")
                {
                    e.Row.Cells["約當張數"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["額度結果"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["約當張數"].ToolTipText = "沒額度";
                    e.Row.Cells["額度結果"].ToolTipText = "沒額度";
                }
                
                //double precision
                if (result+0.00001 >= equivalentNum)
                {
                    e.Row.Cells["約當張數"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["額度結果"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["約當張數"].ToolTipText = "額度OK";
                    e.Row.Cells["額度結果"].ToolTipText = "額度OK";
                }

                if (result+0.00001 < equivalentNum && result > 0)
                {
                    e.Row.Cells["約當張數"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["額度結果"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["約當張數"].ToolTipText = "部分額度";
                    e.Row.Cells["額度結果"].ToolTipText = "部分額度";
                }
            }
            else
                e.Row.Appearance.BackColor = Color.White;

            string underlyingID = "";
            underlyingID = e.Row.Cells["標的代號"].Value.ToString();
            string issuable = "NA";
            string accNI = "N";
            string reissuable = "NA";

            string toolTip1 = "標的發行檢查=N";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的虧損";
            string toolTip4 = "今日未達增額標準";
            string sqlTemp2 = "SELECT IsNull(Issuable,'NA') Issuable, CASE WHEN [AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp2.Count > 0)
            {
                foreach (DataRowView drTemp2 in dvTemp2)
                {
                    issuable = drTemp2["Issuable"].ToString();
                    accNI = drTemp2["AccNetIncome"].ToString();
                }
            }


            if (!isEdit)
            {
                if (accNI == "Y" && result >= equivalentNum)
                {
                    e.Row.Cells["標的代號"].ToolTipText = toolTip3;
                    e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Blue;
                }

                if (issuable == "NA")
                {
                    e.Row.ToolTipText = toolTip2;
                    e.Row.Appearance.ForeColor = Color.Red;
                }
                else if (issuable == "N")
                {
                    e.Row.ToolTipText = toolTip1;
                    e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
                }

            }

            if (applyKind == "增額")
            {
                string sqlTemp3 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [EDIS].[dbo].[WarrantReIssuable] WHERE WarrantName = '" + warrantName + "'";
                DataView dvTemp3 = DeriLib.Util.ExecSqlQry(sqlTemp3, GlobalVar.loginSet.edisSqlConnString);
                if (dvTemp3.Count > 0)
                {
                    foreach (DataRowView drTemp3 in dvTemp3)
                    {
                        reissuable = drTemp3["ReIssuable"].ToString();
                    }
                }

                if (!isEdit && reissuable == "NA")
                {
                    e.Row.Cells["權證名稱"].ToolTipText = toolTip4;
                    e.Row.Cells["權證名稱"].Appearance.ForeColor = Color.Red;
                }
            }

        }

        private void toolStripButtonReload_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e)
        {
            if (e.Cell.Column.Key == "行使比例" || e.Cell.Column.Key == "張數")
            {
                double cr = e.Cell.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["行使比例"].Value);
                double issueNum = e.Cell.Row.Cells["張數"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["張數"].Value);
                double equivalentNum = cr * issueNum;
                e.Cell.Row.Cells["約當張數"].Value = equivalentNum;

            }
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e)
        {
            if (e.Cell.Column.Key == "標的代號")
            {
                string target = (string)e.Cell.Value;
                FrmIssueCheck frmIssueCheck = null;

                foreach (Form iForm in Application.OpenForms)
                {
                    if (iForm.GetType() == typeof(FrmIssueCheck))
                    {
                        frmIssueCheck = (FrmIssueCheck)iForm;
                        break;
                    }
                }

                if (frmIssueCheck != null)
                    frmIssueCheck.BringToFront();
                else
                {
                    frmIssueCheck = new FrmIssueCheck();
                    frmIssueCheck.StartPosition = FormStartPosition.CenterScreen;
                    frmIssueCheck.Show();
                }
                frmIssueCheck.selectUnderlying(target);
            }
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show();
            }
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string applyKind = "";
            applyKind = ultraGrid1.ActiveRow.Cells["類型"].Value.ToString();

            string warrantName = "";
            warrantName = ultraGrid1.ActiveRow.Cells["權證名稱"].Value.ToString();

            DialogResult result = MessageBox.Show("刪除此檔 "+applyKind+warrantName+"，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                string serialNum="";
                serialNum = ultraGrid1.ActiveRow.Cells["序號"].Value.ToString();                

                string sql1 = "DELETE FROM [ApplyTotalList] WHERE SerialNum='" + serialNum + "'";
                string sql2 = "DELETE FROM [ApplyOfficial] WHERE SerialNumber='" + serialNum + "'";
                string sql3 = "DELETE FROM [ReIssueOfficial] WHERE SerialNum='" + serialNum + "'";
                SqlCommand cmd1 = new SqlCommand(sql1, conn);
                SqlCommand cmd2 = new SqlCommand(sql2, conn);
                SqlCommand cmd3 = new SqlCommand(sql3, conn);

                conn.Open();
                cmd1.ExecuteNonQuery();
                cmd1.Dispose();
                if (applyKind == "新發")
                {
                    cmd2.ExecuteNonQuery();
                    cmd2.Dispose();
                }
                else if (applyKind == "增額")
                {
                    cmd3.ExecuteNonQuery();
                    cmd3.Dispose();
                }
                conn.Close();

                LoadData();

                GlobalUtility.logInfo("Info", GlobalVar.globalParameter.userID + " 刪除一檔" + applyKind + ": " + warrantName);
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Info");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 刪除一檔" + applyKind + ": " + warrantName);
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/
            }
            
        }



    }
}
