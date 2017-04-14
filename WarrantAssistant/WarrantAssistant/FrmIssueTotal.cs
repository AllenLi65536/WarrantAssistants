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
    public partial class FrmIssueTotal:Form
    {
        private DataTable dt = new DataTable();
        private string userID = GlobalVar.globalParameter.userID;
        private bool isEdit = false;

        private static Dictionary<int, string> reasonString = new Dictionary<int, string> {
            { 0," "},
            { 1,"技術面偏多，股價持續看好，因此發行認購權證吸引投資人。" },
            { 2,"基本面良好，股價具有漲升的條件，因此發行認購權證吸引投資人。"},
            { 3, "營運動能具提升潛力，因此發行認購權證吸引投資人。"},
            { 4, "提供投資人槓桿避險工具"},
            { 5, "持續針對不同的履約條件、存續期間及認購認售等發行新條件，提供投資人更多元投資選擇"}
        };

        public FrmIssueTotal() {
            InitializeComponent();
        }

        private void FrmIssueTotal_Load(object sender, EventArgs e) {
            InitialGrid();
            LoadData();
        }

        private void InitialGrid() {
            dt.Columns.Add("市場", typeof(string));
            dt.Columns.Add("序號", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("發行價格", typeof(double));
            dt.Columns.Add("1500W", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("CP", typeof(string));
            dt.Columns.Add("股價", typeof(double));
            dt.Columns.Add("履約價", typeof(double));
            dt.Columns.Add("期間", typeof(int));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("HV", typeof(double));
            dt.Columns.Add("IV", typeof(double));
            dt.Columns.Add("IVOri", typeof(double));
            dt.Columns.Add("重設比", typeof(double));
            dt.Columns.Add("界限比", typeof(double));
            dt.Columns.Add("財務費用", typeof(double));
            dt.Columns.Add("張數", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("額度結果", typeof(double));
            dt.Columns.Add("發行原因", typeof(string));

            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            band0.Columns["張數"].Format = "###,###";
            band0.Columns["約當張數"].Format = "###,###";
            band0.Columns["額度結果"].Format = "###,###";

            band0.Columns["交易員"].Width = 60;
            band0.Columns["權證名稱"].Width = 150;
            band0.Columns["發行價格"].Width = 70;
            band0.Columns["標的代號"].Width = 70;
            band0.Columns["1500W"].Width = 70;
            band0.Columns["市場"].Width = 50;
            band0.Columns["類型"].Width = 70;
            band0.Columns["CP"].Width = 40;
            band0.Columns["股價"].Width = 70;
            band0.Columns["履約價"].Width = 70;
            band0.Columns["期間"].Width = 40;
            band0.Columns["行使比例"].Width = 70;
            band0.Columns["HV"].Width = 40;
            band0.Columns["IV"].Width = 40;
            band0.Columns["重設比"].Width = 70;
            band0.Columns["界限比"].Width = 70;
            band0.Columns["財務費用"].Width = 70;
            band0.Columns["張數"].Width = 80;
            band0.Columns["約當張數"].Width = 80;
            band0.Columns["額度結果"].Width = 80;
            band0.Columns["發行原因"].Width = 300;

            band0.Columns["1500W"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["標的代號"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["類型"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["CP"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["期間"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["發行價格"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["股價"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["履約價"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["HV"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["IV"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["重設比"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["界限比"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["財務費用"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["約當張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["額度結果"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["發行原因"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            band0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            band0.Columns["序號"].Hidden = true;
            band0.Columns["IVOri"].Hidden = true;

            // To sort multi-column using SortedColumns property
            // This enables multi-column sorting
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

            // It is good practice to clear the sorted columns collection
            band0.SortedColumns.Clear();

            SetButton();
        }

        private void LoadData() {
            try {
                dt.Rows.Clear();
                string sql = @"SELECT  a.SerialNumber
                                      ,SUBSTRING(a.TraderID,4,4) TraderID
	                                  ,b.WarrantName
	                                  ,a.UnderlyingID
                                      ,a.Apply1500W
	                                  ,b.Market
	                                  ,a.Type
	                                  ,a.CP
	                                  ,IsNull(c.MPrice,0) MPrice
                                      ,a.K
                                      ,a.T
                                      ,a.R
                                      ,a.HV
                                      ,CASE WHEN a.Apply1500W='Y' THEN a.IVNew ELSE a.IV END IVNew
                                      ,a.ResetR
                                      ,a.BarrierR
                                      ,a.FinancialR
                                      ,a.IssueNum
                                      ,b.EquivalentNum
                                      ,b.Result
                                      ,a.IV
                                      ,CASE WHEN a.CP='C' THEN d.Reason ELSE d.ReasonP END Reason
                                  FROM [EDIS].[dbo].[ApplyOfficial] a
                                  LEFT JOIN [EDIS].[dbo].[ApplyTotalList] b ON a.SerialNumber=b.SerialNum
                                  LEFT JOIN [EDIS].[dbo].[WarrantPrices] c on a.UnderlyingID=c.CommodityID
                                left join Underlying_TraderIssue d on a.UnderlyingID=d.UID 
                                  ORDER BY b.Market desc, a.SerialNumber"; //or (a.UnderlyingID = 'IX0001' and d.UID ='TWA00')

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    dr["序號"] = drv["SerialNumber"].ToString();
                    dr["交易員"] = drv["TraderID"].ToString();
                    dr["權證名稱"] = drv["WarrantName"].ToString();
                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    dr["1500W"] = drv["Apply1500W"].ToString();
                    dr["市場"] = drv["Market"].ToString();
                    dr["張數"] = drv["IssueNum"];
                    dr["約當張數"] = drv["EquivalentNum"];
                    dr["額度結果"] = drv["Result"];
                    dr["IVOri"] = drv["IV"];

                    double underlyingPrice = 0.0;
                    underlyingPrice = Convert.ToDouble(drv["MPrice"]);
                    dr["股價"] = underlyingPrice;
                    double k = Convert.ToDouble(drv["K"]);
                    dr["履約價"] = k;
                    int t = Convert.ToInt32(drv["T"]);
                    dr["期間"] = t;
                    double cr = Convert.ToDouble(drv["R"]);
                    dr["行使比例"] = cr;
                    dr["HV"] = Convert.ToDouble(drv["HV"]);
                    double vol = Convert.ToDouble(drv["IVNew"]) / 100;
                    dr["IV"] = Convert.ToDouble(drv["IVNew"]);

                    double resetR = Convert.ToDouble(drv["ResetR"]) / 100;
                    dr["重設比"] = Convert.ToDouble(drv["ResetR"]);
                    double barrierR = Convert.ToDouble(drv["BarrierR"]);
                    dr["界限比"] = barrierR;
                    double financialR = Convert.ToDouble(drv["FinancialR"]) / 100;
                    dr["財務費用"] = Convert.ToDouble(drv["FinancialR"]);
                    string warrantType = drv["Type"].ToString();
                    dr["類型"] = warrantType;
                    CallPutType cp = drv["CP"].ToString() == "C" ? CallPutType.Call : CallPutType.Put;
                    dr["CP"] = drv["CP"].ToString();
                    dr["發行原因"] = drv["Reason"] == DBNull.Value ? " " : reasonString[Convert.ToInt32(drv["Reason"])];

                    double price = 0.0;
                    if (underlyingPrice != 0) {
                        if (warrantType == "牛熊證")
                            price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                        else if (warrantType == "重設型")
                            price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                        else
                            price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                    }
                    dr["發行價格"] = Math.Round(price, 2);

                    dt.Rows.Add(dr);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateData() {
            try {
                string cmdText = "UPDATE [ApplyOfficial] SET K=@K, T=@T, HV=@HV, IV=@IV, ResetR=@ResetR, BarrierR=@BarrierR, FinancialR=@FinancialR, Type=@Type, CP=@CP, Apply1500W=@Apply1500W, MDate=@MDate WHERE SerialNumber=@SerialNumber";
                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                pars.Add(new SqlParameter("@K", SqlDbType.Float));
                pars.Add(new SqlParameter("@T", SqlDbType.Int));
                pars.Add(new SqlParameter("@HV", SqlDbType.Float));
                pars.Add(new SqlParameter("@IV", SqlDbType.Float));
                pars.Add(new SqlParameter("@ResetR", SqlDbType.Float));
                pars.Add(new SqlParameter("@BarrierR", SqlDbType.Float));
                pars.Add(new SqlParameter("@FinancialR", SqlDbType.Float));
                pars.Add(new SqlParameter("@Type", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@CP", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@Apply1500W", SqlDbType.VarChar));
                //pars.Add(new SqlParameter("@TraderID", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                pars.Add(new SqlParameter("@SerialNumber", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    double k = r.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價"].Value);
                    double t = r.Cells["期間"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["期間"].Value);
                    double hv = r.Cells["HV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["HV"].Value);
                    double iv = 0.0;
                    double resetR = r.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["重設比"].Value);
                    double barrierR = r.Cells["界限比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["界限比"].Value);
                    double financialR = r.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["財務費用"].Value);
                    string type = r.Cells["類型"].Value.ToString();
                    string cp = r.Cells["CP"].Value.ToString();
                    string apply1500w = r.Cells["1500W"].Value.ToString();
                    string serialNumber = r.Cells["序號"].Value.ToString();
                    //string traderID = "000"+r.Cells["交易員"].Value.ToString();

                    if (apply1500w == "Y")
                        iv = r.Cells["IVOri"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IVOri"].Value);
                    else
                        iv = r.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IV"].Value);

                    h.SetParameterValue("@K", k);
                    h.SetParameterValue("@T", t);
                    h.SetParameterValue("@HV", hv);
                    h.SetParameterValue("@IV", iv);
                    h.SetParameterValue("@ResetR", resetR);
                    h.SetParameterValue("@BarrierR", barrierR);
                    h.SetParameterValue("@FinancialR", financialR);
                    h.SetParameterValue("@Type", type);
                    h.SetParameterValue("@CP", cp);
                    h.SetParameterValue("@Apply1500W", apply1500w);
                    //h.SetParameterValue("@TraderID", traderID);
                    h.SetParameterValue("@MDate", DateTime.Now);
                    h.SetParameterValue("@SerialNumber", serialNumber);
                    h.ExecuteCommand();
                }
                h.Dispose();

                GlobalUtility.logInfo("Info", GlobalVar.globalParameter.userID + " 更新發行總表");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Info");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 更新發行總表");
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton() {
            if (isEdit) {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Default;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["CP"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["履約價"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["期間"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["HV"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["IV"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["重設比"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["界限比"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["財務費用"].CellActivation = Activation.AllowEdit;

                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行價格"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellAppearance.BackColor = Color.LightGray;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["股價"].CellAppearance.BackColor = Color.LightGray;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellAppearance.BackColor = Color.LightGray;

                toolStripButtonReload.Visible = false;
                toolStripButtonEdit.Visible = false;
                toolStripButtonConfirm.Visible = true;
                toolStripButtonCancel.Visible = true;

                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].Hidden = true;

            } else {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["權證名稱"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行價格"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["CP"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["股價"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["履約價"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["期間"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["HV"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["IV"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["重設比"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["界限比"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["財務費用"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].CellActivation = Activation.NoEdit;

                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行價格"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellAppearance.BackColor = Color.White;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["股價"].CellAppearance.BackColor = Color.White;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellAppearance.BackColor = Color.White;

                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["額度結果"].Hidden = false;

                toolStripButtonReload.Visible = true;
                toolStripButtonEdit.Visible = true;
                toolStripButtonConfirm.Visible = false;
                toolStripButtonCancel.Visible = false;

                if (GlobalVar.globalParameter.userGroup == "TR") {
                    toolStripButtonEdit.Visible = false;
                }

            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e) {
            LoadData();
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

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            string is1500W = "N";
            is1500W = e.Row.Cells["1500W"].Value.ToString();

            if (is1500W == "Y")
                e.Row.Cells["1500W"].Appearance.ForeColor = Color.Blue;

            string underlyingID = "";
            string cp = "C";
            underlyingID = e.Row.Cells["標的代號"].Value.ToString();
            cp = e.Row.Cells["CP"].Value.ToString();
            string issuable = "Y";
            string putIssuable = "Y";
            string toolTip1 = "發行檢查=N";
            string sqlTemp2 = "SELECT [Issuable], [PutIssuable] FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            foreach (DataRowView drTemp2 in dvTemp2) {
                issuable = drTemp2["Issuable"].ToString();
                putIssuable = drTemp2["PutIssuable"].ToString();
            }
            if (underlyingID != "") {

                if (issuable == "N") {
                    e.Row.ToolTipText = toolTip1;
                    e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
                }

                if (cp == "P" && putIssuable == "N") {
                    e.Row.Cells["CP"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["CP"].ToolTipText = "Put not issuable";
                }

            }

            double issuePrice = e.Row.Cells["發行價格"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["發行價格"].Value);
            if (issuePrice <= 0.6 || issuePrice > 3) {
                e.Row.Cells["發行價格"].Appearance.ForeColor = Color.Red;
                e.Row.Cells["發行價格"].ToolTipText = " <= 0.6 or > 3";
            }

            string warrantType = e.Row.Cells["類型"].Value == DBNull.Value ? "一般型" : e.Row.Cells["類型"].Value.ToString();
            double k = e.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["履約價"].Value);
            double underlyingPrice = e.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["股價"].Value);

            //Check for moneyness constraint
            e.Row.Cells["履約價"].Appearance.ForeColor = Color.Black;
            if (warrantType != "牛熊證") {
                if ((cp == "C" && k / underlyingPrice >= 1.5) || (cp == "P" && k / underlyingPrice <= 0.5)) {
                    e.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["履約價"].ToolTipText = "履約價超過價外50%";
                }
            }

            if (!isEdit && DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime) {
                string warrantName = e.Row.Cells["權證名稱"].Value.ToString();
                string applyStatus = "";
                string serialNum = e.Row.Cells["序號"].Value.ToString();
                double issueNum = 0.0;
                issueNum = Convert.ToDouble(e.Row.Cells["張數"].Value);

                double equivalentNum = Convert.ToDouble(e.Row.Cells["約當張數"].Value);
                double result = e.Row.Cells["額度結果"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["額度結果"].Value);

                string sqlTemp = "SELECT [ApplyStatus] FROM [EDIS].[dbo].[Apply_71] WHERE SerialNum = '" + serialNum + "'";
                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                foreach (DataRowView drTemp in dvTemp) {
                    applyStatus = drTemp["ApplyStatus"].ToString();
                }
                if (applyStatus == "排隊中" && issueNum != 10000) {
                    e.Row.Cells["張數"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["張數"].ToolTipText = "排隊中";
                }

                if (applyStatus == "X 沒額度") {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["權證名稱"].ToolTipText = "沒額度";
                }

                if (result >= equivalentNum) {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["權證名稱"].ToolTipText = "額度OK";
                }

                if (result < equivalentNum && result > 0) {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["權證名稱"].ToolTipText = "部分額度";
                }

            }
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key != "交易員" && e.Cell.Column.Key != "權證名稱" && e.Cell.Column.Key != "發行價格" && e.Cell.Column.Key != "標的代號" && e.Cell.Column.Key != "市場" && e.Cell.Column.Key != "1500W") {
                double price = 0.0;

                double underlyingPrice = 0.0;
                underlyingPrice = e.Cell.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["股價"].Value);
                double k = 0.0;
                k = e.Cell.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["履約價"].Value);
                int t = 0;
                t = e.Cell.Row.Cells["期間"].Value == DBNull.Value ? 0 : Convert.ToInt32(e.Cell.Row.Cells["期間"].Value);
                double cr = 0.0;
                cr = e.Cell.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["行使比例"].Value);
                double vol = 0.0;
                vol = e.Cell.Row.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["IV"].Value) / 100;
                double resetR = 0.0;
                resetR = e.Cell.Row.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["重設比"].Value) / 100;
                double financialR = 0.0;
                financialR = e.Cell.Row.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["財務費用"].Value) / 100;
                string warrantType = "一般型";
                warrantType = e.Cell.Row.Cells["類型"].Value == DBNull.Value ? "一般型" : e.Cell.Row.Cells["類型"].Value.ToString();
                string cpType = "C";
                cpType = e.Cell.Row.Cells["CP"].Value == DBNull.Value ? "C" : e.Cell.Row.Cells["CP"].Value.ToString();

                if (warrantType != "一般型" && warrantType != "牛熊證" && warrantType != "重設型") {
                    if (warrantType == "2")
                        warrantType = "牛熊證";
                    else if (warrantType == "3")
                        warrantType = "重設型";
                    else
                        warrantType = "一般型";
                }

                if (cpType != "C" && cpType != "P") {
                    if (cpType == "2")
                        cpType = "P";
                    else
                        cpType = "C";
                }

                CallPutType cp = CallPutType.Call;
                if (cpType == "P")
                    cp = CallPutType.Put;
                else
                    cp = CallPutType.Call;

                if (warrantType == "牛熊證")
                    price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                else if (warrantType == "重設型")
                    price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                else
                    price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);

                /*e.Cell.Row.Cells["履約價"].Appearance.ForeColor = Color.Black;
                if (warrantType != "牛熊證") {
                    if (cpType == "C" && k / underlyingPrice >= 1.5) {
                        e.Cell.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                    } else if (cpType == "P" && k / underlyingPrice <= 0.5) {
                        e.Cell.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                    }
                }*/

                double shares = 0.0;
                shares = e.Cell.Row.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(e.Cell.Row.Cells["張數"].Value);
                /*
                string is1500W = "N";
                is1500W = e.Cell.Row.Cells["1500W"].Value == DBNull.Value ? "N" : (string)e.Cell.Row.Cells["1500W"].Value;
                if (e.Cell.Column.Key == "1500W" && is1500W=="Y")
                {
                    double totalValue = 0.0;
                    totalValue = price * shares * 1000;
                    while (totalValue < 15000000)
                    {
                        vol += 0.01;
                        if (warrantType == "牛熊證")
                            price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                        else if (warrantType == "重設型")
                            price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                        else
                            price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                        totalValue = price * shares * 1000;
                    }
                    e.Cell.Row.Cells["IV"].Value = Math.Round(vol * 100, 0);
                }
                 * */

                e.Cell.Row.Cells["發行價格"].Value = Math.Round(price, 2);
            }

            string is1500W = "N";
            is1500W = e.Cell.Row.Cells["1500W"].Value.ToString();
            if (e.Cell.Column.Key == "1500W") {
                if (is1500W == "N")
                    e.Cell.Row.Cells["IV"].Value = e.Cell.Row.Cells["IVOri"].Value;
            }
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "標的代號") {
                string target = (string) e.Cell.Value;
                FrmIssueCheck frmIssueCheck = null;

                foreach (Form iForm in Application.OpenForms) {
                    if (iForm.GetType() == typeof(FrmIssueCheck)) {
                        frmIssueCheck = (FrmIssueCheck) iForm;
                        break;
                    }
                }

                if (frmIssueCheck != null)
                    frmIssueCheck.BringToFront();
                else {
                    frmIssueCheck = new FrmIssueCheck();
                    frmIssueCheck.StartPosition = FormStartPosition.CenterScreen;
                    frmIssueCheck.Show();
                }
                frmIssueCheck.selectUnderlying(target);
            }

            if (e.Cell.Column.Key == "CP") {
                string target = (string) e.Cell.Row.Cells["標的代號"].Value;
                FrmIssueCheckPut frmIssueCheckPut = null;

                foreach (Form iForm in Application.OpenForms) {
                    if (iForm.GetType() == typeof(FrmIssueCheckPut)) {
                        frmIssueCheckPut = (FrmIssueCheckPut) iForm;
                        break;
                    }
                }

                if (frmIssueCheckPut != null)
                    frmIssueCheckPut.BringToFront();
                else {
                    frmIssueCheckPut = new FrmIssueCheckPut();
                    frmIssueCheckPut.StartPosition = FormStartPosition.CenterScreen;
                    frmIssueCheckPut.Show();
                }
                frmIssueCheckPut.selectUnderlying(target);
            }
        }


    }
}
