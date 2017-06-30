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
using EDLib.SQL;


namespace WarrantAssistant
{
    public partial class FrmApply:Form
    {
        //HeaderUIElement
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dt = new DataTable();
        private bool isEdit = false;
        public string userID = GlobalVar.globalParameter.userID;
        public string userName;
        private int applyCount = 0;

        public FrmApply() {
            InitializeComponent();
        }

        private void FrmApply_Load(object sender, EventArgs e) {
            toolStripLabel1.Text = "使用者: " + userName;
            toolStripLabel2.Text = "";
            InitialGrid();
            LoadData();

        }

        private void InitialGrid() {
            //dt.Columns.Add("編號", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("履約價", typeof(double));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("HV", typeof(double));
            dt.Columns.Add("IV", typeof(double));
            dt.Columns.Add("期間(月)", typeof(int));
            dt.Columns.Add("張數", typeof(double));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("CP", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("重設比", typeof(double));
            dt.Columns.Add("界限比", typeof(double));
            dt.Columns.Add("財務費用", typeof(double));
            dt.Columns.Add("獎勵", typeof(bool));
            dt.Columns["獎勵"].ReadOnly = false;
            dt.Columns.Add("1500W", typeof(bool));
            dt.Columns["1500W"].ReadOnly = false;
            dt.Columns.Add("確認", typeof(bool));
            dt.Columns["確認"].ReadOnly = false;
            dt.Columns.Add("發行原因", typeof(string));
            dt.Columns.Add("發行價格", typeof(double));
            dt.Columns.Add("標的名稱", typeof(string));
            dt.Columns.Add("股價", typeof(double));
            dt.Columns.Add("Delta", typeof(double));
            dt.Columns.Add("跳動價差", typeof(double));
            dt.Columns.Add("IV*", typeof(double));
            dt.Columns.Add("發行價格*", typeof(double));
            dt.Columns.Add("跌停價*", typeof(double));
            dt.Columns.Add("市場", typeof(string));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("今日額度", typeof(double));
            dt.Columns.Add("獎勵額度", typeof(double));

            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Header.Appearance.
            band0.Columns["張數"].Format = "###,###";
            band0.Columns["約當張數"].Format = "###,###";
            band0.Columns["今日額度"].Format = "###,###";
            band0.Columns["獎勵額度"].Format = "###,###";

            band0.Columns["類型"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            band0.Columns["CP"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            band0.Columns["交易員"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            band0.Columns["發行原因"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["刪除"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;

            //ultraGrid1.DisplayLayout.Bands[0].Columns["編號"].Width = 75;

            band0.Columns["標的代號"].Width = 60;
            band0.Columns["履約價"].Width = 60;
            band0.Columns["期間(月)"].Width = 55;
            band0.Columns["行使比例"].Width = 60;
            band0.Columns["HV"].Width = 50;
            band0.Columns["IV"].Width = 50;
            band0.Columns["張數"].Width = 60;
            band0.Columns["重設比"].Width = 60;
            band0.Columns["界限比"].Width = 60;
            band0.Columns["財務費用"].Width = 60;
            band0.Columns["類型"].Width = 60;
            band0.Columns["CP"].Width = 30;
            band0.Columns["交易員"].Width = 70;
            band0.Columns["獎勵"].Width = 40;
            band0.Columns["確認"].Width = 40;
            band0.Columns["1500W"].Width = 50;
            band0.Columns["發行價格"].Width = 60;
            band0.Columns["發行原因"].Width = 150;
            band0.Columns["標的名稱"].Width = 70;
            band0.Columns["股價"].Width = 60;
            band0.Columns["Delta"].Width = 70;
            band0.Columns["跳動價差"].Width = 70;
            band0.Columns["市場"].Width = 40;
            band0.Columns["IV*"].Width = 60;
            band0.Columns["發行價格*"].Width = 60;
            band0.Columns["跌停價*"].Width = 60;
            band0.Columns["約當張數"].Width = 60;
            band0.Columns["今日額度"].Width = 60;
            band0.Columns["獎勵額度"].Width = 60;
           

            band0.Columns["發行價格"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["標的名稱"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["股價"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["Delta"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["跳動價差"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["IV*"].CellAppearance.BackColor = Color.LightBlue;
            band0.Columns["發行價格*"].CellAppearance.BackColor = Color.LightBlue;
            band0.Columns["跌停價*"].CellAppearance.BackColor = Color.LightBlue;
            band0.Columns["市場"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["約當張數"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["今日額度"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["獎勵額度"].CellAppearance.BackColor = Color.LightGray;

            //band0.Columns["標的代號"].SortIndicator = SortIndicator.None;
                        
            // To sort multi-column using SortedColumns property
            // This enables multi-column sorting
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

            // It is good practice to clear the sorted columns collection
            band0.SortedColumns.Clear();

            // You can sort (as well as group rows by) columns by using SortedColumns 
            // property off the band
            //band0.SortedColumns.Add("ContactName", false, false);

            //ultraGrid1.DisplayLayout.Bands[0].Columns["可發行股數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["截至前一日"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["本日累積發行"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["累計%"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;


            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].CellActivation = Activation.AllowEdit;


            SetButton();
        }

        private void LoadData() {
            try {
                dt.Rows.Clear();
                string sql = @"SELECT a.UnderlyingID
                                  ,a.K
                                  ,a.T
                                  ,a.R
                                  ,a.HV
                                  ,a.IV
                                  ,a.IssueNum
                                  ,a.ResetR
                                  ,a.BarrierR
                                  ,a.FinancialR
                                  ,a.Type
                                  ,a.CP
                                  ,a.TraderID
                                  ,CASE WHEN a.UseReward='Y' THEN 1 ELSE 0 END UseReward
                                  ,CASE WHEN a.ConfirmChecked='Y' THEN 1 ELSE 0 END ConfirmChecked
                                  ,CASE WHEN a.Apply1500W='Y' THEN 1 ELSE 0 END Apply1500W
	                              ,b.UnderlyingName
	                              ,c.MPrice
	                              ,b.Market
	                              ,(a.IssueNum*a.R) as EquivalentNum
	                              ,IsNull(b.[IssueCredit],0) IssueCredit
	                              ,IsNull(b.[RewardIssueCredit],0) RewardIssueCredit
                                  ,CASE WHEN a.CP='C' THEN d.Reason ELSE d.ReasonP END Reason
                              FROM [EDIS].[dbo].[ApplyTempList] a ";
                sql += @"LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] b on a.UnderlyingID=b.UnderlyingID
                   LEFT JOIN [EDIS].[dbo].[WarrantPrices] c on a.UnderlyingID=c.CommodityID 
                    left join [Underlying_TraderIssue] d on d.UID = a.UnderlyingID 
                   WHERE a.UserID='" + userID + "' ";//or (a.UnderlyingID = 'IX0001' and d.UID ='TWA00')
                sql += "ORDER BY a.MDate";

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                if (dv.Count > 0) {
                    foreach (DataRowView drv in dv) {
                        DataRow dr = dt.NewRow();

                        //dr["編號"] = drv["SerialNum"].ToString();

                        string underlyingID = drv["UnderlyingID"].ToString();
                        dr["標的代號"] = underlyingID;
                        double k = Convert.ToDouble(drv["K"]);
                        dr["履約價"] = k;
                        // ultraGrid1.DisplayLayout.Bands[0].Columns["履約價
                        int t = Convert.ToInt32(drv["T"]);
                        dr["期間(月)"] = t;
                        double cr = Convert.ToDouble(drv["R"]);
                        dr["行使比例"] = cr;
                        dr["HV"] = Convert.ToDouble(drv["HV"]);
                        double vol = Convert.ToDouble(drv["IV"]) / 100;
                        dr["IV"] = Convert.ToDouble(drv["IV"]);
                        double shares = Convert.ToDouble(drv["IssueNum"]);
                        dr["張數"] = Convert.ToDouble(drv["IssueNum"]);
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
                        dr["交易員"] = drv["TraderID"].ToString();
                        dr["獎勵"] = drv["UseReward"];
                        dr["確認"] = drv["ConfirmChecked"];
                        //if (drv["Reason"] == DBNull.Value)
                        dr["發行原因"] = drv["Reason"] == DBNull.Value ? 0 : Convert.ToInt32(drv["Reason"]);
                        //else
                        //  dr["發行原因"] = Convert.ToInt32(drv["Reason"]);
                        dr["1500W"] = drv["Apply1500W"];
                        dr["標的名稱"] = drv["UnderlyingName"].ToString();
                        double underlyingPrice = 0.0;
                        underlyingPrice = Convert.ToDouble(drv["MPrice"]);
                        dr["股價"] = underlyingPrice;
                        dr["市場"] = drv["Market"].ToString();
                        dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                        double credit = (double) drv["IssueCredit"];
                        credit = Math.Floor(credit);
                        double rewardCredit = (double) drv["RewardIssueCredit"];
                        rewardCredit = Math.Floor(rewardCredit);
                        dr["今日額度"] = credit;
                        dr["獎勵額度"] = rewardCredit;


                        double price = 0.0;
                        double delta = 0.0;
                        if (underlyingPrice != 0) {
                            if (warrantType == "牛熊證")
                                price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                            else if (warrantType == "重設型")
                                price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                            else
                                price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);

                            if (warrantType == "牛熊證")
                                delta = 1.0;
                            else
                                delta = Pricing.Delta(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;

                        }
                        dr["發行價格"] = Math.Round(price, 2);

                        double jumpSize = 0.0;
                        double multiplier = 0.0;
                        if (underlyingID.Substring(0, 2) == "00") {
                            if (underlyingPrice <= 50)
                                multiplier = 0.01;
                            else
                                multiplier = 0.05;
                        } else {
                            if (underlyingPrice <= 10)
                                multiplier = 0.01;
                            else if (underlyingPrice > 10 && underlyingPrice <= 50)
                                multiplier = 0.05;
                            else if (underlyingPrice > 50 && underlyingPrice <= 100)
                                multiplier = 0.1;
                            else if (underlyingPrice > 100 && underlyingPrice <= 500)
                                multiplier = 0.5;
                            else if (underlyingPrice > 500 && underlyingPrice <= 1000)
                                multiplier = 1;
                            else if (underlyingPrice > 1000)
                                multiplier = 5;
                        }
                        jumpSize = delta * multiplier;

                        double vol_ = vol;
                        double price_ = price;
                        double lowerLimit = 0.0;
                        double totalValue = 0.0;
                        totalValue = price_ * shares * 1000;
                        double volLimit = 2 * vol_;
                        while (totalValue < 15000000 && vol_ < volLimit) {
                            vol_ += 0.01;
                            if (warrantType == "牛熊證")
                                price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol_, t, financialR, cr);
                            else if (warrantType == "重設型")
                                price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                            else
                                price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                            totalValue = price_ * shares * 1000;
                        }
                        lowerLimit = price_ - underlyingPrice * 0.1 * cr;
                        lowerLimit = Math.Max(0.01, lowerLimit);

                        dr["IV*"] = vol_ * 100;
                        dr["發行價格*"] = Math.Round(price_, 2);
                        dr["跌停價*"] = Math.Round(lowerLimit, 2);

                        dr["Delta"] = Math.Round(delta, 4);
                        dr["跳動價差"] = Math.Round(jumpSize, 4);


                        dt.Rows.Add(dr);
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private bool CheckReason() {
            bool undoneReason = true;
            string sql2 = @"SELECT [UnderlyingID]";
            sql2 += " FROM [EDIS].[dbo].[ApplyTempList] as A left join Underlying_TraderIssue as B on A.UnderlyingID = B.UID "; //or(A.UnderlyingID = 'IX0001' and B.UID ='TWA00')
            sql2 += " WHERE [UserID]='" + userID + "' AND [ConfirmChecked]='Y' and ((Reason=0  and A.CP='C') or (ReasonP = 0 and A.CP = 'P'))";//

            DataTable noReason = MSSQL.ExecSqlQry(sql2, conn);// new DataTable("noReason");            
            
            foreach (DataRow Row in noReason.Rows) {
                MessageBox.Show(Row["UnderlyingID"] + " 未輸入發行原因");
                undoneReason = false;
            }
           
            return undoneReason;
        }
        private void CheckRelation() {

        }

        private void UpdateData() {
            try {
                SqlCommand cmd = new SqlCommand("DELETE FROM [ApplyTempList] WHERE UserID='" + userID + "'", conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                conn.Close();

                string sql = @"INSERT INTO [ApplyTempList] (SerialNum, UnderlyingID, K, T, R, HV, IV, IssueNum, ResetR, BarrierR, FinancialR, Type, CP, UseReward, ConfirmChecked, Apply1500W, UserID, MDate, TempName, TempType, TraderID, IVNew) ";
                sql += "VALUES(@SerialNum, @UnderlyingID, @K, @T, @R, @HV, @IV, @IssueNum, @ResetR, @BarrierR, @FinancialR, @Type, @CP, @UseReward, @ConfirmChecked, @Apply1500W, @UserID, @MDate, @TempName ,@TempType, @TraderID, @IVNew)";
                List<SqlParameter> ps = new List<SqlParameter>();
                ps.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@K", SqlDbType.Float));
                ps.Add(new SqlParameter("@T", SqlDbType.Int));
                ps.Add(new SqlParameter("@R", SqlDbType.Float));
                ps.Add(new SqlParameter("@HV", SqlDbType.Float));
                ps.Add(new SqlParameter("@IV", SqlDbType.Float));
                ps.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@ResetR", SqlDbType.Float));
                ps.Add(new SqlParameter("@BarrierR", SqlDbType.Float));
                ps.Add(new SqlParameter("@FinancialR", SqlDbType.Float));
                ps.Add(new SqlParameter("@Type", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@CP", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@UseReward", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ConfirmChecked", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@Apply1500W", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@UserID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                ps.Add(new SqlParameter("@TempName", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@TempType", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@TraderID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@IVNew", SqlDbType.Float));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sql, ps);



                int i = 1;
                applyCount = 0;
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string underlyingID = r.Cells["標的代號"].Value.ToString();
                    if (underlyingID != "") {
                        string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + i.ToString("0#");
                        string underlyingName = r.Cells["標的名稱"].Value.ToString();
                        double k = r.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價"].Value);
                        int t = r.Cells["期間(月)"].Value == DBNull.Value ? 6 : Convert.ToInt32(r.Cells["期間(月)"].Value);
                        double cr = r.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["行使比例"].Value);
                        double hv = r.Cells["HV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["HV"].Value);
                        double iv = r.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IV"].Value);
                        double issueNum = r.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(r.Cells["張數"].Value);
                        double resetR = r.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["重設比"].Value);
                        double barrierR = r.Cells["界限比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["界限比"].Value);
                        double financialR = r.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["財務費用"].Value);
                        string type = r.Cells["類型"].Value.ToString();
                        if (type != "一般型" && type != "牛熊證" && type != "重設型") {
                            if (type == "2")
                                type = "牛熊證";
                            else if (type == "3")
                                type = "重設型";
                            else
                                type = "一般型";
                        }

                        string cp = r.Cells["CP"].Value.ToString();
                        if (cp != "C" && cp != "P") {
                            if (cp == "2")
                                cp = "P";
                            else
                                cp = "C";
                        }
                        bool isReward = false;
                        isReward = r.Cells["獎勵"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["獎勵"].Value);
                        string useReward = "N";

                        if (isReward)
                            useReward = "Y";
                        else
                            useReward = "N";

                        bool confirmed = false;
                        confirmed = r.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["確認"].Value);
                        string confirmChecked = "N";
                        if (confirmed) {
                            confirmChecked = "Y";
                            applyCount++;
                        } else
                            confirmChecked = "N";

                        List<SqlParameter> reasonL = new List<SqlParameter>();
                        SQLCommandHelper underlyReason;
                        reasonL.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
                        reasonL.Add(new SqlParameter("@Reason", SqlDbType.Int));
                        if (cp == "C") {
                            underlyReason = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, "Update [Underlying_TraderIssue] set Reason = @Reason where UID = @UnderlyingID", reasonL);
                        } else {
                            underlyReason = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, "Update [Underlying_TraderIssue] set ReasonP = @Reason where UID = @UnderlyingID", reasonL);
                        }
                        int reason = r.Cells["發行原因"].Value == DBNull.Value ? 0 : Convert.ToInt32(r.Cells["發行原因"].Value);

                        bool apply1500Wbool = false;
                        apply1500Wbool = r.Cells["1500W"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["1500W"].Value);
                        string apply1500W = "N";
                        if (apply1500Wbool)
                            apply1500W = "Y";
                        else
                            apply1500W = "N";


                        DateTime expiryDate;
                        expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);
                        expiryDate = expiryDate.AddDays(-1);
                        string sqlTemp = "SELECT TOP 1 TradeDate from TradeDate WHERE IsTrade='Y' AND TradeDate >= '" + expiryDate.ToString("yyyy-MM-dd") + "'";
                        DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                        foreach (DataRowView drTemp in dvTemp) {
                            expiryDate = Convert.ToDateTime(drTemp["TradeDate"]);
                        }
                        string expiryMonth = "";
                        int month = expiryDate.Month;
                        if (month >= 10) {
                            if (month == 10)
                                expiryMonth = "A";
                            if (month == 11)
                                expiryMonth = "B";
                            if (month == 12)
                                expiryMonth = "C";
                        } else
                            expiryMonth = month.ToString();

                        string expiryYear = "";
                        expiryYear = expiryDate.AddYears(-1).ToString("yyyy");
                        expiryYear = expiryYear.Substring(expiryYear.Length - 1, 1);

                        string warrantType = "";
                        string tempType = "";

                        if (type == "牛熊證") {
                            if (cp == "P") {
                                warrantType = "熊";
                                tempType = "4";
                            } else {
                                warrantType = "牛";
                                tempType = "3";
                            }
                        } else {
                            if (cp == "P") {
                                warrantType = "售";
                                tempType = "2";
                            } else {
                                warrantType = "購";
                                tempType = "1";
                            }
                        }

                        string tempName = "";
                        tempName = underlyingName + "凱基" + expiryYear + expiryMonth + warrantType;

                        string traderID = "";
                        traderID = r.Cells["交易員"].Value == DBNull.Value ? userID : r.Cells["交易員"].Value.ToString();

                        double ivNew = r.Cells["IV*"].Value == DBNull.Value ? 0.0 : (double) r.Cells["IV*"].Value;

                        h.SetParameterValue("@SerialNum", serialNumber);
                        h.SetParameterValue("@UnderlyingID", underlyingID);
                        h.SetParameterValue("@K", k);
                        h.SetParameterValue("@T", t);
                        h.SetParameterValue("@R", cr);
                        h.SetParameterValue("@HV", hv);
                        h.SetParameterValue("@IV", iv);
                        h.SetParameterValue("@IssueNum", issueNum);
                        h.SetParameterValue("@ResetR", resetR);
                        h.SetParameterValue("@BarrierR", barrierR);
                        h.SetParameterValue("@FinancialR", financialR);
                        h.SetParameterValue("@Type", type);
                        h.SetParameterValue("@CP", cp);
                        h.SetParameterValue("@UseReward", useReward);
                        h.SetParameterValue("@ConfirmChecked", confirmChecked);
                        h.SetParameterValue("@Apply1500W", apply1500W);
                        h.SetParameterValue("@UserID", userID);
                        h.SetParameterValue("@MDate", DateTime.Now);
                        h.SetParameterValue("@TempName", tempName);
                        h.SetParameterValue("@TempType", tempType);
                        h.SetParameterValue("@TraderID", traderID);
                        h.SetParameterValue("@IVNew", ivNew);

                        h.ExecuteCommand();
                        underlyReason.SetParameterValue("@Reason", reason);
                        //if (underlyingID == "IX0001")
                        //   underlyingID = "TWA00";
                        underlyReason.SetParameterValue("@UnderlyingID", underlyingID);
                        underlyReason.ExecuteCommand();
                        i++;
                    }
                }

                h.Dispose();
                GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔發行");
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Log");
                hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 編輯/更新" + (i-1) + "檔發行");
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void OfficiallyApply() {
            try {
                //CheckReason();
                //CheckRelation();

                UpdateData();
                string sql1 = "DELETE FROM [EDIS].[dbo].[ApplyOfficial] WHERE [UserID]='" + userID + "'";
                string sql2 = @"INSERT INTO [EDIS].[dbo].[ApplyOfficial] ([SerialNumber],[UnderlyingID],[K],[T],[R],[HV],[IV],[IssueNum],[ResetR],[BarrierR],[FinancialR],[Type],[CP],[UseReward],[Apply1500W],[TempName],[TraderID],[MDate],UserID, IVNew)
                                SELECT [SerialNum],[UnderlyingID],[K],[T],[R],[HV],[IV],[IssueNum],[ResetR],[BarrierR],[FinancialR],[Type],[CP],[UseReward],[Apply1500W],[TempName],[TraderID],[MDate],UserID, IVNew";
                //sql2 += "'"+userID + "', [MDate]" ;
                sql2 += " FROM [EDIS].[dbo].[ApplyTempList]";
                sql2 += " WHERE [UserID]='" + userID + "' AND [ConfirmChecked]='Y'";

                string sql3 = "DELETE FROM [EDIS].[dbo].[ApplyTotalList] WHERE [UserID]='" + userID + "' AND [ApplyKind]='1'";
                string sql4 = @"INSERT INTO [EDIS].[dbo].[ApplyTotalList] ([ApplyKind],[SerialNum],[Market],[UnderlyingID],[WarrantName],[CR] ,[IssueNum],[EquivalentNum],[Credit],[RewardCredit],[Type],[CP],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT '1',a.SerialNumber, b.Market, a.UnderlyingID, a.TempName, a.R, a.IssueNum, ROUND(a.R*a.IssueNum, 2), b.IssueCredit, b.RewardIssueCredit, a.Type, a.CP, a.UseReward,'N', a.TraderID, GETDATE(), a.UserID
                                FROM [EDIS].[dbo].[ApplyOfficial] a
                                LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID";
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

                string sql5 = "SELECT [SerialNum], [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND UserID='" + userID + "'";
                DataView dv = DeriLib.Util.ExecSqlQry(sql5, GlobalVar.loginSet.edisSqlConnString);

                string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName WHERE SerialNum=@SerialNum";
                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                pars.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

                foreach (DataRowView dr in dv) {
                    //double count = 0.0;
                    string serialNum = dr["SerialNum"].ToString();
                    string warrantName = dr["WarrantName"].ToString();

                    string sqlTemp = "select top (1) WarrantName from (SELECT [WarrantName] FROM [EDIS].[dbo].[WarrantBasic] WHERE SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 1) + "' union ";
                    sqlTemp += " SELECT [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND [SerialNum]<" + serialNum + " AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 1) + "') as tb1 ";
                    sqlTemp += " order by SUBSTRING(WarrantName,len(WarrantName)-1,len(WarrantName)) desc";
                    DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                    int count = 0;
                    if (dvTemp.Count > 0) {
                        string lastWarrantName = dvTemp[0][0].ToString();                        
                        if (!int.TryParse(lastWarrantName.Substring(lastWarrantName.Length - 2, 2), out count))
                            MessageBox.Show("parse failed "+ lastWarrantName);
                    }

                    //if (dvTemp.Count > 0)
                     //   count += dvTemp.Count;

                    /*sqlTemp = "SELECT [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND [SerialNum]<" + serialNum + " AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 1) + "'";
                    dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                    if (dvTemp.Count > 0)
                        count += dvTemp.Count;*/

                    warrantName = warrantName + (count + 1).ToString("0#");

                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@SerialNum", serialNum);
                    h.ExecuteCommand();
                }
                h.Dispose();

                toolStripLabel2.Text = DateTime.Now + "申請" + applyCount + "檔權證發行成功";
                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 申請" + applyCount + "檔權證發行");
                /* string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                 List<SqlParameter> psInfo = new List<SqlParameter>();
                 psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                 psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                 psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                 psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                 SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                 hInfo.SetParameterValue("@MDate", DateTime.Now);
                 hInfo.SetParameterValue("@InformationType", "Info");
                 hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 申請"+applyCount+"檔權證發行");
                 hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                 hInfo.ExecuteCommand();
                 hInfo.Dispose();*/
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton() {
            if (isEdit) {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["履約價"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["期間(月)"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["HV"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["IV"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["重設比"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["界限比"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["財務費用"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["CP"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行原因"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行價格"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["標的名稱"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["股價"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellActivation = Activation.AllowEdit;

                buttonEdit.Visible = false;
                buttonConfirm.Visible = true;
                buttonDelete.Visible = true;
                buttonCancel.Visible = true;
                toolStripButton1.Visible = false;
                toolStripSeparator2.Visible = false;
                toolStripButton2.Visible = false;
                toolStripSeparator3.Visible = false;
                toolStripButton3.Visible = false;

                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].Hidden = true;

                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = true;

                /*
                for (int x = 0; x < 20; x++)
                {
                    ultraGrid1.DisplayLayout.Bands[0].AddNew();
                    ultraGrid1.Rows[x].Cells[1].Value = (x + 1).ToString();
                    //ultraGrid1.Rows[x].Cells[9].Value = DateTime.Now;
                }

                ultraGrid1.ActiveRowScrollRegion.ScrollRowIntoView(ultraGrid1.Rows[0]);
                //ultraGrid1.Rows[0].Selected = true;
                 * */
            } else {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["履約價"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["期間(月)"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["行使比例"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["HV"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["IV"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["重設比"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["界限比"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["財務費用"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["CP"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行價格"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行原因"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["標的名稱"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["股價"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["Delta"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["跳動價差"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["IV*"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["發行價格*"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Bands[0].Columns["跌停價*"].CellActivation = Activation.NoEdit;


                buttonEdit.Visible = true;
                buttonConfirm.Visible = false;
                buttonDelete.Visible = false;
                buttonCancel.Visible = false;
                toolStripButton1.Visible = true;
                toolStripSeparator2.Visible = true;
                toolStripButton2.Visible = true;
                toolStripSeparator3.Visible = true;
                toolStripButton3.Visible = true;

                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].Hidden = false;

                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].Hidden = false;
                ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = false;
            }
        }

        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;

            if (!e.Layout.ValueLists.Exists("MyValueList")) {
                ValueList v;
                v = e.Layout.ValueLists.Add("MyValueList");
                v.ValueListItems.Add(1, "一般型");
                v.ValueListItems.Add(2, "牛熊證");
                v.ValueListItems.Add(3, "重設型");
            }
            e.Layout.Bands[0].Columns["類型"].ValueList = e.Layout.ValueLists["MyValueList"];

            if (!e.Layout.ValueLists.Exists("MyValueList2")) {
                ValueList v2;
                v2 = e.Layout.ValueLists.Add("MyValueList2");
                v2.ValueListItems.Add(1, "C");
                v2.ValueListItems.Add(2, "P");
            }
            e.Layout.Bands[0].Columns["CP"].ValueList = e.Layout.ValueLists["MyValueList2"];

            if (!e.Layout.ValueLists.Exists("MyValueList3")) {
                ValueList v3;
                v3 = e.Layout.ValueLists.Add("MyValueList3");
                v3.ValueListItems.Add("0005986", "0005986");
                v3.ValueListItems.Add("0007643", "0007643");
                v3.ValueListItems.Add("0008570", "0008570");
                v3.ValueListItems.Add("0008730", "0008730");
            }
            e.Layout.Bands[0].Columns["交易員"].ValueList = e.Layout.ValueLists["MyValueList3"];

            if (!e.Layout.ValueLists.Exists("MyValueList4")) {
                ValueList v;
                v = e.Layout.ValueLists.Add("MyValueList4");
                v.ValueListItems.Add(0, " ");
                v.ValueListItems.Add(1, "技術面偏多，股價持續看好，因此發行認購權證吸引投資人。");
                v.ValueListItems.Add(2, "基本面良好，股價具有漲升的條件，因此發行認購權證吸引投資人。");
                v.ValueListItems.Add(3, "營運動能具提升潛力，因此發行認購權證吸引投資人。");
                v.ValueListItems.Add(4, "提供投資人槓桿避險工具");
                v.ValueListItems.Add(5, "持續針對不同的履約條件、存續期間及認購認售等發行新條件，提供投資人更多元投資選擇");
            }
            e.Layout.Bands[0].Columns["發行原因"].ValueList = e.Layout.ValueLists["MyValueList4"];

        }

        private void buttonEdit_Click(object sender, EventArgs e) {
            isEdit = true;
            SetButton();
        }

        private void buttonConfirm_Click(object sender, EventArgs e) {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            //if (!CheckData())
            //   return;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void buttonDelete_Click(object sender, EventArgs e) {
            isEdit = true;

            DialogResult result = MessageBox.Show("將全部刪除，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                SqlCommand cmd = new SqlCommand("DELETE FROM [ApplyTempList] WHERE UserID='" + userID + "'", conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                conn.Close();
            }
            LoadData();
            SetButton();
        }

        private void buttonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            string cp = "C";
            string underlyingID = e.Row.Cells["標的代號"].Value.ToString();
            cp = e.Row.Cells["CP"].Value.ToString();
            string underlyingName = e.Row.Cells["標的名稱"].Value.ToString();
            double price = e.Row.Cells["發行價格"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["發行價格"].Value);
            double price_ = e.Row.Cells["發行價格*"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["發行價格*"].Value);
            double vol_ = e.Row.Cells["IV*"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["IV*"].Value);
            double lowerLimit = e.Row.Cells["跌停價*"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["跌停價*"].Value);
            double strike = e.Row.Cells["履約價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["履約價"].Value);
            double spot = e.Row.Cells["股價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["股價"].Value);
            string type = e.Row.Cells["類型"].Value.ToString();
            string traderID = "NA";
            string issuable = "Y";
            string putIssuable = "Y";
            string toolTip1 = "非本季標的";
            string toolTip2 = "發行檢查=N";
            string toolTip3 = "非此使用者所屬標的";
            string toolTip4 = "此檔Put須告知主管";
            string sqlTemp = "SELECT [TraderID], [Issuable], [PutIssuable] FROM [EDIS].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
            if (dvTemp.Count > 0) {
                foreach (DataRowView drTemp in dvTemp) {
                    traderID = "000" + drTemp["TraderID"].ToString();
                    issuable = drTemp["Issuable"].ToString();
                    putIssuable = drTemp["PutIssuable"].ToString();
                }
            }
            if (underlyingID != "") {
                if (underlyingName == "" & !isEdit) {
                    e.Row.ToolTipText = toolTip1;
                    e.Row.Appearance.ForeColor = Color.Red;
                } else {
                    if (issuable == "N") {
                        e.Row.Cells["標的代號"].ToolTipText = toolTip2;
                        e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
                    }

                    if (cp == "P" && putIssuable == "N") {
                        e.Row.Cells["CP"].Appearance.ForeColor = Color.Red;
                        e.Row.Cells["CP"].ToolTipText = toolTip4;
                    }

                    if (traderID != userID) {
                        e.Row.Appearance.BackColor = Color.LightYellow;
                        e.Row.ToolTipText = toolTip3;
                    }
                }
                if (price != 0.0 && (price <= 0.6 || price > 3.0))
                    e.Row.Cells["發行價格"].Appearance.ForeColor = Color.Red;
                else
                    e.Row.Cells["發行價格"].Appearance.ForeColor = Color.Black;

                //Check for moneyness constraint
                e.Row.Cells["履約價"].Appearance.ForeColor = Color.Black;
                if (type != "牛熊證") {
                    if (cp == "C" && strike / spot >= 1.5 || cp == "P" && strike / spot <= 0.5) {
                        e.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                        e.Row.Cells["履約價"].ToolTipText = "履約價超過價外50%";
                    }
                }

                if (price != 0.0 && (price <= lowerLimit)) {
                    e.Row.Cells["IV*"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["跌停價*"].Appearance.ForeColor = Color.Red;
                } else {
                    e.Row.Cells["IV*"].Appearance.ForeColor = Color.Black;
                    e.Row.Cells["跌停價*"].Appearance.ForeColor = Color.Black;
                }
            }

        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "標的代號" && ultraGrid1.ActiveRow.Cells["標的代號"].Value != DBNull.Value) {
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
                frmIssueCheck.SelectUnderlying(target);
            }

            if (e.Cell.Column.Key == "CP") {
                string target = (string) e.Cell.Row.Cells[0].Value;
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
                frmIssueCheckPut.SelectUnderlying(target);
            }
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e) {
            DialogResult result = MessageBox.Show("刪除此檔，標的:"+ultraGrid1.ActiveRow.Cells["標的代號"].Value+" 履約價:"+ ultraGrid1.ActiveRow.Cells["履約價"].Value + "，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                ultraGrid1.ActiveRow.Delete();
                UpdateData();
            }
            LoadData();
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right) {
                contextMenuStrip1.Show();
            }
        }

        private void ultraGrid1_BeforeRowsDeleted(object sender, BeforeRowsDeletedEventArgs e) {
            e.DisplayPromptMsg = false;
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "標的代號") {
                string underlyingID = "";
                underlyingID = e.Cell.Row.Cells["標的代號"].Value.ToString();
                string underlyingName = "";
                string traderID = "";
                double underlyingPrice = 0.0;
                string sqlTemp = @"SELECT a.[UnderlyingName]
	                                      ,IsNull(IsNull(b.MPrice, IsNull(b.BPrice,b.APrice)),0) MPrice
                                          ,'000'+a.TraderID TraderID
                                      FROM [EDIS].[dbo].[WarrantUnderlying] a
                                      LEFT JOIN [EDIS].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID ";
                sqlTemp += "WHERE  CAST(UnderlyingID as varbinary(100)) = CAST('" + underlyingID + "' as varbinary(100))";
                // sqlTemp += "WHERE UnderlyingID = '" + underlyingID + "'";
                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                foreach (DataRowView drTemp in dvTemp) {
                    underlyingName = drTemp["UnderlyingName"].ToString();
                    traderID = drTemp["TraderID"].ToString();
                    underlyingPrice = Convert.ToDouble(drTemp["MPrice"]);
                }

                e.Cell.Row.Cells["獎勵"].Value = false;
                e.Cell.Row.Cells["1500W"].Value = false;
                e.Cell.Row.Cells["標的名稱"].Value = underlyingName;
                e.Cell.Row.Cells["交易員"].Value = traderID;
                e.Cell.Row.Cells["股價"].Value = underlyingPrice;
            }

            if (e.Cell.Column.Key == "履約價" || e.Cell.Column.Key == "期間(月)" || e.Cell.Column.Key == "行使比例" || e.Cell.Column.Key == "IV" || e.Cell.Column.Key == "重設比" || e.Cell.Column.Key == "財務費用" || e.Cell.Column.Key == "類型" || e.Cell.Column.Key == "CP" || e.Cell.Column.Key == "張數") {
                double price = 0.0;
                double delta = 0.0;
                double jumpSize = 0.0;
                double multiplier = 0.0;

                string underlyingID = e.Cell.Row.Cells["標的代號"].Value == DBNull.Value ? "" : e.Cell.Row.Cells["標的代號"].Value.ToString();
                double underlyingPrice = 0.0;
                underlyingPrice = e.Cell.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["股價"].Value);
                double k = 0.0;
                k = e.Cell.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["履約價"].Value);
                int t = 0;
                t = e.Cell.Row.Cells["期間(月)"].Value == DBNull.Value ? 0 : Convert.ToInt32(e.Cell.Row.Cells["期間(月)"].Value);
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
                double shares = 0.0;
                shares = e.Cell.Row.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(e.Cell.Row.Cells["張數"].Value);
                bool is1500W = false;
                is1500W = e.Cell.Row.Cells["1500W"].Value == DBNull.Value ? false : (bool) e.Cell.Row.Cells["1500W"].Value;

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

                if (underlyingPrice != 0.0 && underlyingID != "") {
                    if (warrantType == "牛熊證")
                        price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                    else if (warrantType == "重設型")
                        price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                    else
                        price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);

                    if (warrantType == "牛熊證")
                        delta = 1.0;
                    else
                        delta = Pricing.Delta(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;

                    if (underlyingID.Substring(0, 2) == "00") {
                        if (underlyingPrice <= 50)
                            multiplier = 0.01;
                        else
                            multiplier = 0.05;
                    } else {
                        if (underlyingPrice <= 10)
                            multiplier = 0.01;
                        else if (underlyingPrice > 10 && underlyingPrice <= 50)
                            multiplier = 0.05;
                        else if (underlyingPrice > 50 && underlyingPrice <= 100)
                            multiplier = 0.1;
                        else if (underlyingPrice > 100 && underlyingPrice <= 500)
                            multiplier = 0.5;
                        else if (underlyingPrice > 500 && underlyingPrice <= 1000)
                            multiplier = 1;
                        else if (underlyingPrice > 1000)
                            multiplier = 5;
                    }
                }

                jumpSize = delta * multiplier;

                e.Cell.Row.Cells["發行價格"].Value = Math.Round(price, 2);
                e.Cell.Row.Cells["Delta"].Value = Math.Round(delta, 4);
                e.Cell.Row.Cells["跳動價差"].Value = Math.Round(jumpSize, 4);

                double vol_ = vol;
                double price_ = price;
                double lowerLimit = 0.0;
                double totalValue = 0.0;
                totalValue = price_ * shares * 1000;
                double volLimit = 2 * vol_;
                while (totalValue < 15000000 && vol_ != 0.0 && price != 0.0 && vol_ < volLimit) {
                    vol_ += 0.01;
                    if (warrantType == "牛熊證")
                        price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol_, t, financialR, cr);
                    else if (warrantType == "重設型")
                        price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                    else
                        price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                    totalValue = price_ * shares * 1000;
                }
                lowerLimit = price_ - underlyingPrice * 0.1 * cr;
                lowerLimit = Math.Max(0.01, lowerLimit);

                e.Cell.Row.Cells["IV*"].Value = vol_ * 100;
                e.Cell.Row.Cells["發行價格*"].Value = Math.Round(price_, 2);
                e.Cell.Row.Cells["跌停價*"].Value = Math.Round(lowerLimit, 2);

            }
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
                        GlobalUtility.LogInfo("Announce", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證發行");
                        /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                        List<SqlParameter> psInfo = new List<SqlParameter>();
                        psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                        psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                        psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                        psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                        SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                        hInfo.SetParameterValue("@MDate", DateTime.Now);
                        hInfo.SetParameterValue("@InformationType", "Announce");
                        hInfo.SetParameterValue("@InformationContent", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證發行");
                        hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                        hInfo.ExecuteCommand();
                        hInfo.Dispose();*/
                    } else
                        LoadData();
                } else {
                    OfficiallyApply();
                    LoadData();
                }
            }
        }

        private void ultraGrid1_CellChange(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "確認" || e.Cell.Column.Key == "1500W" || e.Cell.Column.Key == "獎勵")
                ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
        }

        private void ultraGrid1_DoubleClickHeader(object sender, DoubleClickHeaderEventArgs e) {
            if (e.Header.Column.Key == "確認") {
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    r.Cells["確認"].Value = true;
                }
                UpdateData();
                LoadData();
            }
        }

        private void FrmApply_FormClosed(object sender, FormClosedEventArgs e) {
            //UpdateData();
        }

        private void ultraGrid1_AfterRowInsert(object sender, RowEventArgs e) {
            //UpdateData();
        }

        private void toolStripButton2_Click(object sender, EventArgs e) {
            LoadData();
        }

        private void toolStripButton3_Click(object sender, EventArgs e) {
            UpdateData();
            LoadData();
        }


    }
}
