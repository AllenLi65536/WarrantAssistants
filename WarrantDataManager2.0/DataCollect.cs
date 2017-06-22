using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using EDLib.SQL;

namespace WarrantDataManager2._0
{
    public class DataCollect
    {
        public CMADODB5.CMConnection cn = new CMADODB5.CMConnection();
        public string arg = "5"; //%
        public string srvLocation = "10.60.0.191";
        public string cnPort = "";
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);

        public void updateWarrantUnderlying() {
            deleteWarrantUnderlying();
            insertWarrantUnderlying();
        }

        public void updateWarrantBasic() {
            deleteWarrantBasic();
            insertWarrantBasic();
        }

        public void updateWarrantUnderlyingCredit() {
            deleteWarrantUnderlyingCredit();
            insertWarrantUnderlyingCredit();
        }

        public void updateWarrantPrices() {
            deleteWarrantPrices();
            insertWarrantPrices();
        }

        public void updateWarrantUnderlyingSummary() {
            deleteWarrantUnderlyingSummary();
            insertWarrantUnderlyingSummary();
        }

        public void updateApplyLists() {
            deleteApplyLists();
        }

        private void deleteWarrantUnderlying() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantUnderlying]", conn);
        }

        private void insertWarrantUnderlying() {
            try {
                conn.Open();
                //更新可發行標的代號，標的名稱，交易員代號，交易員名稱，標的全名

                MSSQL.ExecSqlCmd(@"INSERT INTO [EDIS].[dbo].[WarrantUnderlying] (UnderlyingID, UnderlyingIDCMoney, UnderlyingName, TraderID, TraderName, StockType, FullName) 
                                   SELECT a.[WRTCAN_STKID], a.[WRTCAN_CMONEY_ID], b.[FLGDAT_FLGDTA], ISNULL(c.[TraderAccount],'7643'), ISNULL(c.[TraderName],'Aaron'), a.[WRTCAN_STOCKTYPE], a.[WRTCAN_FULL_NAME] 
                                   FROM [10.7.0.52].[WAFT].[dbo].[V_CANDIDATE] a 
                                   LEFT JOIN [10.7.0.52].[WAFT].[dbo].[V_FLAGDATA_STOCK_UNDERLYING_NAME_LIST] b ON a.[WRTCAN_STKID]=b.[FLGDAT_FLGVAR] 
                                   LEFT JOIN [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] c ON a.[WRTCAN_STKID]=c.UID COLLATE Chinese_Taiwan_Stroke_CI_AS
                                   WHERE a.[WRTCAN_CAN_ISSUE]='1'", conn);
                // LEFT JOIN [10.10.1.30].[EDIS].[dbo].[Underlying_TraderIssue] c ON a.[WRTCAN_STKID]=c.UID COLLATE Chinese_Taiwan_Stroke_CI_AS


                //先預設市場是TSE，以免有些比對不到
                MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[WarrantUnderlying] SET [Market]='TSE'", conn);

                //先從權證系統找市場                
                MSSQL.ExecSqlCmd(@"UPDATE [EDIS].[dbo].[WarrantUnderlying] 
                                   SET [Market]=substring(B.[ISUQTA_MKTTYPE],4,3) 
                                   FROM [10.7.0.52].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] B 
                                   WHERE [UnderlyingID]=B.[ISUQTA_STKID] COLLATE Chinese_Taiwan_Stroke_CI_AS AND B.[ISUQTA_DATE]=(SELECT MAX([ISUQTA_DATE]) FROM [10.7.0.52].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA])", conn);
                conn.Close();

                string sql = "SELECT [股票代號], isNull([上市上櫃],'1') 市場, IsNull([統一編號], '00000000') 統一編號 FROM [上市櫃公司基本資料] WHERE ";
                //DataView dv = DeriLib.Util.ExecSqlQry("SELECT [UnderlyingIDCMoney] FROM [WarrantUnderlying] ORDER BY [UnderlyingIDCMoney]", GlobalVar.loginSet.edisSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry("SELECT [UnderlyingIDCMoney] FROM [WarrantUnderlying] ORDER BY [UnderlyingIDCMoney]", GlobalVar.loginSet.edisSqlConnString);

                string cStr = "";
                foreach (DataRow dr in dv.Rows)
                    cStr += "'" + dr["UnderlyingIDCMoney"].ToString() + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                string cmdText = "UPDATE [WarrantUnderlying] SET UnifiedID=@UnifiedID WHERE UnderlyingIDCMoney=@UnderlyingIDCMoney";
                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                pars.Add(new SqlParameter("@UnderlyingIDCMoney", SqlDbType.VarChar));
                //pars.Add(new SqlParameter("@Market", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@UnifiedID", SqlDbType.VarChar));
                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

                for (; !rs.EOF; rs.MoveNext()) {
                    string commodityIDCMoney = rs.Fields["股票代號"].Value;
                    /*
                    string marketN = rs.Fields["市場"].Value;
                    string market="";
                    if (marketN=="1")
                        market="TSE";
                    else if (marketN=="2")
                        market="OTC";
                    else
                        market="";
                    */
                    string unifiedID = rs.Fields["統一編號"].Value;

                    h.SetParameterValue("@UnderlyingIDCMoney", commodityIDCMoney);
                    //h.SetParameterValue("@Market", market);
                    h.SetParameterValue("@UnifiedID", unifiedID);

                    h.ExecuteCommand();
                }
                h.Dispose();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void deleteWarrantBasic() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantBasic]", conn);
        }

        private void insertWarrantBasic() {
            MSSQL.ExecSqlCmd(@"INSERT INTO [EDIS].[dbo].[WarrantBasic]
                               SELECT a.wid
                                     ,a.wname
                                     ,a.stkid
                                     ,a.stkname
                                     ,a.type 
                                     ,a.strike_now 
                                     ,a.duration
                                     ,IsNull(b.CR,0) CR
                                     ,a.HedgeVol
                                     ,a.IssueVol
                                     ,a.ResetExPrice1*100
                                     ,a.uplimitpercent
                                     ,a.issueprice
                                     ,a.MktTyp 
                                     ,a.TraderID
                                     ,a.Is_Reward_Credit
                                     ,a.issuedate 
                                     ,a.marketdate 
                                     ,a.maturitydate 
                                     ,a.SELF_INCREASE_NUM_TOTAL  
                                     ,a.ISSUE_NUM_WHEN_ISSUE
                                FROM [10.19.1.20].[EDIS].[dbo].[WARRANTS] a
                                LEFT JOIN [10.19.1.20].[EDIS].[dbo].[WarrantBasics] b on a.wname=b.WName and b.MaturityDate >= DATEADD(month,-3,getdate())
                                WHERE a.ISSUECOMNAME='凱基' and a.maturitydate >= DATEADD(month,-3,getdate())
                                ORDER BY a.issuedate DESC, a.wid", conn);
        }

        private void deleteWarrantUnderlyingCredit() {
            conn.Open();
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantUnderlyingCredit]", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantReward]", conn);
            conn.Close();
        }

        private void insertWarrantUnderlyingCredit() {
            string sql = @"INSERT INTO [WarrantReward]
                           SELECT UnderlyingId, SUM([exeRatio]*([FurthurIssueNum]/1000+[IssueNum]/1000)), COUNT(WarrantID)
                           FROM [EDIS].[dbo].[WarrantBasic]
                           WHERE isReward='1' AND IssueDate > ";

            /*DateTime dt = DateTime.Now;
            DateTime startQuarter = dt.AddMonths(0 - (dt.Month - 1) % 3).AddDays(1 - dt.Day);
            string startQuarterDate = startQuarter.ToString("yyyy-MM-dd");*/
            sql += "'" + GlobalVar.globalParameter.firstTradeDateQ.ToString("yyyyMMdd") + "'";
            sql += " GROUP BY UnderlyingID, isReward;";

            conn.Open();
            MSSQL.ExecSqlCmd(@"INSERT INTO EDIS.dbo.WarrantUnderlyingCredit (UnderlyingID, MDate, DataDate, Market, AvailableShares, IssuedPercent, CanIssue, CanFurthurIssue)
                                            SELECT
                                                QUOTA.ISUQTA_STKID, QUOTA.ISUQTA_CREATME, QUOTA.ISUQTA_DATE, SUBSTRING(QUOTA.ISUQTA_MKTTYPE,4,3), (QUOTA.ISUQTA_FOR_WARRANT_SHARES/1000), QUOTA.ISUQTA_ISSUED_PERCENT,
                                                (CANDI.尚可發行額度- QUOTA.ISUQTA_ISSUED_PERCENT) / 100.0 * QUOTA.ISUQTA_FOR_WARRANT_SHARES / 1000.0,
                                                (CANDI.增額發行額度- QUOTA.ISUQTA_ISSUED_PERCENT) / 100.0 * QUOTA.ISUQTA_FOR_WARRANT_SHARES / 1000.0
                                            FROM 
                                            [10.7.0.52].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] AS QUOTA,
                                            (
                                            SELECT
                                                WRTCAN_STKID,
                                                CASE WHEN WRTCAN_STOCKTYPE = 'DE' THEN 100 ELSE 22 END AS 尚可發行額度,
                                                CASE WHEN WRTCAN_STOCKTYPE = 'DE' THEN 100 ELSE 30 END AS 增額發行額度
                                                FROM [10.7.0.52].[WAFT].[dbo].[V_CANDIDATE]
                                                ) AS CANDI
                                            WHERE QUOTA.ISUQTA_DATE= ( SELECT MAX(ISUQTA_DATE) FROM [10.7.0.52].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] )
                                            AND QUOTA.ISUQTA_STKID = CANDI.WRTCAN_STKID", conn);
            MSSQL.ExecSqlCmd(sql, conn);
            conn.Close();
        }

        private void deleteWarrantPrices() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantPrices]", conn);
        }

        private void insertWarrantPrices() {
            MSSQL.ExecSqlCmd(@"INSERT INTO EDIS.dbo.WarrantPrices 
                               SELECT CASE WHEN (A.[CommodityId]='1000') THEN 'IX0001' ELSE A.[CommodityId] END
                                             ,isnull(A.[LastPrice],0)
                                             ,A.[tradedate]
                                             ,isnull(B.[BuyPriceBest1],0)
                                             ,isnull(B.[SellPriceBest1],0)
                                             ,B.[MDate]
                               FROM [10.60.0.37].[TsQuote].[dbo].[vwprice2] A
                               LEFT JOIN [10.60.0.37].[TsQuote].[dbo].[PBest5] B ON A.CommodityId=B.CommodityId", conn);
        }

        private void deleteWarrantUnderlyingSummary() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantUnderlyingSummary]", conn);
        }

        private void insertWarrantUnderlyingSummary() {
            //更新標的代號，標的名稱，交易員代號，市場，額度，累計損益
            /*SqlCommand cmd = new SqlCommand(@"INSERT INTO EDIS.dbo.WarrantUnderlyingSummary (UnderlyingID, UnderlyingName, TraderID, Market, PutIssuable, IssueCredit, IssuedPercent, AccNetIncome)
                                              SELECT a.[UnderlyingID], a.[UnderlyingName], a.[TraderID], a.[Market], IsNull(c.CanIssuePut,'Y'), b.CanIssue, b.IssuedPercent, IsNull(c.AccNetIncome,0)
                                              FROM [EDIS].[dbo].[WarrantUnderlying] a
                                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingCredit] b on a.UnderlyingID=b.UnderlyingID
                                              LEFT JOIN [EDIS].[dbo].[WarrantIssueCheck] c on a.UnderlyingID=c.UnderlyingID", conn);
            /*SqlCommand cmd = new SqlCommand(@"Update EDIS.dbo.WarrantUnderlyingSummary  set UnderlyingID=i.UnderlyingID , UnderlyingName=i.[UnderlyingName], TraderID = i.[TraderID], Market= i.[Market], PutIssuable= i.canIssueP, IssueCredit=i.canIssue, IssuedPercent=i.IssuedPercent, AccNetIncome=i.accNI
                                               from (SELECT a.[UnderlyingID], a.[UnderlyingName], a.[TraderID], a.[Market], IsNull(c.CanIssuePut,'Y') canIssueP, b.CanIssue canIssue, b.IssuedPercent, IsNull(c.AccNetIncome,0) accNI
                                              FROM [EDIS].[dbo].[WarrantUnderlying] a
                                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingCredit] b on a.UnderlyingID=b.UnderlyingID
                                              LEFT JOIN [EDIS].[dbo].[WarrantIssueCheck] c on a.UnderlyingID=c.UnderlyingID) i where i.UnderlyingID =WarrantUnderlyingSummary.UnderlyingID ", conn);*/

            conn.Open();
            MSSQL.ExecSqlCmd(@"INSERT INTO EDIS.dbo.WarrantUnderlyingSummary (UnderlyingID, UnderlyingName, TraderID, Market, PutIssuable, IssueCredit, IssuedPercent, AccNetIncome)
                                              SELECT a.[UnderlyingID], a.[UnderlyingName], a.[TraderID], a.[Market], IsNull(c.CanIssuePut,'Y'), b.CanIssue, b.IssuedPercent, IsNull(c.AccNetIncome,0)
                                              FROM [EDIS].[dbo].[WarrantUnderlying] a
                                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingCredit] b on a.UnderlyingID=b.UnderlyingID
                                              LEFT JOIN [EDIS].[dbo].[WarrantIssueCheck] c on a.UnderlyingID=c.UnderlyingID", conn);

            //更新獎勵額度
            string sql = @"SELECT a.[UnderlyingID], a.[AvailableShares], IsNull(b.[UsedRewardNum],0) UsedRewardNum
                           FROM [EDIS].[dbo].[WarrantUnderlyingCredit] a
                           LEFT JOIN [EDIS].[dbo].[WarrantReward] b on a.UnderlyingID=b.UnderlyingID
                           ORDER BY [UnderlyingID]";
            //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
            DataTable dv = MSSQL.ExecSqlQry(sql, conn);

            string cmdText = "UPDATE [WarrantUnderlyingSummary] SET RewardIssueCredit=@RewardIssueCredit WHERE UnderlyingID=@UnderlyingID";
            List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
            pars.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
            pars.Add(new SqlParameter("@RewardIssueCredit", SqlDbType.Float));
            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

            foreach (DataRow dr in dv.Rows) {
                string underlyingID = dr["UnderlyingID"].ToString();
                double availableShares = Convert.ToDouble(dr["AvailableShares"]);
                double used = Convert.ToDouble(dr["UsedRewardNum"]);
                double remainCredit = 0.0;
                //若本季為A級券商
                if (GlobalVar.globalParameter.isLevelA == true)
                    remainCredit = availableShares * GlobalVar.globalParameter.givenRewardPercent - used;

                h.SetParameterValue("@UnderlyingID", underlyingID);
                h.SetParameterValue("@RewardIssueCredit", remainCredit);

                h.ExecuteCommand();
            }
            h.Dispose();

            //更新是否可發行
            //先預設都可以發行            
            MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[WarrantUnderlyingSummary] SET [Issuable]='Y'", conn);

            //從WarrantIssueCheck比對
            string sql2 = @"SELECT [UnderlyingID]
                                  ,IsNull([CashDividendDate],'2030-12-31') CashDividendDate
                                  ,IsNull([StockDividendDate],'2030-12-31') StockDividendDate
                                  ,IsNull([PublicOfferingDate],'2030-12-31') PublicOfferingDate
                                  ,IsNull([DisposeEndDate],'1990-12-31') DisposeEndDate
                                  ,[WatchCount]
                                  ,[WarningScore]
                              FROM [EDIS].[dbo].[WarrantIssueCheck]";
            //DataView dv2 = DeriLib.Util.ExecSqlQry(sql2, GlobalVar.loginSet.edisSqlConnString);
            DataTable dv2 = MSSQL.ExecSqlQry(sql2, conn);
            conn.Close();

            string cmdText2 = "UPDATE [WarrantUnderlyingSummary] SET Issuable=@Issuable WHERE UnderlyingID=@UnderlyingID";
            List<System.Data.SqlClient.SqlParameter> pars2 = new List<System.Data.SqlClient.SqlParameter>();
            pars2.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
            pars2.Add(new SqlParameter("@Issuable", SqlDbType.VarChar));
            SQLCommandHelper h2 = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText2, pars2);

            foreach (DataRow dr2 in dv2.Rows) {
                bool issuable = true;
                DateTime applyDate = DateTime.Today;
                DateTime issueDate = GlobalVar.globalParameter.nextTradeDate1;
                string underlyingID = dr2["UnderlyingID"].ToString();
                DateTime cashDividendDate = Convert.ToDateTime(dr2["CashDividendDate"]);
                DateTime stockDividendDate = Convert.ToDateTime(dr2["StockDividendDate"]);
                DateTime publicOfferingDate = Convert.ToDateTime(dr2["PublicOfferingDate"]);
                DateTime disposeEndDate = Convert.ToDateTime(dr2["DisposeEndDate"]);
                int watchCount = Convert.ToInt32(dr2["WatchCount"]);
                int warningScore = Convert.ToInt32(dr2["WarningScore"]);
                if (cashDividendDate == issueDate)
                    issuable = false;
                else if (stockDividendDate == issueDate)
                    issuable = false;
                else if (publicOfferingDate == issueDate)
                    issuable = false;
                else if (disposeEndDate.AddMonths(3) > applyDate)
                    issuable = false;
                else if (watchCount >= 2)
                    issuable = false;
                else if (warningScore > 0)
                    issuable = false;
                else
                    issuable = true;

                string issuablesString = "Y";
                if (issuable == false)
                    issuablesString = "N";

                h2.SetParameterValue("@UnderlyingID", underlyingID);
                h2.SetParameterValue("@Issuable", issuablesString);

                h2.ExecuteCommand();
            }
            h2.Dispose();
        }

        private void deleteApplyLists() {
            conn.Open();
            MSSQL.ExecSqlCmd("DELETE FROM [ApplyOfficial]", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [ReIssueOfficial]", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [ApplyTotalList]", conn);
            MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[ApplyTempList] SET ConfirmChecked='N'", conn);
            MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[ReIssueTempList] SET ConfirmChecked='N'", conn);
            conn.Close();
        }
    }
}
