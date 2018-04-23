using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using EDLib.SQL;
using System.Linq;

namespace WarrantDataManager
{
    public static class DataCollect
    {
        public static CMADODB5.CMConnection cn = new CMADODB5.CMConnection();
        public static string arg = "5"; //%
        public static string srvLocation = "10.60.0.191";
        public static string cnPort = "";
        public static SqlConnection conn = new SqlConnection(LoginSet.edisSqlConnString);

        public static WorkState UpdateWarrantUnderlying() {
            try {
                DeleteWarrantUnderlying();
                InsertWarrantUnderlying();
                return WorkState.Successful;
            } catch (Exception) {
                return WorkState.Exception;
            }
        }

        public static WorkState UpdateWarrantBasic() {
            try {
                DeleteWarrantBasic();
                InsertWarrantBasic();
                return WorkState.Successful;
            } catch (Exception) {
                return WorkState.Exception;
            }
        }

        public static WorkState UpdateWarrantUnderlyingCredit() {
            try {
                DeleteWarrantUnderlyingCredit();
                InsertWarrantUnderlyingCredit();
                return WorkState.Successful;
            } catch (Exception) {
                return WorkState.Exception;
            }
        }

        public static WorkState UpdateWarrantPrices() {
            try {
                DeleteWarrantPrices();
                InsertWarrantPrices();
                return WorkState.Successful;
            } catch (Exception) {
                return WorkState.Exception;
            }
        }

        public static WorkState UpdateWarrantUnderlyingSummary() {
            try {
                DeleteWarrantUnderlyingSummary();
                InsertWarrantUnderlyingSummary();
                return WorkState.Successful;
            } catch (Exception e) {
                MessageBox.Show(e.Message);
                return WorkState.Exception;
            }
        }

        public static WorkState UpdateApplyLists() {
            try {
                DeleteApplyLists();
                return WorkState.Successful;
            } catch (Exception) {
                return WorkState.Exception;
            }
        }

        private static void DeleteWarrantUnderlying() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantUnderlying]", conn);
        }

        private static void InsertWarrantUnderlying() {
            try {
                conn.Open();
                //更新可發行標的代號，標的名稱，交易員代號，交易員名稱，標的全名
                /*MSSQL.ExecSqlCmd(@"INSERT INTO [EDIS].[dbo].[WarrantUnderlying] (UnderlyingID, UnderlyingIDCMoney, UnderlyingName, TraderID, TraderName, StockType, FullName) 
                                   SELECT a.[WRTCAN_STKID], a.[WRTCAN_CMONEY_ID], b.[FLGDAT_FLGDTA], ISNULL(c.[TraderAccount],'7643'), ISNULL(c.[TraderName],'Aaron'), a.[WRTCAN_STOCKTYPE], a.[WRTCAN_FULL_NAME] 
                                   FROM [10.100.10.131].[WAFT].[dbo].[V_CANDIDATE] a 
                                   LEFT JOIN [10.100.10.131].[WAFT].[dbo].[V_FLAGDATA_STOCK_UNDERLYING_NAME_LIST] b ON a.[WRTCAN_STKID]=b.[FLGDAT_FLGVAR] 
                                   LEFT JOIN [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] c ON a.[WRTCAN_STKID]=c.UID COLLATE Chinese_Taiwan_Stroke_CI_AS
                                   WHERE a.[WRTCAN_CAN_ISSUE]='1'", conn);*/
                // LEFT JOIN [10.10.1.30].[EDIS].[dbo].[Underlying_TraderIssue] c ON a.[WRTCAN_STKID]=c.UID COLLATE Chinese_Taiwan_Stroke_CI_AS
                MSSQL.ExecSqlCmd(@"INSERT INTO [EDIS].[dbo].[WarrantUnderlying] (UnderlyingID, UnderlyingIDCMoney, UnderlyingName, TraderID, TraderName, StockType, FullName) 
select C.WRTCAN_STKID, C.WRTCAN_CMONEY_ID, C.WRTCAN_SHORT_NAME, C.TraderAccount, C.TraderName, C.WRTCAN_STOCKTYPE, C.WRTCAN_FULL_NAME from 
(SELECT A.WRTCAN_STKID, A.WRTCAN_CMONEY_ID, A.WRTCAN_SHORT_NAME, ISNULL(B.TraderAccount,'7643') as TraderAccount, ISNULL(B.TraderName,'Aaron') as TraderName, A.WRTCAN_STOCKTYPE, A.WRTCAN_FULL_NAME,    
    CASE WHEN (WRTCAN_STOCKTYPE = 'DI' OR WRTCAN_STOCKTYPE = 'DE') AND (AUT.FLGDAT_FLGVAR is null OR AUT.FLGDAT_FLGVAR = 0 OR AUT.FLGDAT_FLGVAR < CONVERT(VARCHAR, GETDATE(), 112)) THEN '未授權'                       
    WHEN WRTCAN_STOCKTYPE = 'DS' AND A.WRTCAN_STKID IN('2883', '6005') THEN '未授權'
    WHEN (WRTCAN_STOCKTYPE = 'DS' OR WRTCAN_STOCKTYPE = 'DR') AND A.WRTCAN_SOURCE = 'STOCK_A' AND (C.FLGDAT_FLGVAR <> 'A' OR C.FLGDAT_FLGVAR is null) THEN '非A級券商'               
    ELSE '1'
    END as CHECK_CAN_ISSUE
FROM [10.100.10.131].[WAFT].[dbo].[CANDIDATE] as A WITH(NOLOCK)
LEFT JOIN  [10.100.10.131].EDAISYS.dbo.FLAGDATAS as AUT WITH(NOLOCK)
    ON WRTCAN_INSNBR = AUT.FLGDAT_FLGNBR AND AUT.FLGDAT_FLGNAM = 'WRT_AUTHORIZATION_MAINTAIN' AND AUT.FLGDAT_FLGNBR = A.WRTCAN_INSNBR 
LEFT JOIN  [10.100.10.131].EDAISYS.dbo.FLAGDATAS as C WITH (NOLOCK)
     ON C.FLGDAT_FLGNAM = 'WRT_MARKET_RATING' and  convert(varchar(10), GETDATE(), 112) between C.FLGDAT_FLGNBR and C.FLGDAT_ORDERS
LEFT JOIN [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] as B ON A.[WRTCAN_STKID]=B.UID COLLATE Chinese_Taiwan_Stroke_CI_AS
WHERE A.WRTCAN_DATE = ( SELECT MAX(WRTCAN_DATE) FROM  [10.100.10.131].WAFT.dbo.CANDIDATE WHERE WRTCAN_VER = 1) AND A.WRTCAN_VER = 1)as C
WHERE C.CHECK_CAN_ISSUE = '1'", conn);


                //先預設市場是TSE，以免有些比對不到
                MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[WarrantUnderlying] SET [Market]='TSE'", conn);

                //先從權證系統找市場                
                MSSQL.ExecSqlCmd(@"UPDATE [EDIS].[dbo].[WarrantUnderlying] 
                                   SET [Market]=substring(B.[ISUQTA_MKTTYPE],4,3) 
                                   FROM [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] B 
                                   WHERE [UnderlyingID]=B.[ISUQTA_STKID] COLLATE Chinese_Taiwan_Stroke_CI_AS AND B.[ISUQTA_DATE]=(SELECT MAX([ISUQTA_DATE]) FROM [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA])", conn);
                conn.Close();

                string sql = "SELECT [股票代號], isNull([上市上櫃],'1') 市場, IsNull([統一編號], '00000000') 統一編號 FROM [上市櫃公司基本資料] WHERE ";
                //DataView dv = DeriLib.Util.ExecSqlQry("SELECT [UnderlyingIDCMoney] FROM [WarrantUnderlying] ORDER BY [UnderlyingIDCMoney]", LoginSet.edisSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry("SELECT [UnderlyingIDCMoney] FROM [WarrantUnderlying] ORDER BY [UnderlyingIDCMoney]", LoginSet.edisSqlConnString);

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
                SQLCommandHelper h = new SQLCommandHelper(LoginSet.edisSqlConnString, cmdText, pars);

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

        private static void DeleteWarrantBasic() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantBasic]", conn);
        }

        private static void InsertWarrantBasic() {
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

        private static void DeleteWarrantUnderlyingCredit() {
            conn.Open();
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantUnderlyingCredit]", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantReward]", conn);
            conn.Close();
        }

        private static void InsertWarrantUnderlyingCredit() {
            string sql = @"INSERT INTO [WarrantReward]"
        + " select ID, sum(sum1), sum(count1)"
        + " from ((SELECT UnderlyingId as ID, SUM([exeRatio]*([IssueNum]/1000)) as sum1, COUNT(WarrantID) as count1"
                + " FROM [EDIS].[dbo].[WarrantBasic] "
                + $" WHERE isReward='1' AND IssueDate > '{GlobalVar.globalParameter.firstTradeDateQ.ToString("yyyyMMdd")}'"
                + " GROUP BY UnderlyingId)"
            + " UNION "
                + " (Select UnderlyingID as ID, Sum(ISNULL(RewardQuotaUsed, 0)) as sum1, count(RewardQuotaUsed) as count1"
                + " FROM [EDIS].[dbo].[ReIssueReward]"
                + $" WHERE MDate >= '{GlobalVar.globalParameter.firstTradeDateQ.ToString("yyyyMMdd")}'"
                + " GROUP BY UnderlyingID)) as A"
        + " Group By ID;";

            /*string sql = @"INSERT INTO [WarrantReward]
                           SELECT UnderlyingId, SUM([exeRatio]*([FurthurIssueNum]/1000+[IssueNum]/1000)), COUNT(WarrantID)
                           FROM [EDIS].[dbo].[WarrantBasic]
                           WHERE isReward='1' AND IssueDate > "
                      + "'" + GlobalVar.globalParameter.firstTradeDateQ.ToString("yyyyMMdd") + "'"
                      + " GROUP BY UnderlyingID, isReward;";*/
            /*DateTime dt = DateTime.Now;
            DateTime startQuarter = dt.AddMonths(0 - (dt.Month - 1) % 3).AddDays(1 - dt.Day);
            string startQuarterDate = startQuarter.ToString("yyyy-MM-dd");*/

            conn.Open();
            MSSQL.ExecSqlCmd(@"INSERT INTO EDIS.dbo.WarrantUnderlyingCredit (UnderlyingID, MDate, DataDate, Market, AvailableShares, IssuedPercent, CanIssue, CanFurthurIssue, CanIssueDelta)
                               SELECT distinct
                                A.ISUQTA_STKID, A.ISUQTA_CREATME, A.ISUQTA_DATE, SUBSTRING(A.ISUQTA_MKTTYPE,4,3), (A.ISUQTA_FOR_WARRANT_SHARES/1000), A.ISUQTA_ISSUED_PERCENT,
                                (B.CanIssue- A.ISUQTA_ISSUED_PERCENT) / 100.0 * A.ISUQTA_FOR_WARRANT_SHARES / 1000.0,
                                (B.CanFurthurIssue- A.ISUQTA_ISSUED_PERCENT) / 100.0 * A.ISUQTA_FOR_WARRANT_SHARES / 1000.0
								, (B.CanIssue- A.ISUQTA_ISSUED_PERCENT) / 100.0 * A.ISUQTA_FOR_WARRANT_SHARES / 1000.0 - (B.CanIssue- Q1.ISUQTA_ISSUED_PERCENT) / 100.0 * Q1.ISUQTA_FOR_WARRANT_SHARES / 1000.0
                               from [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] as A, 
							    [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] as Q1, 						  
								(select WRTCAN_STKID,
                                    case when WRTCAN_STOCKTYPE = 'DE' then 100 else 22 end as CanIssue,
                                    case when WRTCAN_STOCKTYPE = 'DE' then 100 else 30 end as CanFurthurIssue
									FROM [10.100.10.131].[WAFT].[dbo].[CANDIDATE]
									where WRTCAN_DATE = (select max(WRTCAN_DATE) from [10.100.10.131].[WAFT].[dbo].[CANDIDATE])) AS B
                                where A.ISUQTA_DATE = (select MAX(ISUQTA_DATE) from [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] )
								and Q1.ISUQTA_DATE = (select MAX(ISUQTA_DATE) from [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA]  where ISUQTA_DATE < (select MAX(ISUQTA_DATE) from [10.100.10.131].[EXTSRC].[dbo].[V_WRT_ISSUE_QUOTA] ))
                                and A.ISUQTA_STKID = Q1.ISUQTA_STKID and A.ISUQTA_STKID = B.WRTCAN_STKID
								order by A.ISUQTA_STKID", conn);//V_CANDIDATE
            MSSQL.ExecSqlCmd(sql, conn);
            conn.Close();
        }

        private static void DeleteWarrantPrices() {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantPrices]", conn);
        }

        private static void InsertWarrantPrices() {
            MSSQL.ExecSqlCmd(@"INSERT INTO EDIS.dbo.WarrantPrices 
                               SELECT DISTINCT CASE WHEN (A.[CommodityId]='1000') THEN 'IX0001' ELSE A.[CommodityId] END
                                             ,isnull(A.[LastPrice],0)
                                             ,A.[tradedate]
                                             ,isnull(B.[BuyPriceBest1],0)
                                             ,isnull(B.[SellPriceBest1],0)
                                             ,B.[MDate]
                               FROM [10.60.0.37].[TsQuote].[dbo].[vwprice2] A
                               LEFT JOIN [10.60.0.37].[TsQuote].[dbo].[PBest5] B ON A.CommodityId=B.CommodityId", conn);
        }

        private static void UpdateLastPrices() {
            MSSQL.ExecSqlCmd(@"UPDATE [dbo].[WarrantPrices]
   SET [MPrice] = B.MPrice
      ,[MDateTime] =  GetDate()     
from (SELECT [T730010] as ID  
      ,Case when [T730050] <> 0 then [T730050]
	  else T730060 END as Mprice	 
	  --,[T730030]
      --,[T730040] 	 
  FROM [10.60.0.37].[DeriPosition].[dbo].[PTOS_HHPT73M] 
Union SELECT [T730010]  
      ,Case when [T730050] <> 0 then [T730050]
	  else T730060 END
	  --,[T730030]
      --,[T730040]  	 
  FROM [10.60.0.37].[DeriPosition].[dbo].[PTOS_OCPT73M] ) as B
 WHERE CommodityID = B.ID COLLATE Chinese_Taiwan_Stroke_CI_AS", conn);
        }

        private static void DeleteWarrantUnderlyingSummary() {

        }

        private static void InsertWarrantUnderlyingSummary() {
            //更新標的代號，標的名稱，交易員代號，市場，額度，累計損益
            /*SqlCommand cmd = new SqlCommand(@"Update EDIS.dbo.WarrantUnderlyingSummary  set UnderlyingID=i.UnderlyingID , UnderlyingName=i.[UnderlyingName], TraderID = i.[TraderID], Market= i.[Market], PutIssuable= i.canIssueP,
             *  IssueCredit=i.canIssue, IssuedPercent=i.IssuedPercent, AccNetIncome=i.accNI
                                               from (SELECT a.[UnderlyingID], a.[UnderlyingName], a.[TraderID], a.[Market], IsNull(c.CanIssuePut,'Y') canIssueP, b.CanIssue canIssue, b.IssuedPercent, IsNull(c.AccNetIncome,0) accNI
                                              FROM [EDIS].[dbo].[WarrantUnderlying] a
                                              LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingCredit] b on a.UnderlyingID=b.UnderlyingID
                                              LEFT JOIN [EDIS].[dbo].[WarrantIssueCheck] c on a.UnderlyingID=c.UnderlyingID) i where i.UnderlyingID =WarrantUnderlyingSummary.UnderlyingID ", conn);*/

            conn.Open();
            DataTable notIssuable = MSSQL.ExecSqlQry("Select UnderlyingID, TraderID from EDIS.dbo.WarrantUnderlyingSummary where Issuable='N'", conn);
            DataTable issuableAnnounce = MSSQL.ExecSqlQry($"SELECT [InformationContent], MUser FROM [dbo].[InformationLog] where MDate >= '{DateTime.Today.ToString("yyyyMMdd")}' and InformationType = 'AnnounceIssue' ", conn);

            MSSQL.ExecSqlCmd("DELETE FROM [WarrantUnderlyingSummary]", conn);

            if (GlobalVar.globalParameter.isLevelA)
                MSSQL.ExecSqlCmd("INSERT INTO EDIS.dbo.WarrantUnderlyingSummary (UnderlyingID, UnderlyingName, TraderID, Market, PutIssuable, IssueCredit, IssueCreditDelta, IssuedPercent, AccNetIncome, Issuable, RewardIssueCredit) "
                          + $" SELECT a.[UnderlyingID], a.[UnderlyingName], a.[TraderID], a.[Market], IsNull(c.CanIssuePut,'Y'), Floor(b.CanIssue), b.CanIssueDelta, b.IssuedPercent, IsNull(c.AccNetIncome,0), 'Y', Floor(b.AvailableShares * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(d.[UsedRewardNum],0)) "
                          + " FROM [EDIS].[dbo].[WarrantUnderlying] a "
                          + " LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingCredit] b on a.UnderlyingID=b.UnderlyingID "
                          + " LEFT JOIN [EDIS].[dbo].[WarrantIssueCheck] c on a.UnderlyingID=c.UnderlyingID "
                          + " LEFT JOIN [EDIS].[dbo].[WarrantReward] d on a.UnderlyingID=d.UnderlyingID", conn);
            else
                MSSQL.ExecSqlCmd("INSERT INTO EDIS.dbo.WarrantUnderlyingSummary (UnderlyingID, UnderlyingName, TraderID, Market, PutIssuable, IssueCredit, IssueCreditDelta, IssuedPercent, AccNetIncome, Issuable, RewardIssueCredit) "
                           + " SELECT a.[UnderlyingID], a.[UnderlyingName], a.[TraderID], a.[Market], IsNull(c.CanIssuePut,'Y'), Floor(b.CanIssue), b.CanIssueDelta, b.IssuedPercent, IsNull(c.AccNetIncome,0), 'Y', 0"
                           + " FROM [EDIS].[dbo].[WarrantUnderlying] a "
                           + " LEFT JOIN [EDIS].[dbo].[WarrantUnderlyingCredit] b on a.UnderlyingID=b.UnderlyingID "
                           + " LEFT JOIN [EDIS].[dbo].[WarrantIssueCheck] c on a.UnderlyingID=c.UnderlyingID ", conn);

            //從WarrantIssueCheck比對
            string sql2 = "(SELECT [UnderlyingID] FROM [EDIS].[dbo].[WarrantIssueCheck] "
                            + $" Where IsNull([CashDividendDate],'20301231') = '{GlobalVar.globalParameter.nextTradeDate1.ToString("yyyyMMdd")}' "
                            + $" or IsNull([StockDividendDate],'20301231') = '{GlobalVar.globalParameter.nextTradeDate1.ToString("yyyyMMdd")}' "
                            + $" or IsNull([PublicOfferingDate],'20301231') = '{GlobalVar.globalParameter.nextTradeDate1.ToString("yyyyMMdd")}' "
                            + $" or DATEADD(month, 3, IsNull([DisposeEndDate],'19901231')) > '{DateTime.Today.ToString("yyyyMMdd")}' "
                            + " or WatchCount >= 2 or WarningScore > 0)";

            MSSQL.ExecSqlCmd($"UPDATE [WarrantUnderlyingSummary] SET Issuable='N' WHERE UnderlyingID in {sql2}", conn);

            DataTable issuable = MSSQL.ExecSqlQry("Select UnderlyingID from EDIS.dbo.WarrantUnderlyingSummary where Issuable='Y'", conn);
            var notIssuable2Issuable = from row in notIssuable.AsEnumerable()
                                       from row2 in issuable.AsEnumerable()
                                       where row[0].ToString() == row2[0].ToString()
                                       select row;

            foreach (DataRow row in notIssuable2Issuable.OrderBy(x => x[1].ToString())) {
                MSSQL.ExecSqlCmd($"INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values( GETDATE(), 'Announce', '{row[0].ToString()}可以發行', '{row[1].ToString()}')", conn);
                MSSQL.ExecSqlCmd($"INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values( GETDATE(), 'AnnounceIssue', '{row[0].ToString()}', '{row[1].ToString()}')", conn);
            }

            DataTable notIssuableAnnounce = MSSQL.ExecSqlQry("Select UnderlyingID from EDIS.dbo.WarrantUnderlyingSummary where Issuable='N'", conn);
            var issuable2NotIssuable = from row in issuableAnnounce.AsEnumerable()
                                       from row2 in notIssuableAnnounce.AsEnumerable()
                                       where row[0].ToString() == row2[0].ToString()
                                       select row;

            foreach (var row in issuable2NotIssuable) {//.OrderBy(x => x[1].ToString())
                MSSQL.ExecSqlCmd($"INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values( GETDATE(), 'Log', '{row[0].ToString()}不可發行', '{row[1].ToString()}')", conn);
                MSSQL.ExecSqlCmd($"Delete from [InformationLog] where MDate >='{DateTime.Today.ToString("yyyyMMdd")}' and InformationType ='Announce' and InformationContent ='{row[0].ToString()}可以發行'", conn);
                MSSQL.ExecSqlCmd($"Delete from [InformationLog] where MDate >='{DateTime.Today.ToString("yyyyMMdd")}' and InformationType ='AnnounceIssue' and InformationContent ='{row[0].ToString()}'", conn);
            }

            conn.Close();
        }

        private static void DeleteApplyLists() {
            conn.Open();
            MSSQL.ExecSqlCmd("INSERT INTO [dbo].[ReIssueReward] ([UnderlyingId], [RewardQuotaUsed], [MDate])"
                + " (select UnderlyingID, exeRatio * ReIssueNum, GETDATE() from ReIssueOfficial where UseReward = 'Y')", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [ApplyOfficial]", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [ReIssueOfficial]", conn);
            MSSQL.ExecSqlCmd("DELETE FROM [ApplyTotalList]", conn);
            MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[ApplyTempList] SET ConfirmChecked='N'", conn);
            MSSQL.ExecSqlCmd("UPDATE [EDIS].[dbo].[ReIssueTempList] SET ConfirmChecked='N'", conn);
            conn.Close();
        }
    }
}
