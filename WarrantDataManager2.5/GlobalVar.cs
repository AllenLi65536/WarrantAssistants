using System;
using System.Data;
using EDLib.SQL;
using System.Windows.Forms;

namespace WarrantDataManager
{
    public class GlobalVar
    {
        public static MainForm mainForm;
        public static AutoWork autoWork;
        //public static ErrProcess errProcess;
        public static GlobalParameter globalParameter;
    }

    public class GlobalUtility
    {
        public static void Start() {
            LoadGlobalParameters();
            //GlobalVar.errProcess = new ErrProcess();
            GlobalVar.autoWork = new AutoWork();
            //GlobalVar.warrantPriceUpdator = new WarrantPriceUpdator();
            //GlobalVar.warrantPriceProcess = new WarrantPriceProcess();
        }


        public static void LoadGlobalParameters() {
            if (GlobalVar.globalParameter == null)
                GlobalVar.globalParameter = new GlobalParameter();

            DataTable dt = MSSQL.ExecSqlQry("SELECT [InterestRate],[GivenRewardPercent] FROM [EDIS].[dbo].[Global]", LoginSet.edisSqlConnString);
            GlobalVar.globalParameter.interestRate = Convert.ToDouble(dt.Rows[0]["InterestRate"]);
            GlobalVar.globalParameter.givenRewardPercent = Convert.ToDouble(dt.Rows[0]["GivenRewardPercent"]);
           
            CheckIsLevelA();
            CheckIsTodayTradeDate();
            GetNextTradeDate();
            GetLastTradeDate();
            GetFirstTradeDateOfQuarter();
        }

        private static void CheckIsLevelA() {
            DataTable isA = MSSQL.ExecSqlQry(@"select top 1 FLGDAT_FLGVAR from EDAISYS.dbo.FLAGDATAS 
                                            where FLGDAT_FLGNAM = 'WRT_MARKET_RATING'
                                            order by FLGDAT_ORDERS desc",
                                            LoginSet.edaisysConnString);
            if (isA.Rows.Count > 0 && isA.Rows[0][0].ToString() == "A") {
                MSSQL.ExecSqlCmd("Update [Global] Set IsLevelA = 1", LoginSet.edisSqlConnString);
                GlobalVar.globalParameter.isLevelA = true;
            } else {
                MSSQL.ExecSqlCmd("Update [Global] Set IsLevelA = 0", LoginSet.edisSqlConnString);
                GlobalVar.globalParameter.isLevelA = false;
            }
            //MessageBox.Show(GlobalVar.globalParameter.isLevelA.ToString());
        }
        private static void CheckIsTodayTradeDate() {
            DataTable dv = MSSQL.ExecSqlQry("SELECT IsTrade FROM [TradeDate] WHERE CONVERT(VARCHAR, TradeDate, 112) = CONVERT(VARCHAR, GETDATE(), 112)", LoginSet.tsquoteSqlConnString);
            if (dv.Rows[0]["IsTrade"].ToString() == "Y")
                GlobalVar.globalParameter.isTodayTradeDate = true;
            else
                GlobalVar.globalParameter.isTodayTradeDate = false;
        }

        private static void GetLastTradeDate() {
            try {
                string sql = "SELECT TOP 1 TradeDate FROM TradeDate WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate desc";

                DataTable dv = MSSQL.ExecSqlQry(sql, LoginSet.tsquoteSqlConnString);

                GlobalVar.globalParameter.lastTradeDate = Convert.ToDateTime(dv.Rows[0]["TradeDate"]);
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[GlobalUtil_GetLastTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetNextTradeDate() {
            try {
                string sql = "SELECT TOP 3 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)>CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate";
                DataTable dv = MSSQL.ExecSqlQry(sql, LoginSet.tsquoteSqlConnString);

                GlobalVar.globalParameter.nextTradeDate1 = Convert.ToDateTime(dv.Rows[0]["TradeDate"]);
                GlobalVar.globalParameter.nextTradeDate2 = Convert.ToDateTime(dv.Rows[1]["TradeDate"]);
                GlobalVar.globalParameter.nextTradeDate3 = Convert.ToDateTime(dv.Rows[2]["TradeDate"]);
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[FrmIssueTable_GetNextTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetFirstTradeDateOfQuarter() {
            DateTime dt = DateTime.Now;
            string startQuarter = dt.AddMonths(0 - (dt.Month - 1) % 3).AddDays(1 - dt.Day).ToString("yyyyMMdd");

            string sql = "SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND TradeDate >= '" + startQuarter + "' ORDER BY TradeDate";
            DataTable dv = MSSQL.ExecSqlQry(sql, LoginSet.tsquoteSqlConnString);
            GlobalVar.globalParameter.firstTradeDateQ = Convert.ToDateTime(dv.Rows[0][0]);
        }

        public static void Close() {
            if (GlobalVar.autoWork != null)
                GlobalVar.autoWork.Dispose();
        }
    }

    public static class LoginSet
    {
        public static string edisSqlConnString = "SERVER=10.10.1.30;DATABASE=EDIS;UID=WarrantWeb;PWD=WarrantWeb";
        public static string tsquoteSqlConnString = "SERVER=10.60.0.37;DATABASE=TsQuote;UID=WarrantWeb;PWD=WarrantWeb";
        public static string warrantSysSqlConnString = "SERVER=BSSDB;DATABASE=WAFT;UID=warpap;PWD=warpap";
        public static string edaisysConnString = "SERVER=BSSDB;DATABASE=EDAISYS;UID=eduser;PWD=eduser";
    }

    public class GlobalParameter
    {
        public double interestRate = 0.0;
        public double givenRewardPercent = 0.0;
        public bool isLevelA;
        public bool isTodayTradeDate = false;
        public DateTime lastTradeDate;
        public DateTime nextTradeDate1;
        public DateTime nextTradeDate2;
        public DateTime nextTradeDate3;
        public DateTime firstTradeDateQ;

    }
}
