using System;
using System.Data;
using EDLib.SQL;

namespace WarrantDataManager2._0
{
    public class GlobalVar
    {
        public static MainForm mainForm;
        public static AutoWork autoWork;
        //public static ErrProcess errProcess;
        public static GlobalParameter globalParameter;
        public static LoginSet loginSet;
    }

    public class GlobalUtility
    {
        public static void Start() {
            loadLoginSet();
            loadGlobalParameters();
            //GlobalVar.errProcess = new ErrProcess();
            GlobalVar.autoWork = new AutoWork();
            //GlobalVar.warrantPriceUpdator = new WarrantPriceUpdator();
            //GlobalVar.warrantPriceProcess = new WarrantPriceProcess();
        }

        public static void loadLoginSet() {
            GlobalVar.loginSet = new LoginSet();
            GlobalVar.loginSet.edisSqlConnString = "SERVER=10.10.1.30;DATABASE=EDIS;UID=WarrantWeb;PWD=WarrantWeb";
            GlobalVar.loginSet.tsquoteSqlConnString = "SERVER=10.60.0.37;DATABASE=TsQuote;UID=WarrantWeb;PWD=WarrantWeb";
            GlobalVar.loginSet.warrantSysSqlConnString = "SERVER=10.7.0.52;DATABASE=WAFT;UID=eduser;PWD=eduser";
        }

        public static void loadGlobalParameters() {
            if (GlobalVar.globalParameter == null)
                GlobalVar.globalParameter = new GlobalParameter();

            GlobalVar.globalParameter.interestRate = 0.025;
            //A集券商獎勵額度目前為1%
            GlobalVar.globalParameter.givenRewardPercent = 0.01;
            //本季是否為A級券商
            //GlobalVar.globalParameter.isLevelA = true;
            checkIsLevelA();
            checkIsTodayTradeDate();
            getNextTradeDate();
            getLastTradeDate();
            getFirstTradeDateOfQuarter();
        }

        private static void checkIsLevelA() {
            DataTable isA = MSSQL.ExecSqlQry("select FLGDAT_FLGDSC from EDAISYS.dbo.FLAGDATAS "
                                         + " where FLGDAT_FLGVAR + 1911 = DATEPART(yyyy, GETDATE())"
                                         + " AND FLGDAT_FLGDTA = DATEPART(q, GETDATE())"
                                         + " AND FLGDAT_FLGNAM = 'WRT_MARKET_RATING'",
                                         "SERVER=10.7.0.52;DATABASE=EDAISYS;UID=eduser;PWD=eduser");
            if (isA.Rows.Count > 0 && isA.Rows[0][0].ToString() == "A")
                GlobalVar.globalParameter.isLevelA = true;
            else
                GlobalVar.globalParameter.isLevelA = false;            
        }
        private static void checkIsTodayTradeDate() {
            //DataView dv = DeriLib.Util.ExecSqlQry("SELECT IsTrade FROM [TradeDate] WHERE CONVERT(VARCHAR, TradeDate, 112) = CONVERT(VARCHAR, GETDATE(), 112)", GlobalVar.loginSet.tsquoteSqlConnString);
            DataTable dv = MSSQL.ExecSqlQry("SELECT IsTrade FROM [TradeDate] WHERE CONVERT(VARCHAR, TradeDate, 112) = CONVERT(VARCHAR, GETDATE(), 112)", GlobalVar.loginSet.tsquoteSqlConnString);
            if (dv.Rows[0]["IsTrade"].ToString() == "Y")
                GlobalVar.globalParameter.isTodayTradeDate = true;
            else
                GlobalVar.globalParameter.isTodayTradeDate = false;            
        }

        private static void getLastTradeDate() {
            try {
                string sql = "SELECT TOP 1 TradeDate FROM TradeDate WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate desc";
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);

                GlobalVar.globalParameter.lastTradeDate = Convert.ToDateTime(dv.Rows[0]["TradeDate"]);
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[GlobalUtil_GetLastTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void getNextTradeDate() {
            try {
                string sql = "SELECT TOP 3 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)>CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate";
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);

                GlobalVar.globalParameter.nextTradeDate1 = Convert.ToDateTime(dv.Rows[0]["TradeDate"]);
                GlobalVar.globalParameter.nextTradeDate2 = Convert.ToDateTime(dv.Rows[1]["TradeDate"]);
                GlobalVar.globalParameter.nextTradeDate3 = Convert.ToDateTime(dv.Rows[2]["TradeDate"]);
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[FrmIssueTable_GetNextTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void getFirstTradeDateOfQuarter() {
            DateTime dt = DateTime.Now;
            string startQuarter = dt.AddMonths(0 - (dt.Month - 1) % 3).AddDays(1 - dt.Day).ToString("yyyyMMdd");

            string sql = "SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND TradeDate >= '" + startQuarter + "' ORDER BY TradeDate";
            //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
            DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
            GlobalVar.globalParameter.firstTradeDateQ = Convert.ToDateTime(dv.Rows[0][0]);
        }

        public static void close() {
            if (GlobalVar.autoWork != null) { GlobalVar.autoWork.Dispose(); }
        }
    }

    public class LoginSet
    {
        public string edisSqlConnString = "";
        public string tsquoteSqlConnString = "";
        public string warrantSysSqlConnString = "";
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
