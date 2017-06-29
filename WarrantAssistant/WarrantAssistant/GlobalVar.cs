using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.IO;
using System.Windows.Forms;

namespace WarrantAssistant
{
    
    public class GlobalVar
    {
        public static MainForm mainForm;
        public static AutoWork autoWork;
        public static GlobalParameter globalParameter;
        public static LoginSet loginSet;
    }
    public class GlobalUtility
    {
        public static void Start()
        {
            loadLoginSet();
            loadGlobalParameters();
            //GlobalVar.errProcess = new ErrProcess();
            //GlobalVar.autoWork = new AutoWork();
            //GlobalVar.warrantPriceUpdator = new WarrantPriceUpdator();
            //GlobalVar.warrantPriceProcess = new WarrantPriceProcess();
        }
        public static T MenuItemClick<T>() where T : Form, new() {
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms) {
                if (iForm.GetType() == typeof(T)) {
                    iForm.BringToFront();
                    return (T)iForm;
                }
            }
            T form = new T();
            form.StartPosition = FormStartPosition.CenterScreen;
            form.Show();
            return form;
        }

        public static void SelectUnderlying(string underlyingID, DataGridView dataGridView1) {
            for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                string uID = (string) dataGridView1.Rows[i].Cells[0].Value;
                if (uID == underlyingID)
                    dataGridView1.CurrentCell = dataGridView1.Rows[i - 1].Cells[0];
            }
        }
        public static string GetHtml(string url) {
            string firstResponse = null;
            try {
                WebRequest req = WebRequest.Create(url);
                WebResponse resp = req.GetResponse();
                Stream dataStream = resp.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream, System.Text.Encoding.Default);
                firstResponse = reader.ReadToEnd();

                //Close connection
                req.Abort();
                resp.Close();
                dataStream.Close();
                reader.Close();
            } catch (Exception err) {
                MessageBox.Show(err.ToString());
            }
            return firstResponse;
        }

        public static void LogInfo(string type, string content) {
            string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
            List<SqlParameter> psInfo = new List<SqlParameter>();
            psInfo.Add(new SqlParameter("@MDate" , SqlDbType.DateTime));
            psInfo.Add(new SqlParameter("@InformationType" , SqlDbType.VarChar));
            psInfo.Add(new SqlParameter("@InformationContent" , SqlDbType.VarChar));
            psInfo.Add(new SqlParameter("@MUser" , SqlDbType.VarChar));

            SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString , sqlInfo , psInfo);
            hInfo.SetParameterValue("@MDate" , DateTime.Now);
            hInfo.SetParameterValue("@InformationType" , type);
            hInfo.SetParameterValue("@InformationContent" , content);
            hInfo.SetParameterValue("@MUser" , GlobalVar.globalParameter.userID);
            hInfo.ExecuteCommand();
            hInfo.Dispose();

        }

        public static void loadLoginSet()
        {
            GlobalVar.loginSet = new LoginSet();
            GlobalVar.loginSet.edisSqlConnString = "SERVER=10.10.1.30;DATABASE=EDIS;UID=WarrantWeb;PWD=WarrantWeb";
            GlobalVar.loginSet.tsquoteSqlConnString = "SERVER=10.60.0.37;DATABASE=TsQuote;UID=WarrantWeb;PWD=WarrantWeb";
            GlobalVar.loginSet.warrantSysSqlConnString = "SERVER=10.7.0.52;DATABASE=WAFT;UID=eduser;PWD=eduser";
            GlobalVar.loginSet.warrantSysKeySqlConnString = "SERVER=10.7.0.52;DATABASE=EDAISYS;UID=eduser;PWD=eduser";
        }

        public static void loadGlobalParameters()
        {
            if (GlobalVar.globalParameter == null)
                GlobalVar.globalParameter = new GlobalParameter();

            checkGlobal();

            checkIsTodayTradeDate();
            getNextTradeDate();
            getLastTradeDate();
        }
        private static void checkGlobal()
        {
            DataView dv = DeriLib.Util.ExecSqlQry("SELECT [InterestRate],[GivenRewardPercent],[IsLevelA],[DayPerYear],[ResultTime] FROM [EDIS].[dbo].[Global]", GlobalVar.loginSet.edisSqlConnString);
            GlobalVar.globalParameter.interestRate = Convert.ToDouble(dv[0]["InterestRate"]);
            //A集券商獎勵額度目前為1%
            GlobalVar.globalParameter.givenRewardPercent = Convert.ToDouble(dv[0]["GivenRewardPercent"]);
            //本季是否為A級券商
            GlobalVar.globalParameter.isLevelA = Convert.ToBoolean(dv[0]["IsLevelA"]);
            GlobalVar.globalParameter.dayPerYear = Convert.ToInt32(dv[0]["DayPerYear"]);
            GlobalVar.globalParameter.resultTime = Convert.ToInt32(dv[0]["ResultTime"]);

        }
        private static void checkIsTodayTradeDate()
        {
            DataView dv = DeriLib.Util.ExecSqlQry("SELECT IsTrade FROM [TradeDate] WHERE CONVERT(VARCHAR, TradeDate, 112) = CONVERT(VARCHAR, GETDATE(), 112)", GlobalVar.loginSet.tsquoteSqlConnString);
            if (dv[0]["IsTrade"].ToString() == "Y")
                GlobalVar.globalParameter.isTodayTradeDate = true;
            else
                GlobalVar.globalParameter.isTodayTradeDate = false;
        }

        private static void getLastTradeDate()
        {
            try
            {
                string sql = "SELECT TOP 1 TradeDate FROM TradeDate WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate desc";
                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);

                GlobalVar.globalParameter.lastTradeDate = Convert.ToDateTime(dv[0]["TradeDate"]);
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[GlobalUtil_GetLastTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void getNextTradeDate()
        {
            try
            {
                string sql = "SELECT TOP 3 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)>CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate";
                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);

                GlobalVar.globalParameter.nextTradeDate1 = Convert.ToDateTime(dv[0]["TradeDate"]);
                GlobalVar.globalParameter.nextTradeDate2 = Convert.ToDateTime(dv[1]["TradeDate"]);
                GlobalVar.globalParameter.nextTradeDate3 = Convert.ToDateTime(dv[2]["TradeDate"]);
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[FrmIssueTable_GetNextTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        public static void close()
        {
            if (GlobalVar.autoWork != null) { GlobalVar.autoWork.Dispose(); }
        }
    }

    public class LoginSet
    {
        public string edisSqlConnString = "";
        public string tsquoteSqlConnString = "";
        public string warrantSysSqlConnString = "";
        public string warrantSysKeySqlConnString = "";
    }

    public class GlobalParameter
    {
        public string userID = "";
        public string userDeputy = "";
        public string userGroup = "";
        public string userLevel = "";
        public string userName = "";

        public double interestRate = 0.025;
        public int dayPerYear = 365;
        public double givenRewardPercent = 0.01;
        public bool isLevelA;
        public int resultTime = 640;
        public bool isTodayTradeDate = false;
        public DateTime lastTradeDate;
        public DateTime nextTradeDate1;
        public DateTime nextTradeDate2;
        public DateTime nextTradeDate3;

    }
}
