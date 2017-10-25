using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.IO;
using System.Windows.Forms;
using EDLib.SQL;

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
        public static void Start() {
            LoadLoginSet();
            LoadGlobalParameters();
            //GlobalVar.errProcess = new ErrProcess();
            //GlobalVar.autoWork = new AutoWork();
            //GlobalVar.warrantPriceUpdator = new WarrantPriceUpdator();
            //GlobalVar.warrantPriceProcess = new WarrantPriceProcess();
        }
        public static T MenuItemClick<T>() where T : Form, new() {
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms) {
                if (iForm.GetType() == typeof(T)) {
                    iForm.BringToFront();
                    return (T) iForm;
                }
            }
            T form = new T {
                StartPosition = FormStartPosition.CenterScreen
            };
            form.Show();
            return form;
        }

        public static void SelectUnderlying(string underlyingID, DataGridView dataGridView1) {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                if ((string) dataGridView1.Rows[i].Cells[0].Value == underlyingID)
                    dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
        }
        public static void KeyDecoder(KeyEventArgs e, ref string enteredKey) {
            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) {
                if (enteredKey.Length > 0)
                    enteredKey = enteredKey.Substring(0, enteredKey.Length - 1);
            } else if (e.KeyCode == Keys.Escape)
                enteredKey = "";
            else {
                if (e.KeyCode == Keys.NumPad0 || e.KeyCode == Keys.D0)
                    enteredKey += "0";
                else if (e.KeyCode == Keys.NumPad1 || e.KeyCode == Keys.D1)
                    enteredKey += "1";
                else if (e.KeyCode == Keys.NumPad2 || e.KeyCode == Keys.D2)
                    enteredKey += "2";
                else if (e.KeyCode == Keys.NumPad3 || e.KeyCode == Keys.D3)
                    enteredKey += "3";
                else if (e.KeyCode == Keys.NumPad4 || e.KeyCode == Keys.D4)
                    enteredKey += "4";
                else if (e.KeyCode == Keys.NumPad5 || e.KeyCode == Keys.D5)
                    enteredKey += "5";
                else if (e.KeyCode == Keys.NumPad6 || e.KeyCode == Keys.D6)
                    enteredKey += "6";
                else if (e.KeyCode == Keys.NumPad7 || e.KeyCode == Keys.D7)
                    enteredKey += "7";
                else if (e.KeyCode == Keys.NumPad8 || e.KeyCode == Keys.D8)
                    enteredKey += "8";
                else if (e.KeyCode == Keys.NumPad9 || e.KeyCode == Keys.D9)
                    enteredKey += "9";
                else if (e.KeyCode == Keys.B)
                    enteredKey += "B";
                else if (e.KeyCode == Keys.C)
                    enteredKey += "C";
                else if (e.KeyCode == Keys.P)
                    enteredKey += "P";
                else
                    enteredKey += e.KeyCode.ToString();
            }
            e.Handled = true;
        }

        /*public static string GetHtml(string url) {
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
        }*/

        public static void LogInfo(string type, string content) {
            string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
            List<SqlParameter> psInfo = new List<SqlParameter> {
                new SqlParameter("@MDate", SqlDbType.DateTime),
                new SqlParameter("@InformationType", SqlDbType.VarChar),
                new SqlParameter("@InformationContent", SqlDbType.VarChar),
                new SqlParameter("@MUser", SqlDbType.VarChar)
            };

            SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
            hInfo.SetParameterValue("@MDate", DateTime.Now);
            hInfo.SetParameterValue("@InformationType", type);
            hInfo.SetParameterValue("@InformationContent", content);
            hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
            hInfo.ExecuteCommand();
            hInfo.Dispose();

        }

        public static void LoadLoginSet() {
            GlobalVar.loginSet = new LoginSet {
                edis20SqlConnString = "SERVER=10.19.1.20;DATABASE=EDIS;UID=WarrantWeb;PWD=WarrantWeb",
                edisSqlConnString = "SERVER=10.10.1.30;DATABASE=EDIS;UID=WarrantWeb;PWD=WarrantWeb",
                tsquoteSqlConnString = "SERVER=10.60.0.37;DATABASE=TsQuote;UID=WarrantWeb;PWD=WarrantWeb",
                warrantSysSqlConnString = "SERVER=10.7.0.52;DATABASE=WAFT;UID=eduser;PWD=eduser",
                warrantSysKeySqlConnString = "SERVER=10.7.0.52;DATABASE=EDAISYS;UID=eduser;PWD=eduser"
            };
        }

        public static void LoadGlobalParameters() {
            if (GlobalVar.globalParameter == null)
                GlobalVar.globalParameter = new GlobalParameter();

            CheckGlobal();

            CheckIsTodayTradeDate();
            GetNextTradeDate();
            GetLastTradeDate();
        }
        private static void CheckGlobal() {
            //DataView dv = DeriLib.Util.ExecSqlQry("SELECT [InterestRate],[GivenRewardPercent],[IsLevelA],[DayPerYear],[ResultTime] FROM [EDIS].[dbo].[Global]", GlobalVar.loginSet.edisSqlConnString);
            DataTable dt = MSSQL.ExecSqlQry("SELECT [InterestRate],[GivenRewardPercent],[IsLevelA],[DayPerYear],[ResultTime] FROM [EDIS].[dbo].[Global]", GlobalVar.loginSet.edisSqlConnString);
            GlobalVar.globalParameter.interestRate = Convert.ToDouble(dt.Rows[0]["InterestRate"]);
            //A集券商獎勵額度目前為1%
            GlobalVar.globalParameter.givenRewardPercent = Convert.ToDouble(dt.Rows[0]["GivenRewardPercent"]);
            //本季是否為A級券商
            GlobalVar.globalParameter.isLevelA = Convert.ToBoolean(dt.Rows[0]["IsLevelA"]);
            GlobalVar.globalParameter.dayPerYear = Convert.ToInt32(dt.Rows[0]["DayPerYear"]);
            GlobalVar.globalParameter.resultTime = Convert.ToInt32(dt.Rows[0]["ResultTime"]);

            GlobalVar.globalParameter.traders = new List<string>();
            dt = MSSQL.ExecSqlQry("Select UserID from Trader", "SERVER=10.101.10.5;DATABASE=WMM3;UID=hedgeuser;PWD=hedgeuser");
            foreach (DataRow row in dt.Rows) 
                GlobalVar.globalParameter.traders.Add(row[0].ToString());
            
        }
        private static void CheckIsTodayTradeDate() {
            //DataView dv = DeriLib.Util.ExecSqlQry("SELECT IsTrade FROM [TradeDate] WHERE CONVERT(VARCHAR, TradeDate, 112) = CONVERT(VARCHAR, GETDATE(), 112)", GlobalVar.loginSet.tsquoteSqlConnString);
            DataTable dv = MSSQL.ExecSqlQry("SELECT IsTrade FROM [TradeDate] WHERE CONVERT(VARCHAR, TradeDate, 112) = CONVERT(VARCHAR, GETDATE(), 112)", GlobalVar.loginSet.tsquoteSqlConnString);
            if (dv.Rows[0]["IsTrade"].ToString() == "Y")
                GlobalVar.globalParameter.isTodayTradeDate = true;
            else
                GlobalVar.globalParameter.isTodayTradeDate = false;
        }

        private static void GetLastTradeDate() {
            try {
                string sql = "SELECT TOP 1 TradeDate FROM TradeDate WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate desc";
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
                GlobalVar.globalParameter.lastTradeDate = Convert.ToDateTime(dv.Rows[0]["TradeDate"]);
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[GlobalUtil_GetLastTradeDate][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetNextTradeDate() {
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

        public static void Close() {
            if (GlobalVar.autoWork != null) { GlobalVar.autoWork.Dispose(); }
        }
    }

    public class LoginSet
    {
        public string edis20SqlConnString = "";
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
        public List<string> traders;

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
