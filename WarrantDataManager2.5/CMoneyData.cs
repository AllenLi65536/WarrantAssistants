using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using EDLib.SQL;

namespace WarrantDataManager
{
    public static class CMoneyData
    {
        public static CMADODB5.CMConnection cn = new CMADODB5.CMConnection();
        public static string arg = "5"; //%
        public static string srvLocation = "10.60.0.191";
        public static string cnPort = "";
        private static Dictionary<string, CommodityData> data = new Dictionary<string, CommodityData>();
        private static List<string> tw50Stocks = new List<string>();

        public static WorkState LoadCMoneyData() {
            try {
                GetTW50Stcoks();
                LoadCommodityData();
                GetPricesAndPERatio();
                GetEarning();
                GetDividendDates();
                GetPublicOfferingAdjustDate();
                GetDisposeEndDate();
                GetWatchStock();
                GetWarningScore();
                GetAccNetIncome();
                UpdateIssueCheck();
                return WorkState.Successful;
            } catch (Exception) {
                return WorkState.Exception;
            }
        }

        /*public static void updateWarrantUnderlying() {
            loadCommodityData();
            refreshCommodityBasics();
        }*/

        public static CommodityBasicList GetCommodityBasics() {
            CommodityBasicList cBL = new CommodityBasicList();

            try {
                string sql = "SELECT [股票代號], [股票名稱], isNull([上市上櫃],'1') 市場, IsNull([公司名稱], '') 公司名稱, IsNull([統一編號], '00000000') 統一編號 FROM [上市櫃公司基本資料] WHERE ";
                List<string> datas = new List<string>();
                //DataView dv = DeriLib.Util.ExecSqlQry("SELECT WRTCAN_CMONEY_ID FROM [V_CANDIDATE] ORDER BY WRTCAN_CMONEY_ID", LoginSet.warrantSysSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry("SELECT WRTCAN_CMONEY_ID FROM [CANDIDATE] WHERE WRTCAN_DATE = (select max(WRTCAN_DATE) from [WAFT].[dbo].[CANDIDATE]) ORDER BY WRTCAN_CMONEY_ID", LoginSet.warrantSysSqlConnString);//V_CANDIDATE
                string cStr = "";
                foreach (DataRow dr in dv.Rows)
                    cStr += "'" + dr["WRTCAN_CMONEY_ID"].ToString() + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string cid = Convert.ToString(rs.Fields["股票代號"].Value);
                    string cnm = Convert.ToString(rs.Fields["股票名稱"].Value);
                    string mktN = Convert.ToString(rs.Fields["市場"].Value);
                    string mkt = "";
                    if (mktN == "1")
                        mkt = "TSE";
                    else if (mktN == "2")
                        mkt = "OTC";
                    else
                        mkt = "";
                    string uid = Convert.ToString(rs.Fields["統一編號"].Value);
                    string fnm = Convert.ToString(rs.Fields["公司名稱"].Value);

                    if (uid != null && cid != null && cnm != null && mkt != null && fnm != null && uid != "00000000" && uid != "0" && uid != "" && !datas.Contains(uid)) {
                        CommodityBasic b = new CommodityBasic(cid, cnm, mkt, uid, fnm);
                        datas.Add(uid);
                        cBL.add(b);
                    }
                }

            } catch (Exception ex) {
                MessageBox.Show("GetCommodityBasics" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getCommodityBasics][" + ex.Message + "][" + ex.StackTrace + "]");
            }

            return cBL;
        }

        private static void LoadCommodityData() {
            try {
                data.Clear();
                string sql = "SELECT UnderlyingID, UnderlyingIDCMoney, UnderlyingName FROM [WarrantUnderlying] WHERE StockType='DS' or StockType='DR' ORDER BY UnderlyingID";
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, LoginSet.edisSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry(sql, LoginSet.edisSqlConnString);

                foreach (DataRow dr in dv.Rows) {
                    string commodityID = dr["UnderlyingID"].ToString();
                    string commodityIDCMoney = dr["UnderlyingIDCMoney"].ToString();
                    string commodityName = dr["UnderlyingName"].ToString();

                    if (!data.ContainsKey(commodityID))
                        data.Add(commodityID, new CommodityData(commodityID, commodityName, tw50Stocks.Contains(commodityIDCMoney)));

                }
            } catch (Exception ex) {
                MessageBox.Show("LoadCommodityData" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_loadCommodityData][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void RefreshCommodityBasics() {
            try {
                string sql = "SELECT [股票代號], [股票名稱], isNull([上市上櫃],'1') 市場, IsNull([公司名稱], '') 公司名稱, IsNull([統一編號], '00000000') 統一編號 FROM [上市櫃公司基本資料] WHERE ";

                //DataView dv = DeriLib.Util.ExecSqlQry("SELECT WRTCAN_CMONEY_ID FROM [V_CANDIDATE] ORDER BY WRTCAN_CMONEY_ID", LoginSet.warrantSysSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry("SELECT WRTCAN_CMONEY_ID FROM [CANDIDATE] WHERE WRTCAN_DATE = (select max(WRTCAN_DATE) from [WAFT].[dbo].[CANDIDATE]) ORDER BY WRTCAN_CMONEY_ID", LoginSet.warrantSysSqlConnString); // V_CANDIDATE
                string cStr = "";
                foreach (DataRow dr in dv.Rows)
                    cStr += "'" + dr["WRTCAN_CMONEY_ID"].ToString() + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                string cmdText = "UPDATE [WarrantUnderlying] SET UnderlyingName=@UnderlyingName, Market=@Market, UnifiedID=@UnifiedID, FullName=@FullName WHERE UnderlyingID=@Underlying";
                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                pars.Add(new System.Data.SqlClient.SqlParameter("@UnderlyingID", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@UnderlyingName", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@Market", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@FullName", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@UnifiedID", SqlDbType.VarChar));
                SQLCommandHelper h = new SQLCommandHelper(LoginSet.edisSqlConnString, cmdText, pars);


                for (; !rs.EOF; rs.MoveNext()) {
                    string commodityID = rs.Fields["股票代號"].Value;
                    string commodityName = rs.Fields["股票名稱"].Value;
                    string market = rs.Fields["市場"].Value;
                    string unifiedID = rs.Fields["統一編號"].Value;
                    string fullName = rs.Fields["公司名稱"].Value;

                    h.SetParameterValue("@UnderlyingID", commodityID);
                    h.SetParameterValue("@UnderlyingName", commodityName);
                    h.SetParameterValue("@Market", market);
                    h.SetParameterValue("@UnifiedID", unifiedID);
                    h.SetParameterValue("@FullName", fullName);

                    h.ExecuteCommand();
                }
                h.Dispose();
            } catch (Exception ex) {
                MessageBox.Show("RefreshCommodityBasics" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_refreshIssuableUnderlyingData][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetTW50Stcoks() {
            try {
                tw50Stocks.Clear();

                string sql = "select [標的代號] from ETF持股明細表 where 股票代號 = '0050' and 日期 = '" + GlobalVar.globalParameter.lastTradeDate.ToString("yyyyMMdd") + "'";

                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);
                /*DataTable etf = CMoney.ExecCMoneyQry(sql);
                foreach(DataRow row in etf.Rows) {
                    tw50Stocks.Add(row["標的代號"].ToString());
                }*/

                for (; !rs.EOF; rs.MoveNext())
                    tw50Stocks.Add(Convert.ToString(rs.Fields["標的代號"].Value));


                /*string URL = "http://www.twse.com.tw/ch/trading/indices/twco/tai50i.php";
                
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(URL);
                req.Method = "GET";
                WebResponse response = req.GetResponse();
                string htmlstr = "";

                using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    htmlstr = sr.ReadToEnd();
                }

                response.Close();

                htmlstr = htmlstr.Substring(htmlstr.IndexOf("公眾流通量") + 1);

                while (htmlstr.IndexOf("<tr class=tb2>") >= 0)
                {
                    htmlstr = htmlstr.Substring(htmlstr.IndexOf("<tr class=tb2>") + 14);
                    htmlstr = htmlstr.Substring(htmlstr.IndexOf("<td align=center>") + 17);
                    string cid = htmlstr.Substring(0, htmlstr.IndexOf("</td>"));
                    tw50Stocks.Add(cid);
                }*/
            } catch (Exception ex) {
                MessageBox.Show("Gettw50Stocks" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getBigStocks][" + ex.Message + "][" + ex.StackTrace + "]");
            }

        }

        private static void GetPricesAndPERatio() {
            try {
                string dStr = DateTime.Today.ToString("yyyyMMdd");
                string qStr = DateTime.Today.AddMonths(-3).ToString("yyyyMMdd");
                string yStr = DateTime.Today.AddYears(-1).ToString("yyyyMMdd");

                //昨天，三個月前，一年前的交易日
                //DataView dDv = DeriLib.Util.ExecSqlQry("SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<'" + dStr + "' ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                //DataView qDv = DeriLib.Util.ExecSqlQry("SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<'" + qStr + "' ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                //DataView yDv = DeriLib.Util.ExecSqlQry("SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<'" + yStr + "' ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                DataTable dDv = MSSQL.ExecSqlQry("SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<'" + dStr + "' ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                DataTable qDv = MSSQL.ExecSqlQry("SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<'" + qStr + "' ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                DataTable yDv = MSSQL.ExecSqlQry("SELECT TOP 1 TradeDate FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<'" + yStr + "' ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);

                //實際前一交易日，前三個月的交易日，前一年的交易日
                DateTime dDT = Convert.ToDateTime(dDv.Rows[0]["TradeDate"]);
                DateTime qDT = Convert.ToDateTime(qDv.Rows[0]["TradeDate"]);
                DateTime yDT = Convert.ToDateTime(yDv.Rows[0]["TradeDate"]);

                string sql = "SELECT [日期], [股票代號], IsNull([收盤價],0) 收盤價, IsNull([本益比],0) 本益比 FROM [日收盤表排行] WHERE [日期] IN ('" + dDT.ToString("yyyyMMdd") + "','" + qDT.ToString("yyyyMMdd") + "','" + yDT.ToString("yyyyMMdd") + "') AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";

                //把最後一個逗點刪掉
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號], [日期]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    string date = rs.Fields["日期"].Value;
                    double price = Convert.ToDouble(rs.Fields["收盤價"].Value);
                    double pe = Convert.ToDouble(rs.Fields["本益比"].Value);

                    CommodityData d = data[stockID];
                    if (date == dDT.ToString("yyyyMMdd")) {
                        d.peRatio = pe;
                        d.price = price;
                    } else if (date == qDT.ToString("yyyyMMdd"))
                        d.priceQuarter = price;
                    else if (date == yDT.ToString("yyyyMMdd"))
                        d.priceYear = price;

                }

                foreach (CommodityData d in data.Values) {
                    if (d.priceQuarter == 0)
                        d.returnQuarter = 0;
                    else
                        d.returnQuarter = d.price / d.priceQuarter - 1.0;

                    if (d.priceYear == 0)
                        d.returnYear = 0;
                    else
                        d.returnYear = d.price / d.priceYear - 1.0;
                }

            } catch (Exception ex) {
                MessageBox.Show("getPriceAndPERatio" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getPricesAndPERatio][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetEarning() {
            try {
                string sql = "SELECT [年季], [股票代號], isnull([合併總損益(千)], 0) as [合併總損益(千)] FROM [季合併財報(損益單季)] WHERE [年季] IN (SELECT DISTINCT TOP 4 [年季] FROM [季合併財報(損益單季)] ORDER BY [年季] desc) AND "; // 合併淨損益(千)

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號], [年季]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string quarter = rs.Fields["年季"].Value;
                    string stockID = rs.Fields["股票代號"].Value;
                    double earning = Convert.ToDouble(rs.Fields["合併總損益(千)"].Value);

                    data[stockID].commodityEarning.addQuarterEarning(earning);
                }

            } catch (Exception ex) {
                MessageBox.Show("getEarning" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getEarning][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetDividendDates() {
            try {
                string sql = @"SELECT [年度], [股票代號], IsNull([現金股利合計(元)], 0) 現金股利, IsNull([股票股利合計(元)], 0) 股票股利, IsNull([除息日],'') 除息日, isNull([除權日],'') 除權日
                               FROM [股利政策表] WHERE [年度] = '" + DateTime.Today.AddYears(-1).Year.ToString() + "' AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";

                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    string ec = rs.Fields["除息日"].Value == null ? "" : rs.Fields["除息日"].Value;
                    string es = rs.Fields["除權日"].Value == null ? "" : rs.Fields["除權日"].Value;
                    double cd = Convert.ToDouble(rs.Fields["現金股利"].Value);
                    double sd = Convert.ToDouble(rs.Fields["股票股利"].Value);

                    if (ec != "") { data[stockID].exCashDividendDate = DateTime.ParseExact(ec, "yyyyMMdd", null); }
                    if (es != "") { data[stockID].exStockDividendDate = DateTime.ParseExact(es, "yyyyMMdd", null); }
                    data[stockID].cashDividend = cd;
                    data[stockID].stockDividend = sd;
                }
            } catch (Exception ex) {
                MessageBox.Show("GetDividendDates" + ex.Message);
            }
        }

        private static void GetPublicOfferingAdjustDate() {
            try {
                string sql = "SELECT [年度], [股票代號], IsNull([現增除權日],'') 現增除權日 FROM [股利政策表] WHERE [年度] = '" + DateTime.Today.AddYears(-1).Year.ToString() + "' AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    string es = rs.Fields["現增除權日"].Value == null ? "" : rs.Fields["現增除權日"].Value;

                    if (es != "") { data[stockID].exPODate = DateTime.ParseExact(es, "yyyyMMdd", null); }
                }
            } catch (Exception ex) {
                Console.WriteLine("POD error:" + ex.Message);
                Console.ReadLine();
            }
        }

        private static void GetDisposeEndDate() {
            try {
                string sql = "SELECT [年度], [股票代號], IsNull([處置時間迄],'') 處置結束日 FROM [處置股票] WHERE [年度] >= '" + DateTime.Today.AddMonths(-6).Year.ToString() + "' AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號] ,[處置時間迄] ";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    string es = rs.Fields["處置結束日"].Value == null ? "" : rs.Fields["處置結束日"].Value;

                    if (es != "")
                        data[stockID].disposeEndDate = DateTime.ParseExact(es, "yyyyMMdd", null);
                }
            } catch (Exception ex) {
                MessageBox.Show("getDisposeEndDate" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getDisposeEndDates][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetWatchStock() {
            try {
                //找到前六個交易日的日期
                //WARNING!! Problems happen around Chinese new year!!
                string dStr = "";
                //DataView dv = DeriLib.Util.ExecSqlQry("SELECT TOP 6 CONVERT(VARCHAR, TradeDate,112) TD FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                DataTable dv = MSSQL.ExecSqlQry("SELECT TOP 6 CONVERT(VARCHAR, TradeDate,112) TD FROM [TradeDate] WHERE IsTrade='Y' AND CONVERT(VARCHAR,TradeDate,112)<CONVERT(VARCHAR,GETDATE(),112) ORDER BY TradeDate desc", LoginSet.tsquoteSqlConnString);
                foreach (DataRow dr in dv.Rows)
                    dStr += "'" + dr["TD"].ToString() + "',";
                if (dStr.Length > 0)
                    dStr = dStr.Substring(0, dStr.Length - 1);

                string sql = "SELECT [股票代號], IsNull([注意交易資訊],'') 注意 FROM [注意股票] WHERE [日期] IN (" + dStr + ") AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    CommodityData d = data[stockID];

                    string x = rs.Fields["注意"].Value;

                    if (x != "")
                        d.watchCount++;
                }
            } catch (Exception ex) {
                MessageBox.Show("getWatchStock" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getWatchStock][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetWarningScore() {
            try {
                //string sql = "SELECT [股票代號], [警示指標總符合數] FROM [月財務警示指標] WHERE [年月] = '" + DateTime.Today.ToString("yyyyMM") + "' AND ";
                string sql = "SELECT [股票代號], [警示指標總符合數] FROM [月財務警示指標] WHERE [年月] = (select MAX([年月]) from [月財務警示指標]) AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    int w = Convert.ToInt32(rs.Fields["警示指標總符合數"].Value);

                    data[stockID].warningScore = w;
                }
            } catch (Exception ex) {
                MessageBox.Show("getWarningScore" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getWarningScore][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void GetAccNetIncome() {
            try {
                string sql = "";
                DateTime dQ1 = new DateTime(DateTime.Today.Year, 5, 15);
                DateTime dQ2 = new DateTime(DateTime.Today.Year, 8, 15);
                DateTime dQ3 = new DateTime(DateTime.Today.Year, 11, 15);
                DateTime dY = new DateTime(DateTime.Today.Year, 3, 31);

                if (DateTime.Today < dY)
                    sql = $"SELECT [股票代號], [股票名稱], IsNull([合併總損益累計(千)], 0) 稅後純益 FROM [季合併為主財報(損益累計)] WHERE [年季] = '{(DateTime.Today.Year - 1).ToString()}03' AND "; // [稅後純益累計(千)]
                else if (DateTime.Today < dQ1)
                    sql = $"SELECT [股票代號], [股票名稱], IsNull([合併總損益累計(千)], 0) 稅後純益 FROM [季合併為主財報(損益累計)] WHERE [年季] = '{(DateTime.Today.Year - 1).ToString()}04' AND ";
                else if (DateTime.Today < dQ2)
                    sql = $"SELECT [股票代號], [股票名稱], IsNull([合併總損益累計(千)], 0) 稅後純益 FROM [季合併為主財報(損益累計)] WHERE [年季] = '{DateTime.Today.Year.ToString()}01' AND ";
                else if (DateTime.Today < dQ3)
                    sql = $"SELECT [股票代號], [股票名稱], IsNull([合併總損益累計(千)], 0) 稅後純益 FROM [季合併為主財報(損益累計)] WHERE [年季] = '{DateTime.Today.Year.ToString()}02' AND ";
                else
                    sql = $"SELECT [股票代號], [股票名稱], IsNull([合併總損益累計(千)], 0) 稅後純益 FROM [季合併為主財報(損益累計)] WHERE [年季] = '{DateTime.Today.Year.ToString()}03' AND ";

                string cStr = "";
                foreach (string cID in data.Keys)
                    cStr += "'" + cID + "',";
                if (cStr.Length > 0)
                    cStr = cStr.Substring(0, cStr.Length - 1);

                sql += "[股票代號] IN (" + cStr + ") ORDER BY [股票代號]";
                ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sql);

                for (; !rs.EOF; rs.MoveNext()) {
                    string stockID = rs.Fields["股票代號"].Value;
                    data[stockID].accNetIncome = Convert.ToDouble(rs.Fields["稅後純益"].Value);
                }
            } catch (Exception ex) {
                MessageBox.Show("getAccNetIncome" + ex.Message);
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_getAccNetIncome][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private static void UpdateIssueCheck() {
            try {
                SQLCommandHelper deleteIssueCheck = new SQLCommandHelper(LoginSet.edisSqlConnString, "DELETE FROM [WarrantIssueCheck]", new List<System.Data.SqlClient.SqlParameter>());
                deleteIssueCheck.ExecuteCommand();
                deleteIssueCheck.Dispose();

                SQLCommandHelper deleteIssueCheckPut = new SQLCommandHelper(LoginSet.edisSqlConnString, "DELETE FROM [WarrantIssueCheckPut]", new List<System.Data.SqlClient.SqlParameter>());
                deleteIssueCheckPut.ExecuteCommand();
                deleteIssueCheckPut.Dispose();

                string sqlIssueCheck = "INSERT INTO [WarrantIssueCheck] values (@UnderlyingID, @UnderlyingName, @MDate, @CashDividend, @StockDividend, @CashDividendDate, @StockDividendDate, @PublicOfferingDate, @DisposeEndDate, @WatchCount, @WarningScore, @CanIssuePut, @AccNetIncome)";
                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();

                pars.Add(new System.Data.SqlClient.SqlParameter("@UnderlyingID", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@UnderlyingName", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@MDate", SqlDbType.DateTime));
                pars.Add(new System.Data.SqlClient.SqlParameter("@CashDividend", SqlDbType.Float));
                pars.Add(new System.Data.SqlClient.SqlParameter("@StockDividend", SqlDbType.Float));
                pars.Add(new System.Data.SqlClient.SqlParameter("@CashDividendDate", SqlDbType.Date));
                pars.Add(new System.Data.SqlClient.SqlParameter("@StockDividendDate", SqlDbType.Date));
                pars.Add(new System.Data.SqlClient.SqlParameter("@PublicOfferingDate", SqlDbType.Date));
                pars.Add(new System.Data.SqlClient.SqlParameter("@DisposeEndDate", SqlDbType.Date));
                pars.Add(new System.Data.SqlClient.SqlParameter("@WatchCount", SqlDbType.Int));
                pars.Add(new System.Data.SqlClient.SqlParameter("@WarningScore", SqlDbType.Int));
                pars.Add(new System.Data.SqlClient.SqlParameter("@CanIssuePut", SqlDbType.VarChar));
                pars.Add(new System.Data.SqlClient.SqlParameter("@AccNetIncome", SqlDbType.Float));

                string sqlIssueCheckPut = "INSERT INTO [WarrantIssueCheckPut] values (@UnderlyingID, @UnderlyingName, @MDate, @IsTW50Stocks, @PERatio, @SumEarning, @Price, @PriceQuarter, @PriceYear, @ReturnQuarter, @ReturnYear)";
                List<System.Data.SqlClient.SqlParameter> parsPut = new List<System.Data.SqlClient.SqlParameter>();

                parsPut.Add(new System.Data.SqlClient.SqlParameter("@UnderlyingId", SqlDbType.VarChar));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@UnderlyingName", SqlDbType.VarChar));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@MDate", SqlDbType.DateTime));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@IsTW50Stocks", SqlDbType.VarChar));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@PERatio", SqlDbType.Float));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@SumEarning", SqlDbType.Float));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@Price", SqlDbType.Float));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@PriceQuarter", SqlDbType.Float));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@PriceYear", SqlDbType.Float));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@ReturnQuarter", SqlDbType.Float));
                parsPut.Add(new System.Data.SqlClient.SqlParameter("@ReturnYear", SqlDbType.Float));

                SQLCommandHelper insertIssueCheck = new SQLCommandHelper(LoginSet.edisSqlConnString, sqlIssueCheck, pars);
                SQLCommandHelper insertIssueCheckPut = new SQLCommandHelper(LoginSet.edisSqlConnString, sqlIssueCheckPut, parsPut);

                foreach (CommodityData d in data.Values) {
                    d.checkPutIssueability();

                    insertIssueCheck.SetParameterValue("@UnderlyingID", d.commodityID);
                    insertIssueCheck.SetParameterValue("@UnderlyingName", d.commodityName);
                    insertIssueCheck.SetParameterValue("@MDate", DateTime.Now);
                    insertIssueCheck.SetParameterValue("@CashDividend", d.cashDividend);
                    insertIssueCheck.SetParameterValue("@StockDividend", d.stockDividend);
                    if (d.exCashDividendDate.ToString("yyyyMMdd") == "00010101")
                        insertIssueCheck.SetParameterValue("@CashDividendDate", null);
                    else
                        insertIssueCheck.SetParameterValue("@CashDividendDate", d.exCashDividendDate);

                    if (d.exStockDividendDate.ToString("yyyyMMdd") == "00010101")
                        insertIssueCheck.SetParameterValue("@StockDividendDate", null);
                    else
                        insertIssueCheck.SetParameterValue("@StockDividendDate", d.exStockDividendDate);

                    if (d.exPODate.ToString("yyyyMMdd") == "00010101")
                        insertIssueCheck.SetParameterValue("@PublicOfferingDate", null);
                    else
                        insertIssueCheck.SetParameterValue("@PublicOfferingDate", d.exPODate);

                    if (d.disposeEndDate.ToString("yyyyMMdd") == "00010101")
                        insertIssueCheck.SetParameterValue("@DisposeEndDate", null);
                    else
                        insertIssueCheck.SetParameterValue("@DisposeEndDate", d.disposeEndDate);

                    insertIssueCheck.SetParameterValue("@WatchCount", d.watchCount);
                    insertIssueCheck.SetParameterValue("@WarningScore", d.warningScore);
                    insertIssueCheck.SetParameterValue("@CanIssuePut", d.isPutIssuable ? "Y" : "N");
                    insertIssueCheck.SetParameterValue("@AccNetIncome", d.accNetIncome);

                    insertIssueCheckPut.SetParameterValue("@UnderlyingID", d.commodityID);
                    insertIssueCheckPut.SetParameterValue("@UnderlyingName", d.commodityName);
                    insertIssueCheckPut.SetParameterValue("@MDate", DateTime.Now);
                    insertIssueCheckPut.SetParameterValue("@IsTW50Stocks", d.isTW50Stocks ? "Y" : "N");
                    insertIssueCheckPut.SetParameterValue("@PERatio", d.peRatio);
                    insertIssueCheckPut.SetParameterValue("@SumEarning", d.commodityEarning.sumEarning);
                    insertIssueCheckPut.SetParameterValue("@Price", d.price);
                    insertIssueCheckPut.SetParameterValue("@PriceQuarter", d.priceQuarter);
                    insertIssueCheckPut.SetParameterValue("@PriceYear", d.priceYear);
                    insertIssueCheckPut.SetParameterValue("@ReturnQuarter", d.returnQuarter);
                    insertIssueCheckPut.SetParameterValue("@ReturnYear", d.returnYear);

                    insertIssueCheckPut.ExecuteCommand();
                    insertIssueCheck.ExecuteCommand();

                }

                insertIssueCheck.Dispose();
                insertIssueCheckPut.Dispose();
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_updateIssueCheck][" + ex.Message + "][" + ex.StackTrace + "]");
                MessageBox.Show("UpdateIssueCheck" + ex.ToString());
            }

        }

        /*#region IDisposable成元

        public void Dispose() {
            try {
                cn = null;
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[CMoneyWork_Dispose][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        #endregion*/
    }

    public class CommodityEarning
    {
        public CommodityData commodityData;
        public double sumEarning = 0.0;
        private int sumCount = 0;

        public CommodityEarning(CommodityData commodityData) {
            this.commodityData = commodityData;
            this.sumEarning = 0.0;
            this.sumCount = 0;
        }

        public void addQuarterEarning(double earning) {
            if (sumCount > 3)
                return;
            sumEarning += earning;
            sumCount++;
        }
    }

    public class CommodityBasicList
    {
        private List<CommodityBasic> cbList = new List<CommodityBasic>();

        public void add(CommodityBasic c) {
            if (!cbList.Contains(c))
                cbList.Add(c);
        }

        public void clear() {
            if (cbList.Count > 0)
                cbList.Clear();
        }

        public CommodityBasic getByCommodityID(string id) {
            CommodityBasic basic = null;

            foreach (CommodityBasic cb in cbList) {
                if (cb.commodityID == id) {
                    basic = cb;
                    break;
                }
            }

            return basic;
        }

        public CommodityBasic getByUnifiedID(string id) {
            CommodityBasic basic = null;

            foreach (CommodityBasic cb in cbList) {
                if (cb.unifiedID == id) {
                    basic = cb;
                    break;
                }
            }

            return basic;
        }

        public List<CommodityBasic> CommodityBasics { get { return cbList; } }
    }

    public class CommodityBasic
    {
        public string commodityID = "";
        public string commodityName = "";
        public string market = "";
        public string unifiedID = "";
        public string fullName = "";

        public CommodityBasic(string commodityID, string commodityName, string market, string unifiedID, string fullName) {
            this.commodityID = commodityID;
            this.commodityName = commodityName;
            this.market = market;
            this.unifiedID = unifiedID;
            this.fullName = fullName;
        }
    }

    public class CommodityData
    {
        public string commodityID = "";
        public string commodityName = "";
        public string commodityWarrantName = "";

        public bool isPutIssuable = true;
        //若是台灣50成分股則不受Put發行限制
        public bool isTW50Stocks = false;
        //注意天數(前六個交易日內不得>=2天
        public int watchCount = 0;
        //警示分數，不可以有警示
        public int warningScore = 0;

        public double cashDividend = 0.0;
        public double stockDividend = 0.0;

        public DateTime exCashDividendDate;
        public DateTime exStockDividendDate;
        public DateTime exPODate;
        public DateTime disposeEndDate;

        public double accNetIncome = 0.0;

        public double price = 0.0;
        public double priceQuarter = 0.0;
        public double priceYear = 0.0;
        public double returnQuarter = 0.0;
        public double returnYear = 0.0;
        public double peRatio = 0.0;

        public CommodityEarning commodityEarning;

        public CommodityData(string commodityID, string commodityName, bool isTW50Stocks) {
            this.commodityID = commodityID;
            this.commodityName = commodityName;
            this.isTW50Stocks = isTW50Stocks;
            this.commodityEarning = new CommodityEarning(this);
        }

        //Put不可發行時機(台灣50成分股不適用):1.季報酬不得高於50% 2.年報酬不得高於100% 3.本益比不得高於40 4.淨利需大於0
        public void checkPutIssueability() {
            try {
                if (isTW50Stocks) {
                    isPutIssuable = true;
                    return;
                }

                //如果價格資料不足
                if (price == 0.0 || priceQuarter == 0.0 || priceYear == 0.0) {
                    isPutIssuable = false;
                    return;
                }

                //returnQuarter = price / priceQuarter - 1.0;
                //returnYear = price / priceYear - 1.0;

                if (returnQuarter > 0.5 || returnYear > 1.0) {
                    isPutIssuable = false;
                    return;
                }

                if (peRatio > 40.0 || peRatio <= 0.0) {
                    isPutIssuable = false;
                    return;
                }

                if (commodityEarning.sumEarning < 0.0) {
                    isPutIssuable = false;
                    return;
                }

            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[CommodityData_checkPutIssuability][" + ex.Message + "][" + ex.StackTrace + "]");
                isPutIssuable = false;
            }
        }


    }
}
