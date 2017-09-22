using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;
using EDLib.SQL;

namespace WarrantAssistant
{
    public partial class MainForm:Form
    {
        //private delegate void ShowHandler();        

        public MainForm() {
            InitializeComponent();
        }

        //private SafeQueue workQueue = new SafeQueue();
        private Thread workThread;
        /*
        public void AddWork(Work work)
        {
            workQueue.Enqueue(work);
        }
        */

        private void MainForm_Load(object sender, EventArgs e) {
            GlobalVar.mainForm = this;
            FrmLogIn frmLogin = new FrmLogIn();
            if (!frmLogin.TryIPLogin()) {
                MessageBox.Show("Auto login failed. Please e-mail your IP address to allen.li@kgi.com");

                Close();
                //frmLogin.ShowDialog();
            }
            if (!frmLogin.loginOK)
                Close();
        }

        public void Start() {
            if (GlobalVar.globalParameter.userGroup == "TR") {
                行政ToolStripMenuItem.Visible = false;
                財工ToolStripMenuItem.Visible = false;
            }

            if (GlobalVar.globalParameter.userGroup == "AD") {
                traderToolStripMenuItem.Visible = false;
                財工ToolStripMenuItem.Visible = false;
            }

            代理人發行條件輸入ToolStripMenuItem.Visible = false;
            代理人增額條件輸入ToolStripMenuItem.Visible = false;


            SetUltraGrid(dtInfo, ultraGrid1);
            SetUltraGrid(dtAnnounce, ultraGrid2);

            GlobalVar.autoWork = new AutoWork();

            workThread = new Thread(new ThreadStart(RoutineWork));
            workThread.Start();
        }
        private void RoutineWork() {
            try {
                for (;;) {
                    try {
                        if (ultraGrid1.InvokeRequired)
                            ultraGrid1.Invoke(new System.Action(LoadUltraGrid1));
                        else
                            LoadUltraGrid1();

                        if (ultraGrid2.InvokeRequired)
                            ultraGrid2.Invoke(new System.Action(LoadUltraGrid2));
                        else
                            LoadUltraGrid2();

                    } catch (Exception ex) {
                        MessageBox.Show("In main form routine work in for loop " + ex.Message);
                    }
                    Thread.Sleep(10000);
                }

            } catch (Exception ex) {
                //MessageBox.Show("In main form routine work "+ex.Message);
            }
        }
        /*
        private void RoutineWork()
        {
            try
            {
                for (; ; )
                {
                    while (workQueue.Count > 0)
                    {
                        try
                        {
                            object obj = workQueue.Dequeue();
                            if (obj != null)
                            {
                                Work work = (Work)obj;
                                WorkState workstate = work.DoWork();
                                work.Close();
                            }
                        }
                        catch (ThreadAbortException tex)
                        {
                            //MessageBox.Show(tex.Message);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
            }
        }
        */
        public System.Data.DataTable dtInfo = new System.Data.DataTable();
        public System.Data.DataTable dtAnnounce = new System.Data.DataTable();


        public void SetUltraGrid(System.Data.DataTable dt, UltraGrid grid) {
            dt.Columns.Add("時間", typeof(string));
            dt.Columns.Add("內容", typeof(string));
            dt.Columns.Add("人員", typeof(string));
            grid.DataSource = dt;

            UltraGridBand band0 = grid.DisplayLayout.Bands[0];
            band0.Columns["時間"].Width = 60;
            band0.Columns["人員"].Width = 30;
            //ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            band0.ColHeadersVisible = false;
            grid.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            grid.DisplayLayout.Override.CellAppearance.BorderAlpha = Alpha.Transparent;
            grid.DisplayLayout.Override.RowAppearance.BorderAlpha = Alpha.Transparent;
            band0.Columns[2].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;

            band0.Columns["時間"].CellActivation = Activation.NoEdit;
            band0.Columns["內容"].CellActivation = Activation.NoEdit;
            band0.Columns["人員"].CellActivation = Activation.NoEdit;

        }

        private void LoadUltraGrid(System.Data.DataTable dt, string infoOrAnnounce) {
            try {
                string sql = @"SELECT [MDate]
                                  ,[InformationContent]
                                  ,[MUser]
                              FROM [EDIS].[dbo].[InformationLog]
                              WHERE InformationType='" + infoOrAnnounce + "'";
                sql += "AND CONVERT(VARCHAR,Date,112) >='" + GlobalVar.globalParameter.lastTradeDate.ToString("yyyy-MM-dd") + "' ORDER BY MDate DESC";
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                if (dt.Rows.Count == dv.Rows.Count)
                    return;

                dt.Rows.Clear();
                foreach (DataRow drv in dv.Rows) {
                    DataRow dr = dt.NewRow();

                    DateTime md = Convert.ToDateTime(drv["MDate"]);
                    dr["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                    dr["內容"] = drv["InformationContent"].ToString();
                    dr["人員"] = drv["MUser"].ToString();
                    dt.Rows.Add(dr);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        public void LoadUltraGrid1() {
            LoadUltraGrid(dtInfo, "Info");
        }
        public void LoadUltraGrid2() {
            LoadUltraGrid(dtAnnounce, "Announce");
        }

        private void 標的SummaryToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmUnderlyingSummary>();
        }

        private void 標的發行檢查ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueCheck>();
        }

        private void put發行檢查ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueCheckPut>();
        }

        private void 已發行權證ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmWarrant>();
        }

        private void 可增額列表ToolStripMenuItem1_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssueInput>();
        }

        private void 可增額列表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssuable>();
        }

        private void 試算表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<Frm71>();
        }

        private void 發行條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmApply>();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e) {
            if (workThread != null && workThread.IsAlive) { workThread.Abort(); }
            GlobalUtility.Close();
        }

        private void toolStripButton1_Click(object sender, EventArgs e) {
            string info = "";
            info = toolStripTextBox1.Text;
            if (info != "") {
                GlobalUtility.LogInfo("Announce", info);

                toolStripTextBox1.Text = "";
                LoadUltraGrid2();//
            }
        }

        private void 增額條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssue>();
        }

        private void 搶額度總表含增額ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmApplyTotalList>();
        }

        private void 發行總表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueTotal>();
        }

        private void 增額總表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssueTotal>();
        }

        private void 代理人發行條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms) {
                if (iForm.GetType() == typeof(FrmApply)) {
                    iForm.BringToFront();
                    return;
                }
            }
            FrmApply frmApplyDeputy = new FrmApply();
            frmApplyDeputy.userID = GlobalVar.globalParameter.userDeputy;
            frmApplyDeputy.StartPosition = FormStartPosition.CenterScreen;
            frmApplyDeputy.Show();

        }

        private void 代理人增額條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms) {
                if (iForm.GetType() == typeof(FrmReIssue)) {
                    iForm.BringToFront();
                    return;
                }
            }

            FrmReIssue frmReIssue = new FrmReIssue();
            frmReIssue.userID = GlobalVar.globalParameter.userDeputy;
            frmReIssue.StartPosition = FormStartPosition.CenterScreen;
            frmReIssue.Show();

        }

        private void 已發權證條件發行ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueByCurrent>();
        }

        private void 詳細LOGToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmLog>();
        }

        private void 轉申請發行TXTToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileTSE = "D:\\權證發行_相關Excel\\上傳檔\\TSE申請上傳檔.txt";
            string fileOTC = "D:\\權證發行_相關Excel\\上傳檔\\OTC申請上傳檔.txt";

            //TXTFileWriter tseWriter = new TXTFileWriter(fileTSE);
            //TXTFileWriter otcWriter = new TXTFileWriter(fileOTC);           

            int tseCount = 0;
            int otcCount = 0;

            int tseReissue = 0;
            int otcReissue = 0;

            int tseReward = 0;
            int otcReward = 0;

            try {
                using (StreamWriter tseWriter = new StreamWriter(fileTSE, false, Encoding.GetEncoding("Big5")))
                using (StreamWriter otcWriter = new StreamWriter(fileOTC, false, Encoding.GetEncoding("Big5"))) {

                    string sql = @"SELECT a.ApplyKind
                                      ,a.Market
	                                  ,a.WarrantName
                                      ,a.UnderlyingID
                                      ,a.IssueNum
                                      ,a.CR
                                      ,IsNull(CASE WHEN a.ApplyKind='2' THEN c.T ELSE b.T END,6) T
                                      ,a.Type
                                      ,a.CP
                                      ,CASE WHEN a.UseReward='Y' THEN '1' ELSE '0' END UseReward
                                      ,CASE WHEN a.MarketTmr='Y' THEN '1' Else '0' END MarketTmr
                                  FROM [EDIS].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [EDIS].[dbo].[ApplyOfficial] b ON a.SerialNum=b.SerialNumber
                                  LEFT JOIN [EDIS].[dbo].[WarrantBasic] c ON a.WarrantName=c.WarrantName
                                  ORDER BY a.Market desc, a.Type, a.CP, a.UnderlyingID, a.SerialNum";//a.SerialNum
                    //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                    System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                    if (dv.Rows.Count > 0) {
                        foreach (DataRow dr in dv.Rows) {
                            string applyKind = dr["ApplyKind"].ToString();
                            string market = dr["Market"].ToString();
                            string warrantName = dr["WarrantName"].ToString();
                            string underlyingID = dr["UnderlyingID"].ToString();
                            double issueNum = Convert.ToDouble(dr["IssueNum"]);
                            double cr = Convert.ToDouble(dr["CR"]);
                            int t = Convert.ToInt32(dr["T"]);
                            string type = dr["Type"].ToString();
                            string cp = dr["CP"].ToString();
                            string useReward = dr["UseReward"].ToString();
                            string marketTmr = dr["MarketTmr"].ToString();

                            string markup = "                                     ";
                            int byteLen = System.Text.Encoding.Default.GetBytes(warrantName).Length;
                            warrantName = warrantName + markup.Substring(0, 16 - byteLen);
                            underlyingID = underlyingID.PadRight(12, ' ');
                            string issueNumS = issueNum.ToString().PadLeft(7, '0');
                            //issueNumS = issueNumS.PadLeft(7, '0');
                            string crS = (cr * 10000).ToString();
                            crS = crS.Substring(0, Math.Min(5, crS.Length));
                            crS = crS.PadLeft(5, '0');
                            string tS = t.ToString().PadLeft(2, '0');
                            //tS = tS.PadLeft(2, '0');

                            string tempType = "1";
                            if (type == "牛熊證") {
                                if (cp == "P")
                                    tempType = "4";
                                else
                                    tempType = "3";
                            } else {
                                if (cp == "P")
                                    tempType = "2";
                                else
                                    tempType = "1";
                            }

                            string writestr = "";
                            writestr = warrantName + underlyingID + issueNumS + crS + tS + tempType + useReward + marketTmr;

                            if (market == "TSE") {
                                //tseWriter.WriteFile(writestr);
                                tseWriter.WriteLine(writestr);
                                tseCount++;
                                if (useReward == "1")
                                    tseReward++;
                                if (applyKind == "2")
                                    tseReissue++;
                            } else if (market == "OTC") {
                                //otcWriter.WriteFile(writestr);
                                otcWriter.WriteLine(writestr);
                                otcCount++;
                                if (useReward == "1")
                                    otcReward++;
                                if (applyKind == "2")
                                    otcReissue++;
                            }
                        }
                    }
                }

                string infoStr = "TSE 共" + tseCount + "檔，增額" + tseReissue + "檔，獎勵" + tseReward + "檔。\nOTC共" + otcCount + "檔，增額" + otcReissue + "檔，獎勵" + otcReward + "檔。";

                GlobalUtility.LogInfo("Info", "今日共申請" + (tseCount + otcCount) + "檔權證發行/增額");

                MessageBox.Show("轉TXT檔完成!\n" + infoStr);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }

        }

        private void 權證系統上傳檔ToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileName = "D:\\權證發行_相關Excel\\上傳檔\\權證發行匯入檔.xls";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;

            try {
                string sql = @"SELECT a.UnderlyingID
	                                  ,c.TraderID
                                      ,a.WarrantName
                                      ,c.Type
                                      ,c.CP
                                      ,IsNull(b.MPrice,0) MPrice
                                      ,c.K
                                      ,c.ResetR
                                      ,c.BarrierR
                                      ,c.T
                                      ,a.CR
                                      ,c.HV
                                      ,c.IV
                                      ,a.IssueNum
                                      ,c.FinancialR
                                      ,a.UseReward
                                      ,c.Apply1500W
                                      ,c.SerialNumber
                                  FROM [EDIS].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [EDIS].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID
                                  LEFT JOIN [EDIS].[dbo].[ApplyOfficial] c ON a.SerialNum=c.SerialNumber
                                  WHERE a.ApplyKind='1' AND a.Result+0.00001 >= a.EquivalentNum
                                  ORDER BY a.Market desc, a.Type, a.CP, a.UnderlyingID, a.SerialNum"; //a.SerialNum
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                if (dv.Rows.Count > 0) {
                    int i = 3;
                    app.Visible = true;
                    workBook = app.Workbooks.Open(fileName);
                    //workBook.EnvelopeVisible = false;
                    Worksheet workSheet = (Worksheet) workBook.Sheets[1];
                    workSheet.get_Range("A3:BZ1000").ClearContents();
                    //workSheet.UsedRange.

                    foreach (DataRow dr in dv.Rows) {
                        string date = DateTime.Today.ToString("yyyyMMdd");
                        string underlyingID = dr["UnderlyingID"].ToString();
                        string traderID = dr["TraderID"].ToString().TrimStart('0');
                        if (traderID == "8730" || traderID == "10120")
                            traderID = "7643";
                        string warrantName = dr["WarrantName"].ToString();
                        string type = dr["Type"].ToString();
                        string cp = dr["CP"].ToString();
                        string isReset = "N";
                        if (type == "重設型" || type == "牛熊證")
                            isReset = "Y";
                        double stockPrice = Convert.ToDouble(dr["MPrice"]);
                        double k = Convert.ToDouble(dr["K"]);
                        double resetR = Convert.ToDouble(dr["ResetR"]);
                        double barrierR = Convert.ToDouble(dr["BarrierR"]);
                        if (isReset == "Y")
                            k = Math.Round(resetR / 100 * stockPrice, 2);
                        double barrierP = Math.Round(barrierR / 100 * stockPrice, 2);
                        if (type == "牛熊證") {
                            if (cp == "C")
                                barrierP = Math.Round(Math.Floor(barrierR * stockPrice) / 100, 2);
                            else if (cp == "P")
                                barrierP = Math.Round(Math.Ceiling(barrierR * stockPrice) / 100, 2);
                        }

                        //Check for moneyness constraint
                        if (type != "牛熊證") {
                            if ((cp == "C" && k / stockPrice >= 1.5) || (cp == "P" && k / stockPrice <= 0.5)) {
                                MessageBox.Show(warrantName + " strike price is not valid due to moneyness constraint.");
                                // continue;
                            }
                        }

                        int t = Convert.ToInt32(dr["T"]);
                        double cr = Convert.ToDouble(dr["CR"]);
                        double r = GlobalVar.globalParameter.interestRate * 100;
                        double hv = Convert.ToDouble(dr["HV"]);
                        double iv = Convert.ToDouble(dr["IV"]);
                        double issueNum = Convert.ToDouble(dr["IssueNum"]);
                        double price = 0.0;
                        double financialR = Convert.ToDouble(dr["FinancialR"]);
                        string isReward = dr["UseReward"].ToString();

                        string is1500W = dr["Apply1500W"].ToString();
                        string serialNum = dr["SerialNumber"].ToString();
                        double p = 0.0;
                        double vol = iv / 100;
                        if (is1500W == "Y") {
                            CallPutType cpType = CallPutType.Call;
                            if (cp == "P")
                                cpType = CallPutType.Put;
                            else
                                cpType = CallPutType.Call;

                            if (type == "牛熊證")
                                p = Pricing.BullBearWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                            else if (type == "重設型")
                                p = Pricing.ResetWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                            else
                                p = Pricing.NormalWarrantPrice(cpType, stockPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);

                            double totalValue = p * issueNum * 1000;
                            double volUpperLimmit = vol * 2;
                            while (totalValue < 15000000 && vol < volUpperLimmit) {
                                vol += 0.01;
                                if (type == "牛熊證")
                                    p = Pricing.BullBearWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                                else if (type == "重設型")
                                    p = Pricing.ResetWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                else
                                    p = Pricing.NormalWarrantPrice(cpType, stockPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                totalValue = p * issueNum * 1000;
                            }

                            if (vol < volUpperLimmit) {
                                iv = vol * 100;
                                string cmdText = "UPDATE [ApplyOfficial] SET IVNew=@IVNew WHERE SerialNumber=@SerialNumber";
                                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                                pars.Add(new SqlParameter("@IVNew", SqlDbType.Float));
                                pars.Add(new SqlParameter("@SerialNumber", SqlDbType.VarChar));

                                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

                                h.SetParameterValue("@IVNew", iv);
                                h.SetParameterValue("@SerialNumber", serialNum);
                                h.ExecuteCommand();
                                h.Dispose();
                            }
                        }

                        if (type == "重設型")
                            type = "一般型";
                        if (cp == "P")
                            cp = "認售";
                        else
                            cp = "認購";
                        try {
                            // workSheet.Cells[1][i] = date;
                            workSheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value = date;
                            workSheet.get_Range("B" + i.ToString(), "B" + i.ToString()).Value = underlyingID;
                            workSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Value = traderID;
                            workSheet.get_Range("D" + i.ToString(), "D" + i.ToString()).Value = warrantName;
                            workSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Value = type;
                            workSheet.get_Range("F" + i.ToString(), "F" + i.ToString()).Value = cp;
                            workSheet.get_Range("G" + i.ToString(), "G" + i.ToString()).Value = isReset;
                            workSheet.get_Range("H" + i.ToString(), "H" + i.ToString()).Value = stockPrice;
                            workSheet.get_Range("I" + i.ToString(), "I" + i.ToString()).Value = k;
                            workSheet.get_Range("J" + i.ToString(), "J" + i.ToString()).Value = resetR;
                            workSheet.get_Range("K" + i.ToString(), "K" + i.ToString()).Value = barrierP;
                            workSheet.get_Range("L" + i.ToString(), "L" + i.ToString()).Value = barrierR;
                            workSheet.get_Range("M" + i.ToString(), "M" + i.ToString()).Value = t;
                            workSheet.get_Range("N" + i.ToString(), "N" + i.ToString()).Value = cr;
                            workSheet.get_Range("O" + i.ToString(), "O" + i.ToString()).Value = r;
                            workSheet.get_Range("P" + i.ToString(), "P" + i.ToString()).Value = hv;
                            workSheet.get_Range("Q" + i.ToString(), "Q" + i.ToString()).Value = iv;
                            workSheet.get_Range("R" + i.ToString(), "R" + i.ToString()).Value = issueNum;
                            workSheet.get_Range("S" + i.ToString(), "S" + i.ToString()).Value = price;
                            workSheet.get_Range("T" + i.ToString(), "T" + i.ToString()).Value = financialR;
                            workSheet.get_Range("Y" + i.ToString(), "Y" + i.ToString()).Value = isReward;
                            i++;
                        } catch (Exception ex) {
                            MessageBox.Show("write" + ex.Message);
                        }
                    }

                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + "產發行上傳檔");
                    app.Visible = false;
                    MessageBox.Show("發行上傳檔完成!");
                } else
                    MessageBox.Show("無可發行權證");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (workBook != null) {
                    workBook.Save();
                    workBook.Close();
                }
                if (app != null)
                    app.Quit();
            }
        }

        private void 增額上傳檔ToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileName = "D:\\權證發行_相關Excel\\上傳檔\\增額作業匯入資料.xls";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;
            bool noPrice = false;

            try {
                string sql = @"SELECT a.WarrantName
                                      ,a.IssueNum
                                      ,IsNull(c.MPrice,IsNull(c.BPrice,IsNull(c.APrice,0))) MPrice
                                      ,b.WarrantID
                                  FROM [EDIS].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [EDIS].[dbo].[ReIssueOfficial] b ON a.SerialNum=b.SerialNum
                                  LEFT JOIN [EDIS].[dbo].[WarrantPrices] c ON b.WarrantID=c.CommodityID
                                  WHERE a.ApplyKind='2' AND a.Result + 0.00001 >=a.EquivalentNum
                                  ORDER BY a.Market desc, a.SerialNum";
                //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                if (dv.Rows.Count > 0) {
                    int i = 3;
                    app.Visible = true;
                    workBook = app.Workbooks.Open(fileName);
                    //workBook.EnvelopeVisible = false;
                    Worksheet workSheet = (Worksheet) workBook.Sheets[1];
                    workSheet.get_Range("A3:Z1000").ClearContents();

                    foreach (DataRow dr in dv.Rows) {
                        double warrantPrice = Convert.ToDouble(dr["MPrice"]);
                        if (warrantPrice == 0.0)
                            noPrice = true;

                        workSheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value = dr["WarrantName"].ToString();
                        workSheet.get_Range("B" + i.ToString(), "B" + i.ToString()).Value = "權證增額";
                        workSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Value = DateTime.Today.ToString("yyyyMMdd");
                        workSheet.get_Range("D" + i.ToString(), "D" + i.ToString()).Value = "增額發行";
                        workSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Value = Convert.ToDouble(dr["IssueNum"]) * 1000;
                        workSheet.get_Range("F" + i.ToString(), "F" + i.ToString()).Value = warrantPrice;
                        i++;
                    }
                    if (noPrice)
                        MessageBox.Show("注意!有權證價格為零!");

                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + "產增額上傳檔");
                    app.Visible = false;
                    MessageBox.Show("增額上傳檔完成!");
                } else
                    MessageBox.Show("無可增額權證");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (workBook != null) {
                    workBook.Save();
                    workBook.Close();
                }
                if (app != null)
                    app.Quit();
            }
        }

        private void 關係人列表ToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileName = "D:\\權證發行_相關Excel\\上傳檔\\利害關係人整批查詢上傳格式範例.xls";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;

            try {
                string sql = @"SELECT DISTINCT a.UnderlyingID
                                      ,b.UnifiedID
                                      ,b.FullName
                                  FROM [EDIS].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [EDIS].[dbo].[WarrantUnderlying] b ON a.UnderlyingID=b.UnderlyingID
                                  WHERE a.Result>=a.EquivalentNum AND (b.StockType='DS' OR b.StockType='DR')";
                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                int i = 2;
                if (dv.Count > 0) {
                    app.Visible = true;
                    workBook = app.Workbooks.Open(fileName);
                    //workBook.EnvelopeVisible = false;
                    Worksheet workSheet = (Worksheet) workBook.Sheets[1];
                    workSheet.get_Range("A3:Z1000").ClearContents();

                    foreach (DataRowView dr in dv) {
                        workSheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value = i - 1;
                        workSheet.get_Range("B" + i.ToString(), "B" + i.ToString()).Value = dr["UnifiedID"].ToString();
                        workSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Value = dr["FullName"].ToString();
                        i++;
                    }

                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + "產關係人上傳檔");
                    app.Visible = false;
                    MessageBox.Show("關係人上傳檔完成!");
                } else
                    MessageBox.Show("無關係人需查詢");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (workBook != null) {
                    workBook.Save();
                    workBook.Close();
                }
                if (app != null)
                    app.Quit();
            }
        }

        private void 修正權證名稱ToolStripMenuItem_Click(object sender, EventArgs e) {
            string sql5 = "SELECT [SerialNum], [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1'";
            System.Data.DataTable dv = MSSQL.ExecSqlQry(sql5, GlobalVar.loginSet.edisSqlConnString);

            string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName WHERE SerialNum=@SerialNum";
            List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
            pars.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
            pars.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

            bool updated = false;
            foreach (DataRow dr in dv.Rows) {
                string serialNum = dr["SerialNum"].ToString();
                string warrantName = dr["WarrantName"].ToString();

                string sqlTemp = "select top (1) WarrantName from (SELECT [WarrantName] FROM [EDIS].[dbo].[WarrantBasic] WHERE SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 3) + "' union ";
                sqlTemp += " SELECT [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND [SerialNum]<" + serialNum + " AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 3) + "') as tb1 ";
                sqlTemp += " order by SUBSTRING(WarrantName,len(WarrantName)-1,len(WarrantName)) desc";

                System.Data.DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                int count = 0;
                if (dvTemp.Rows.Count > 0) {
                    string lastWarrantName = dvTemp.Rows[0][0].ToString();
                    if (!int.TryParse(lastWarrantName.Substring(lastWarrantName.Length - 2, 2), out count))
                        MessageBox.Show("parse failed " + lastWarrantName);
                }

                if (warrantName.Substring(warrantName.Length - 2, 2) != (count + 1).ToString("0#")) {
                    updated = true;
                    warrantName = warrantName.Substring(0, warrantName.Length - 2) + (count + 1).ToString("0#");
                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@SerialNum", serialNum);
                    h.ExecuteCommand();
                }            
            }
            h.Dispose();
            if (updated)
                MessageBox.Show("Magic!");
            else
                MessageBox.Show("No magic.");
        }
    }
}
