using System;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Concurrent;

namespace WarrantDataManager
{
    public enum WorkState { Successful = 0, Exception = 1, Failed = 2 }
    public partial class MainForm:Form
    {
        private delegate void ShowHandler(string message);
        public delegate WorkState WorkInQueue();

        //private SafeQueue workQueue = new SafeQueue();
        //private SafeQueue messageQueue = new SafeQueue();
        private ConcurrentQueue<string> messageQueue2 = new ConcurrentQueue<string>();
        private ConcurrentQueue<WorkInQueue> workQueue2 = new ConcurrentQueue<WorkInQueue>();

        private Thread workThread;
        private Thread msgThread;

        public void AddWork(WorkInQueue work) {
            //workQueue.Enqueue(work);
            workQueue2.Enqueue(work);
        }

        public void AddMessage(string message) {
            //messageQueue.Enqueue(message);
            messageQueue2.Enqueue(message);
        }

        public MainForm() {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e) {
            try {
                GlobalVar.mainForm = this;
                workThread = new Thread(new ThreadStart(RoutineWork));
                msgThread = new Thread(new ThreadStart(MessageWork));
                workThread.Start();
                msgThread.Start();

            } catch (Exception ex) {
                MessageBox.Show("MainForm_Load][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e) {
            if (workThread != null && workThread.IsAlive) { workThread.Abort(); }
            if (msgThread != null && msgThread.IsAlive) { msgThread.Abort(); }
            GlobalUtility.close();
        }

        private void RoutineWork() {
            try {
                for (;;) {
                    while (workQueue2.Count > 0) {
                        try {
                            WorkInQueue workInQueue = null;
                            workQueue2.TryDequeue(out workInQueue);
                            if (workInQueue != null) {
                                WorkState workstate = workInQueue.Invoke();
                                //workInQueue.Method.Name
                                if (workstate == WorkState.Successful)
                                    AddMessage("Work[" + workInQueue.Method.Name + "]" + "\t\t" + "Complete Sucessfully");
                                else if (workstate == WorkState.Exception)
                                    AddMessage("Work[" + workInQueue.Method.Name + "]" + "\t\t" + "Failed Due To Exception");
                                else
                                    AddMessage("Work[" + workInQueue.Method.Name + "]" + "\t\t" + "Failed Due To Some Error");
                            }
                            /*object obj = workQueue.Dequeue();
                            if (obj != null) {
                                Work work = (Work) obj;
                                WorkState workstate = work.DoWork();
                                if (workstate == WorkState.Successful)
                                    AddMessage("Work[" + work.workName + "]" + "\t\t" + "Complete Sucessfully");
                                else if (workstate == WorkState.Exception)
                                    AddMessage("Work[" + work.workName + "]" + "\t\t" + "Failed Due To Exception");
                                else
                                    AddMessage("Work[" + work.workName + "]" + "\t\t" + "Failed Due To Some Error");
                                work.Close();
                            }*/
                        } catch (ThreadAbortException tex) {
                            MessageBox.Show(tex.Message);
                        } catch (Exception ex) {
                            MessageBox.Show(ex.Message);
                        }

                    }
                    Thread.Sleep(1000);
                }
            } catch (Exception ex) {
            }
        }

        private void MessageWork() {
            try {
                for (;;) {
                    //while (messageQueue.Count > 0) {
                    while (messageQueue2.Count > 0) {
                        try {
                            string message = "";
                            messageQueue2.TryDequeue(out message);
                            //object obj = messageQueue.Dequeue();
                            if (message != "") {
                                //string message = obj.ToString();
                                if (this.InvokeRequired)
                                    this.BeginInvoke(new ShowHandler(PublicMessage), new object[] { message });
                                else
                                    PublicMessage(message);
                            }
                        } catch (ThreadAbortException tex) {
                        } catch (Exception ex) {
                            //GlobalVar.errProcess.Add(1, "MainForm_MessageWork][" + ex.Message + "][" + ex.StackTrace + "]");
                        }

                    }
                    Thread.Sleep(1000);
                }
            } catch (Exception ex) {
            }
        }

        private void PublicMessage(string message) {
            try {
                if (listBox1.Items.Count > 1000)
                    listBox1.Items.Clear();
                listBox1.Items.Insert(0, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\t" + message);
            } catch (Exception ex) {
                //GlobalVar.errProcess.Add(1, "[MainForm_PublicMessage][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }

        private void UnderlyingDataRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingWork("UnderlyingDataRefresh", "標的資料更新"));
            AddWork(DataCollect.updateWarrantUnderlying);
        }

        private void WarrantDataRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantBasicWork("WarrantDataRefresh", "權證資料更新"));
            AddWork(DataCollect.updateWarrantBasic);
        }

        private void IssueCreditRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingCreditWork("IssueCreditRefresh", "權證額度更新"));
            AddWork(DataCollect.updateWarrantUnderlyingCredit);
        }

        private void IssueCheckRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantIssueCheckWork("IssueCheckRefresh", "發行檢查更新"));
            AddWork(CMoneyData.loadData);
        }

        private void SummaryRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingSummaryWork("SummaryRefresh", "Summary更新"));
            AddWork(DataCollect.updateWarrantUnderlyingSummary);
        }

        private void PricesRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantPricesWork("PricesRefresh", "價格更新"));
            AddWork(DataCollect.updateWarrantPrices);
        }

        private void UpdateAll_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingWork("UnderlyingDataRefresh", "標的資料更新"));
            //AddWork(new WarrantBasicWork("WarrantDataRefresh", "權證資料更新"));
            //AddWork(new WarrantUnderlyingCreditWork("IssueCreditRefresh", "權證額度更新"));
            //AddWork(new WarrantIssueCheckWork("IssueCheckRefresh", "發行檢查更新"));
            //AddWork(new WarrantUnderlyingSummaryWork("SummaryRefresh", "Summary更新"));
            //AddWork(new WarrantPricesWork("PricesRefresh", "價格更新"));
            AddWork(DataCollect.updateWarrantUnderlying);
            AddWork(DataCollect.updateWarrantBasic);
            AddWork(DataCollect.updateWarrantUnderlyingCredit);
            AddWork(CMoneyData.loadData);
            AddWork(DataCollect.updateWarrantUnderlyingSummary);
            AddWork(DataCollect.updateWarrantPrices);
        }

        private void CleanApplyList_Click(object sender, EventArgs e) {
            //AddWork(new CleanApplyList("CleanApplyList", "申請表清空"));
            AddWork(DataCollect.updateApplyLists);
        }


    }
}
