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
            GlobalUtility.Close();
        }

        private void RoutineWork() {
            try {
                for (; ; ) {
                    while (workQueue2.Count > 0) {
                        try {
                            workQueue2.TryDequeue(out WorkInQueue workInQueue);
                            if (workInQueue != null) {
                                WorkState workstate = workInQueue.Invoke();
                                //workInQueue.Method.Name
                                if (workstate == WorkState.Successful)
                                    AddMessage($"Complete Sucessfully\t\t{workInQueue.Method.Name}");
                                else if (workstate == WorkState.Exception)
                                    AddMessage($"Failed Due To Exception\t\t{workInQueue.Method.Name}");
                                else
                                    AddMessage($"Failed Due To Some Error\t\t{workInQueue.Method.Name}");
                            }
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
                for (; ; ) {
                    while (messageQueue2.Count > 0) {
                        try {
                            messageQueue2.TryDequeue(out string message);
                            if (message != "") {
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
            AddWork(DataCollect.UpdateWarrantUnderlying);
        }

        private void WarrantDataRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantBasicWork("WarrantDataRefresh", "權證資料更新"));
            AddWork(DataCollect.UpdateWarrantBasic);
        }

        private void IssueCreditRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingCreditWork("IssueCreditRefresh", "權證額度更新"));
            AddWork(DataCollect.UpdateWarrantUnderlyingCredit);
        }

        private void IssueCheckRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantIssueCheckWork("IssueCheckRefresh", "發行檢查更新"));
            AddWork(CMoneyData.LoadCMoneyData);
        }

        private void SummaryRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingSummaryWork("SummaryRefresh", "Summary更新"));
            AddWork(DataCollect.UpdateWarrantUnderlyingSummary);
        }

        private void PricesRefresh_Click(object sender, EventArgs e) {
            //AddWork(new WarrantPricesWork("PricesRefresh", "價格更新"));
            AddWork(DataCollect.UpdateWarrantPrices);
        }

        private void UpdateAll_Click(object sender, EventArgs e) {
            //AddWork(new WarrantUnderlyingWork("UnderlyingDataRefresh", "標的資料更新"));
            //AddWork(new WarrantBasicWork("WarrantDataRefresh", "權證資料更新"));
            //AddWork(new WarrantUnderlyingCreditWork("IssueCreditRefresh", "權證額度更新"));
            //AddWork(new WarrantIssueCheckWork("IssueCheckRefresh", "發行檢查更新"));
            //AddWork(new WarrantUnderlyingSummaryWork("SummaryRefresh", "Summary更新"));
            //AddWork(new WarrantPricesWork("PricesRefresh", "價格更新"));
            AddWork(DataCollect.UpdateWarrantUnderlying);
            AddWork(DataCollect.UpdateWarrantBasic);
            AddWork(DataCollect.UpdateWarrantUnderlyingCredit);
            AddWork(CMoneyData.LoadCMoneyData);
            AddWork(DataCollect.UpdateWarrantUnderlyingSummary);
            AddWork(DataCollect.UpdateWarrantPrices);
        }

        private void CleanApplyList_Click(object sender, EventArgs e) {
            //AddWork(new CleanApplyList("CleanApplyList", "申請表清空"));
            AddWork(DataCollect.UpdateApplyLists);
        }


    }
}
