using System;
using System.Threading;

namespace WarrantDataManager
{
    public class AutoWork:IDisposable
    {
        private Thread workThread;

        private bool globalDataOK = true;
        private bool underlyingDataOK = true;
        private bool warrantDataOK = true;
        private bool underlyingCreditAfterOK = true;
        private bool underlyingCreditBeforeOK = true;
        private bool issueCheckOK = true;
        private bool summaryOK = true;
        private bool pricesOK = true;
        private bool cleanApplyOK = true;

        public AutoWork() {
            workThread = new Thread(new ThreadStart(Working));
            workThread.Start();
        }

        private void Working() {
            try {
                for (;;) {
                    try {
                        DateTime now = DateTime.Now;

                        //if (now.TimeOfDay.TotalSeconds > 10 && now.TimeOfDay.TotalSeconds < 30)
                        if (now.TimeOfDay.TotalMinutes > 424 && now.TimeOfDay.TotalMinutes < 425) {
                            globalDataOK = false;
                            //cleanApplyOK = false;
                        }

                        //0610
                        if (now.TimeOfDay.TotalMinutes > 430 && now.TimeOfDay.TotalMinutes < 431) {
                            underlyingCreditAfterOK = false;
                        }

                        //早上0730
                        if (now.TimeOfDay.TotalMinutes > 449 && now.TimeOfDay.TotalMinutes < 450) {
                            underlyingDataOK = false;
                            warrantDataOK = false;
                            underlyingCreditBeforeOK = false;
                            issueCheckOK = false;
                            summaryOK = false;
                        }

                        //早上八點
                        if (now.TimeOfDay.TotalMinutes > 478 && now.TimeOfDay.TotalMinutes < 479) {
                            underlyingDataOK = false;
                            warrantDataOK = false;
                            underlyingCreditBeforeOK = false;
                            issueCheckOK = false;
                            summaryOK = false;
                        }

                        //早上九點
                        if (now.TimeOfDay.TotalMinutes > 540 && now.TimeOfDay.TotalMinutes < 541) {
                            underlyingDataOK = false;
                            warrantDataOK = false;
                            underlyingCreditBeforeOK = false;
                            issueCheckOK = false;
                            summaryOK = false;
                            pricesOK = false;
                        }
                        if (now.TimeOfDay.TotalMinutes > 1170 && now.TimeOfDay.TotalMinutes < 1171)
                            cleanApplyOK = false;

                        /*價格更新頻率*/

                        if (now.TimeOfDay.TotalMinutes > 553 && now.TimeOfDay.TotalMinutes < 554)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 568 && now.TimeOfDay.TotalMinutes < 569)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 638 && now.TimeOfDay.TotalMinutes < 639)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 643 && now.TimeOfDay.TotalMinutes < 644)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 648 && now.TimeOfDay.TotalMinutes < 649)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 653 && now.TimeOfDay.TotalMinutes < 654)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 658 && now.TimeOfDay.TotalMinutes < 659)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 810 && now.TimeOfDay.TotalMinutes < 811)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 812 && now.TimeOfDay.TotalMinutes < 813)
                            pricesOK = false;

                        if (now.TimeOfDay.TotalMinutes > 814 && now.TimeOfDay.TotalMinutes < 815)
                            pricesOK = false;

                        /*價格更新頻率End*/

                        //if (now.TimeOfDay.TotalSeconds > 60 && (!globalDataOK))
                        if (now.TimeOfDay.TotalMinutes > 425 && (!globalDataOK)) {
                            GlobalUtility.LoadGlobalParameters();
                            globalDataOK = true;
                        }

                        //if (now.TimeOfDay.TotalSeconds > 60 && (!cleanApplyOK) && GlobalVar.globalParameter.isTodayTradeDate)
                        if (now.TimeOfDay.TotalMinutes > 1171 && (!cleanApplyOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new CleanApplyList("CleanApplyList", "申請表清空"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateApplyLists);
                            cleanApplyOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 450 && (!underlyingDataOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantUnderlyingWork("UnderlyingDataRefresh", "標的資料更新"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateWarrantUnderlying);
                            underlyingDataOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 450 && (!warrantDataOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantBasicWork("WarrantDataRefresh", "權證資料更新"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateWarrantBasic);
                            warrantDataOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 450 && (!underlyingCreditBeforeOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantUnderlyingCreditWork("IssueCreditRefresh", "權證額度更新"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateWarrantUnderlyingCredit);
                            underlyingCreditBeforeOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 450 && (!issueCheckOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantIssueCheckWork("IssueCheckRefresh", "發行檢查更新"));
                            GlobalVar.mainForm.AddWork(CMoneyData.LoadCMoneyData);
                            issueCheckOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 450 && (!summaryOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantUnderlyingSummaryWork("SummaryRefresh", "Summary更新"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateWarrantUnderlyingSummary);
                            summaryOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 540 && (!pricesOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantPricesWork("PricesRefresh", "價格更新"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateWarrantPrices);
                            pricesOK = true;
                        }

                        if (now.TimeOfDay.TotalMinutes > 431 && (!underlyingCreditAfterOK) && GlobalVar.globalParameter.isTodayTradeDate) {
                            //GlobalVar.mainForm.AddWork(new WarrantUnderlyingCreditWork("IssueCreditRefresh", "權證額度更新"));
                            GlobalVar.mainForm.AddWork(DataCollect.UpdateWarrantUnderlyingCredit);
                            underlyingCreditAfterOK = true;
                        }

                    } catch (ThreadAbortException) {
                    } catch (Exception) {
                    }
                    Thread.Sleep(20000);
                }
            } catch (ThreadAbortException) {
            } catch (Exception) {
            }
        }

        #region IDisposable成員

        public void Dispose() {
            if (workThread != null && workThread.IsAlive) { workThread.Abort(); }
        }

        #endregion
    }
}
