using System;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop;
//using Oracle.DataAccess.Client;

namespace WarrantDataManager2._0
{
    public enum WorkState { Successful = 0, Exception = 1, Failed = 2 }

    public class Work
    {
        public string workID = "";
        public string workName = "";
        public DateTime doWorkTime;

        public Work(string workID, string workName)
        {
            this.workID = workID;
            this.workName = workName;
        }

        public virtual WorkState DoWork() { return WorkState.Successful; }
        public virtual void Close() { }
    }

    public class WarrantUnderlyingWork : Work
    {
        private DataCollect dataCollect;
        public WarrantUnderlyingWork(string workID, string workName)
            : base(workID, workName)
        {
            dataCollect = new DataCollect();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (dataCollect != null)
                {
                    dataCollect.updateWarrantUnderlying();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[IssuableUnderlyingDataWork_DoWork][" + ex.Message + "][" + ex.StackTrace + "]");
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
  
    }

    public class WarrantBasicWork : Work
    {
        private DataCollect dataCollect;
        public WarrantBasicWork(string workID, string workName)
            : base(workID, workName)
        {
            dataCollect = new DataCollect();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (dataCollect != null)
                {
                    dataCollect.updateWarrantBasic();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[IssuableUnderlyingDataWork_DoWork][" + ex.Message + "][" + ex.StackTrace + "]");
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }

    public class WarrantUnderlyingCreditWork : Work
    {
        private DataCollect dataCollect;
        public WarrantUnderlyingCreditWork(string workID, string workName)
            : base(workID, workName)
        {
            dataCollect = new DataCollect();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (dataCollect != null)
                {
                    dataCollect.updateWarrantUnderlyingCredit();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //GlobalVar.errProcess.Add(1, "[IssuableUnderlyingDataWork_DoWork][" + ex.Message + "][" + ex.StackTrace + "]");
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }

    public class WarrantIssueCheckWork : Work
    {
        private CMoneyData cMoneyData;
        public WarrantIssueCheckWork(string workID, string workName)
            : base(workID, workName)
        {
            this.cMoneyData = new CMoneyData();
        }

        public override WorkState DoWork()
        {
            cMoneyData.loadData();
            return WorkState.Successful;
        }

        public override void Close()
        {
            if (cMoneyData != null)
                cMoneyData.Dispose();
        }
    }

    public class WarrantUnderlyingSummaryWork : Work
    {
        private DataCollect dataCollect;
        public WarrantUnderlyingSummaryWork(string workID, string workName)
            : base(workID, workName)
        {
            dataCollect = new DataCollect();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (dataCollect != null)
                {
                    dataCollect.updateWarrantUnderlyingSummary();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[IssuableUnderlyingDataWork_DoWork][" + ex.Message + "][" + ex.StackTrace + "]");
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }

    public class WarrantPricesWork : Work
    {
        private DataCollect dataCollect;
        public WarrantPricesWork(string workID, string workName)
            : base(workID, workName)
        {
            dataCollect = new DataCollect();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (dataCollect != null)
                {
                    dataCollect.updateWarrantPrices();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[IssuableUnderlyingDataWork_DoWork][" + ex.Message + "][" + ex.StackTrace + "]");
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }

    public class CleanApplyList : Work
    {
        private DataCollect dataCollect;
        public CleanApplyList(string workID, string workName)
            : base(workID, workName)
        {
            dataCollect = new DataCollect();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (dataCollect != null)
                {
                    dataCollect.updateApplyLists();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[IssuableUnderlyingDataWork_DoWork][" + ex.Message + "][" + ex.StackTrace + "]");
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }
}
