using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop;
using System.Data.SqlClient;
using System.Data;
using System.IO;
//using Oracle.DataAccess.Client;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Configuration;
using System.ComponentModel;
using System.Threading;
using System.Drawing;

namespace WarrantAssistant
{
    public enum WorkState { Successful = 0, Exception = 1, Failed = 2 }

    public class Work
    {
        public string workName = "";
        public DateTime doWorkTime;

        public Work(string workName)
        {
            this.workName = workName;
        }

        public virtual WorkState DoWork() { return WorkState.Successful; }
        public virtual void Close() { }
    }

    public class InfoWork : Work
    {
        //private MainForm mainform;
        public InfoWork(string workName)
            : base(workName)
        {
            //mainform = new MainForm();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (GlobalVar.mainForm != null)
                {
                    //GlobalVar.mainForm.SetUltraGrid1();
                    GlobalVar.mainForm.LoadUltraGrid1();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                GlobalUtility.logInfo("Error", "InfoWork Error: " + ex.Message);
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Error");
                hInfo.SetParameterValue("@InformationContent", "InfoWork Error: "+ex.Message);
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/

                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }

    public class AnnounceWork : Work
    {
        //private MainForm mainform;
        public AnnounceWork(string workName)
            : base(workName)
        {
            //mainform = new MainForm();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (GlobalVar.mainForm != null)
                {
                    //GlobalVar.mainForm.SetUltraGrid2();
                    GlobalVar.mainForm.LoadUltraGrid2();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                GlobalUtility.logInfo("Error", "AnnounceWork Error: " + ex.Message);
                /*string sqlInfo = "INSERT INTO [InformationLog] ([MDate],[InformationType],[InformationContent],[MUser]) values(@MDate, @InformationType, @InformationContent, @MUser)";
                List<SqlParameter> psInfo = new List<SqlParameter>();
                psInfo.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                psInfo.Add(new SqlParameter("@InformationType", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@InformationContent", SqlDbType.VarChar));
                psInfo.Add(new SqlParameter("@MUser", SqlDbType.VarChar));

                SQLCommandHelper hInfo = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, sqlInfo, psInfo);
                hInfo.SetParameterValue("@MDate", DateTime.Now);
                hInfo.SetParameterValue("@InformationType", "Error");
                hInfo.SetParameterValue("@InformationContent", "AnnounceWork Error: " + ex.Message);
                hInfo.SetParameterValue("@MUser", GlobalVar.globalParameter.userID);
                hInfo.ExecuteCommand();
                hInfo.Dispose();*/

                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }
}
