using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace WarrantDataManager
{    
    public class SQLCommandHelper : IDisposable
    {
        public string connString = "";
        public string commandString = "";
        public List<SqlParameter> parameterList;

        private bool isInitialOK = false;
        private SqlCommand cmd;

        public SQLCommandHelper(string connString, string commandString, List<SqlParameter> parameterList)
        {
            this.connString = connString;
            this.commandString = commandString;
            this.parameterList = parameterList;
            InitialCommand();
        }
        private void InitialCommand()
        {
            try
            {
                SqlConnection conn = new SqlConnection();
                conn.ConnectionString = connString;
                cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = commandString;
                foreach (SqlParameter parameter in parameterList)
                    cmd.Parameters.Add(parameter);
                isInitialOK = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("[" + DateTime.Now.ToString("HH:mm:ss.fff") + "][SQLCommandHelper_InitialCommand][" + ex.Message + "]");
            }
        }
        public void SetParameterValue(string parameterName, object value)
        {
            try
            {
                if (value != null)
                    cmd.Parameters[parameterName].Value = value;
                else
                    cmd.Parameters[parameterName].Value = DBNull.Value;
            }
            catch (Exception ex)
            {
                //GlobalVar.errProcess.Add(1, "[SqlCommandHelper_SetParameterValue][" + ex.Message + "][" + ex.StackTrace + "]");
            }
        }
        public void ExecuteCommand()
        {
            try
            {
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"[{DateTime.Now.ToString("HH:mm:ss.fff")}][SQLCommandHelper][{ex.Message}]  ");
                string temp = "";
                for(int i = 0; i < cmd.Parameters.Count; i++) 
                    temp += $"{cmd.Parameters[i].Value} ";
                
                MessageBox.Show($"{cmd.CommandText} {temp}");
                
            }
        }
        #region IDisposable成員
        public void Dispose()
        {
            if (cmd != null && cmd.Connection != null && cmd.Connection.State != ConnectionState.Closed)
            {
                cmd.Connection.Close();
                cmd.Connection.Dispose();
                cmd.Dispose();
            }
        }
        #endregion
    }
}
