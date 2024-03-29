﻿using System;
using System.Data;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmLogIn:Form
    {
        public bool loginOK = false;
        public FrmLogIn() {
            InitializeComponent();
        }

        public bool TryIPLogin() {
            string IP = EDLib.Utility.GetLocalIPAddress();
            string sqlTemp = $"SELECT [UserGroup],[UserLevel],[UserName],[Deputy],[UserID] FROM [EDIS].[dbo].[User] WHERE IP = '{IP}'";
            //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp , GlobalVar.loginSet.edisSqlConnString);
            DataView dvTemp = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString).DefaultView;
            if (dvTemp.Count > 0) {
                foreach (DataRowView drTemp in dvTemp) {
                    GlobalVar.globalParameter.userGroup = drTemp["UserGroup"].ToString();
                    GlobalVar.globalParameter.userLevel = drTemp["UserLevel"].ToString();
                    GlobalVar.globalParameter.userName = drTemp["UserName"].ToString();
                    GlobalVar.globalParameter.userDeputy = drTemp["Deputy"].ToString();
                    GlobalVar.globalParameter.userID = drTemp["UserID"].ToString();
                }
                loginOK = true;
                GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " login.(IP)");
                this.Close();
                GlobalVar.mainForm.Start();
                return true;
            }
            GlobalUtility.LogInfo("Log", IP + " login failed.");
            return false;
        }

        private void Check() {
            string account = textBox1.Text;
            string password = textBox2.Text;

            if (account != "" && password != "") { // Login with UserID               
                string sqlTemp = "SELECT [UserGroup],[UserLevel],[UserName],[Deputy] FROM [EDIS].[dbo].[User] WHERE UserID = '" + account + "' and [UserPasswordEncrypt] = HASHBYTES('SHA1', '" + password + "')";

                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                if (dvTemp.Count > 0) {
                    foreach (DataRowView drTemp in dvTemp) {
                        GlobalVar.globalParameter.userGroup = drTemp["UserGroup"].ToString();
                        GlobalVar.globalParameter.userLevel = drTemp["UserLevel"].ToString();
                        GlobalVar.globalParameter.userName = drTemp["UserName"].ToString();
                        GlobalVar.globalParameter.userDeputy = drTemp["Deputy"].ToString();
                    }
                    GlobalVar.globalParameter.userID = account;
                    loginOK = true;
                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " login.(ID)");
                    this.Close();
                    GlobalVar.mainForm.Start();
                } else {
                    MessageBox.Show("帳號密碼錯誤!");
                    textBox2.Text = "";
                    textBox1.Text = "";
                    textBox1.Focus();
                    GlobalUtility.LogInfo("Log", account + " login failed." + password);
                    Thread.Sleep(5000);
                    return;
                }

                /*if (password == passwordS) {
                    GlobalVar.globalParameter.userID = account;
                    GlobalVar.globalParameter.userDeputy = deputy;
                    GlobalVar.globalParameter.userGroup = group;
                    GlobalVar.globalParameter.userLevel = level;
                    GlobalVar.globalParameter.userName = name;
                    loginOK = true;
                    this.Close();
                    GlobalVar.mainForm.Start();
                } else {
                    MessageBox.Show("密碼錯誤!");
                    textBox2.Text = "";
                    textBox2.Focus();
                    GlobalUtility.logInfo("Log" , account + " login failed."+password);
                    return;
                }*/
            } else {
                MessageBox.Show("請輸入帳號密碼");
                textBox2.Text = "";
                textBox1.Focus();
                return;
            }

        }

        private void button1_Click(object sender, EventArgs e) {
            Check();
        }

        private void FrmLogIn_Load(object sender, EventArgs e) {
            textBox1.Focus();
        }

        private void FrmLogIn_FormClosed(object sender, FormClosedEventArgs e) {

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter)
                Check();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter)
                Check();
        }

        private void FrmLogIn_Shown(object sender, EventArgs e) {
            textBox1.Focus();
        }
    }
}
