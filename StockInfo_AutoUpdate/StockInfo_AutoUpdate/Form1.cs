using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace StockInfo_AutoUpdate
{
    public partial class Form1 : Form
    {
        private string m_sourcePath = @"R:\工具\WarrantAssistant";
        private string m_destPath = @"D:\WarrantAssistant";
        
        public Form1()
        {
            InitializeComponent();
        }

        private void UpdateFiles()
        {
            timer1.Enabled = false;

            if (System.IO.Directory.Exists(m_destPath) == false)
                Directory.CreateDirectory(m_destPath);

            if(System.IO.Directory.Exists(m_sourcePath) == false)
            {
                RunLocal();
                this.Close();
            }

            int iCount = 0;
            string[] arrFile = Directory.GetFiles(this.m_sourcePath);
            foreach(string strFile in arrFile)
            {
                FileInfo fiSource = new FileInfo(strFile);

                bool needUpdate = true;
                if(File.Exists(m_destPath + "\\" + fiSource.Name))
                {
                    FileInfo fiDest = new FileInfo(m_destPath + "\\" + fiSource.Name);
                    if(fiSource.LastWriteTime <= fiDest.LastWriteTime)
                        needUpdate = false;
                }

                if(needUpdate)
                    File.Copy(strFile, m_destPath + "\\" + fiSource.Name, true);

                iCount++;
                progressBar1.Value = Convert.ToInt32(iCount * 1.0 / arrFile.Length * 100);
                Application.DoEvents();
            }

            RunLocal();
            this.Close();
        }

        private void RunLocal()
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(m_destPath + "\\" + "WarrantAssistant.exe"));
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            UpdateFiles();
        }
    }
}
