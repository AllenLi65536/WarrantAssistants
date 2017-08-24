using System;
using System.Windows.Forms;
using System.IO;

namespace StockInfo_AutoUpdate
{
    public partial class Form1:Form
    {
        private string sourcePath = @"R:\工具\WarrantAssistant";
        private string destPath = @"D:\WarrantAssistant";

        public Form1() {
            InitializeComponent();
        }

        private void UpdateFiles() {
            timer1.Enabled = false;

            if (Directory.Exists(destPath) == false)
                Directory.CreateDirectory(destPath);

            if (Directory.Exists(sourcePath) == false) {
                RunLocal();
                Close();
            }

            int iCount = 0;
            string[] sourceFiles = Directory.GetFiles(this.sourcePath);
            foreach (string sourceFile in sourceFiles) {
                FileInfo fiSource = new FileInfo(sourceFile);

                bool needUpdate = true;
                if (File.Exists(destPath + "\\" + fiSource.Name)) {
                    FileInfo fiDest = new FileInfo(destPath + "\\" + fiSource.Name);
                    if (fiSource.LastWriteTime <= fiDest.LastWriteTime)
                        needUpdate = false;
                }

                if (needUpdate)
                    File.Copy(sourceFile, destPath + "\\" + fiSource.Name, true);

                iCount++;
                progressBar1.Value = Convert.ToInt32(iCount * 1.0 / sourceFiles.Length * 100);
                Application.DoEvents();
            }

            RunLocal();
            this.Close();
        }

        private void RunLocal() {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(destPath + "\\" + "WarrantAssistant.exe"));
        }

        private void timer1_Tick(object sender, EventArgs e) {
            UpdateFiles();
        }
    }
}
