using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using EDLib.SQL;
using HtmlAgilityPack;
using System.IO;

namespace WarrantAssistant
{
    public partial class FrmRename:Form
    {
        public FrmRename() {
            InitializeComponent();
        }
        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            e.Layout.Override.CellAppearance.BackColor = Color.LightCyan;           
            //e.Layout.Bands[0].Columns["WName"]
            //e.Layout.Bands[0].Columns["WName"].CellAppearance.ForeColor = Color.Gray;           
        }
        private void FrmRename_Load(object sender, EventArgs e) {
            toolStripLabel1.Text = "";

            //Get key and id
            DataTable dv = MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='10'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string key = dv.Rows[0]["FLGDAT_FLGDTA"].ToString();

            dv = MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='20'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string id = dv.Rows[0]["FLGDAT_FLGDTA"].ToString();

            string twseUrl = "http://siis.twse.com.tw/server-java/t150sa03?step=0&id=9200pd" + id + "&TYPEK=sii&key=" + key;

            if (!ParseHtml(twseUrl))
                return;
        }

        private bool ParseHtml(string url) {
            try {
                string firstResponse = GlobalUtility.GetHtml(url);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(firstResponse);
                HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]/tr[1]/td/table").ChildNodes;

                int count = 0;
                for (int i = 3; i < navNodeChild.Count; i += 2) {
                    string[] split = navNodeChild[i].InnerText.Split(new string[] { " ", "\t", "&nbsp;", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    ultraDataSource1.Rows.Add();
                    ultraDataSource1.Rows[count]["WName"] = split[1];
                    ultraDataSource1.Rows[count++]["SerialNumber"] = split[0];
                }

                return true;
            } catch (Exception e) {
                MessageBox.Show("可能要更新Key");
                //MessageBox.Show(e.Message);
                return false;
            }
        }

        private void RenameFiles_Click(object sender, EventArgs e) {
            string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (Directory.Exists("C:\\WarrantDocuments\\Renamed" + now))
                Directory.Delete("C:\\WarrantDocuments\\Renamed" + now);
            Directory.CreateDirectory("C:\\WarrantDocuments\\Renamed" + now);

            for (int i = 0; i < ultraDataSource1.Rows.Count; i++)
                if (Directory.Exists("C:\\WarrantDocuments\\" + ultraDataSource1.Rows[i]["WName"])) {
                    string[] files = Directory.GetFiles("C:\\WarrantDocuments\\" + ultraDataSource1.Rows[i]["WName"]);
                    foreach (string file in files) {
                        if (Path.GetExtension(file).ToLower() != ".xml")
                            File.Copy(file,
                                 "C:\\WarrantDocuments\\Renamed" + now + "\\" + ultraDataSource1.Rows[i]["SerialNumber"] + "-" + Path.GetFileName(file), true);
                    }
                }
            toolStripLabel1.Text = DateTime.Now + "修改發行檔名完成";
        }

        private void RenameXML_Click(object sender, EventArgs e) {
            string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (Directory.Exists("C:\\WarrantDocuments\\Xml" + now))
                Directory.Delete("C:\\WarrantDocuments\\Xml" + now);
            Directory.CreateDirectory("C:\\WarrantDocuments\\Xml" + now);

            for (int i = 0; i < ultraDataSource1.Rows.Count; i++)
                if (Directory.Exists("C:\\WarrantDocuments\\" + ultraDataSource1.Rows[i]["WName"])) {
                    string[] files = Directory.GetFiles("C:\\WarrantDocuments\\" + ultraDataSource1.Rows[i]["WName"]);
                    foreach (string file in files) {
                        if (Path.GetExtension(file).ToLower() == ".xml")
                            File.Copy(file,
                                "C:\\WarrantDocuments\\Xml" + now + "\\" + ultraDataSource1.Rows[i]["SerialNumber"] + ".xml", true);
                    }
                }
            toolStripLabel1.Text = DateTime.Now + "修改XML檔名完成";
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e) {

        }
    }
}
