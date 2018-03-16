using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using EDLib.SQL;
using HtmlAgilityPack;
using System.IO;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading.Tasks;

namespace WarrantAssistant
{
    public partial class FrmRename:Form
    {
        public FrmRename() {
            InitializeComponent();
        }
        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            e.Layout.Override.CellAppearance.BackColor = Color.LightCyan;            
        }

        string path;// = "C:\\WarrantDocuments";

        private void FrmRename_Load(object sender, EventArgs e) {
            toolStripLabel1.Text = "";
            toolStrip1.Enabled = false;
            
            //Get key and id
            DataTable dv = MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='10'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string key = dv.Rows[0]["FLGDAT_FLGDTA"].ToString();

            dv = MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDTA FROM EDAISYS.dbo.V_FLAGDATAS WHERE FLGDAT_FLGNAM = 'WRT_ISSUE_QUOTA' and FLGDAT_ORDERS='20'"
                , GlobalVar.loginSet.warrantSysKeySqlConnString);
            string id = dv.Rows[0]["FLGDAT_FLGDTA"].ToString();                      

            string twseUrl = $"http://siis.twse.com.tw/server-java/t150sa03?step=0&id=9200pd{id}&TYPEK=sii&key={key}";

            if (!ParseHtml(twseUrl, true))
                return;

            twseUrl = $"http://siis.twse.com.tw/server-java/o_t150sa03?step=0&id=9200pd{id}&TYPEK=otc&key={key}";
            ParseHtml(twseUrl, false);

            // Get document path
            dv = MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDSC FROM EDAISYS.dbo.FLAGDATAS WHERE FLGDAT_FLGNAM = 'DOCUMENT_SETTING' and FLGDAT_FLGVAR='EXPORT_PATH'"
               , GlobalVar.loginSet.warrantSysKeySqlConnString);
            path = dv.Rows[0]["FLGDAT_FLGDSC"].ToString().TrimEnd('\\');
            
            toolStrip1.Enabled = true;
        }

        private string GetHtmlAsync(string url, Encoding encode) {
            try {
                using (WebResponse resp = WebRequest.Create(url).GetResponse()) // GetResponseAsync()
                using (Stream dataStream = resp.GetResponseStream())
                using (StreamReader reader = new StreamReader(dataStream, encode))
                    return reader.ReadToEnd();

            } catch (Exception err) {
                Console.WriteLine(err);
                return err.ToString();
            }
        }

        private bool ParseHtml(string url, bool twse) {
            try {
                string firstResponse = GetHtmlAsync(url, System.Text.Encoding.Default);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(firstResponse);
                HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]/tr[1]/td/table").ChildNodes;
                                               
                for (int i = 3; i < navNodeChild.Count; i += 2) {
                    string[] split = navNodeChild[i].InnerText.Split(new string[] { " ", "\t", "&nbsp;", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    ultraDataSource1.Rows.Add();
                    ultraDataSource1.Rows[ultraDataSource1.Rows.Count - 1]["WName"] = split[1];
                    ultraDataSource1.Rows[ultraDataSource1.Rows.Count - 1]["SerialNumber"] = split[0];
                    ultraDataSource1.Rows[ultraDataSource1.Rows.Count - 1]["Market"] = twse ? "TWSE" : "OTC";
                }

                return true;
            } catch (Exception e) {
                MessageBox.Show("可能要更新Key，或是還沒有資料");
                //MessageBox.Show(e.Message);
                return false;
            }
        }

        private void RenameFiles_Click(object sender, EventArgs e) {
            Rename("Renamed");
            /*string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");

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
            toolStripLabel1.Text = DateTime.Now + "修改發行檔名A完成";*/
        }
        private void RenameFilesB_Click(object sender, EventArgs e) {
            Rename("RenamedB");
            /*string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (Directory.Exists("C:\\WarrantDocuments\\RenamedB" + now))
                Directory.Delete("C:\\WarrantDocuments\\RenamedB" + now);
            Directory.CreateDirectory("C:\\WarrantDocuments\\RenamedB" + now);

            for (int i = 0; i < ultraDataSource1.Rows.Count; i++)
                if (Directory.Exists("C:\\WarrantDocuments\\" + ultraDataSource1.Rows[i]["WName"])) {
                    string[] files = Directory.GetFiles("C:\\WarrantDocuments\\" + ultraDataSource1.Rows[i]["WName"]);
                    foreach (string file in files) {
                        if (Path.GetExtension(file).ToLower() != ".xml" && !Path.GetFileName(file).StartsWith("02") && !Path.GetFileName(file).StartsWith("08")
                             && !Path.GetFileName(file).StartsWith("14") && !Path.GetFileName(file).StartsWith("15") && !Path.GetFileName(file).StartsWith("16")
                             && !Path.GetFileName(file).StartsWith("19") && !Path.GetFileName(file).StartsWith("20") && !Path.GetFileName(file).StartsWith("21"))
                            File.Copy(file,
                                 "C:\\WarrantDocuments\\RenamedB" + now + "\\" + ultraDataSource1.Rows[i]["SerialNumber"] + "-" + Path.GetFileName(file), true);
                    }
                }
            toolStripLabel1.Text = DateTime.Now + "修改發行檔名B完成";*/

        }

        private void RenameXML_Click(object sender, EventArgs e) {
            Rename("Xml");
            /*string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");

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
            toolStripLabel1.Text = DateTime.Now + "修改XML檔名完成";*/
        }

        private void Rename(string type) {
            string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (Directory.Exists($"{path}\\{type}{now}"))
                Directory.Delete($"{path}\\{type}{now}", true);
            Directory.CreateDirectory($"{path}\\{type}{now}");
            if (Directory.Exists($"{path}\\{type}OTC{now}"))
                Directory.Delete($"{path}\\{type}OTC{now}", true);
            Directory.CreateDirectory($"{path}\\{type}OTC{now}");

            for (int i = 0; i < ultraDataSource1.Rows.Count; i++)
                if (Directory.Exists($"{path}\\{ultraDataSource1.Rows[i]["WName"]}")) {
                    string[] files = Directory.GetFiles($"{path}\\{ultraDataSource1.Rows[i]["WName"]}");
                    foreach (string file in files) {
                        string toFile = null;
                        string fileName = Path.GetFileName(file);
                        string fileExtension = Path.GetExtension(file).ToLower();
                        switch (type) {
                            case "Xml":
                                if (fileExtension == ".xml") {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\Xml{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}.xml";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\XmlOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}.xml";
                                }
                                break;
                            case "Renamed":
                                if (fileExtension != ".xml") {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\Renamed{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{fileName}";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{fileName}";
                                }
                                break;
                            case "RenamedB":
                                if (fileExtension != ".xml" && !fileName.StartsWith("02") && !fileName.StartsWith("08")
                                && !fileName.StartsWith("14") && !fileName.StartsWith("15") && !fileName.StartsWith("16")
                                && !fileName.StartsWith("19") && !fileName.StartsWith("20") && !fileName.StartsWith("21")) {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedB{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{Path.GetFileName(file)}";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedBOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{Path.GetFileName(file)}";
                                }
                                break;
                        }
                        if (toFile != null)
                            File.Copy(file, toFile, true);
                    }
                }
            if (!Directory.EnumerateFileSystemEntries($"{path}\\{type}{now}").Any())
                Directory.Delete($"{path}\\{type}{now}");
            if (!Directory.EnumerateFileSystemEntries($"{path}\\{type}OTC{now}").Any())
                Directory.Delete($"{path}\\{type}OTC{now}");

            switch (type) {
                case "Xml":
                    toolStripLabel1.Text = DateTime.Now + "修改XML檔名完成";
                    break;
                case "Renamed":
                    toolStripLabel1.Text = DateTime.Now + "修改發行檔名A完成";
                    break;
                case "RenamedB":
                    toolStripLabel1.Text = DateTime.Now + "修改發行檔名B完成";
                    break;
            }
        }

    }
}
