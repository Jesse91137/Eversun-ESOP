using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using System.Net;
using System.Diagnostics;
using System.IO;
using System.Timers;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;
using System.Net.Sockets;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Security.AccessControl;
using System.Security;
using System.DirectoryServices.AccountManagement;
using System.Collections;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace E_SOP
{
    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        //宣告變數
        #region
        public static ListBox MainListBox;
        public Socket newclient;
        public bool Connected;
        public Thread myThread;
        public delegate void MyInvoke(string str);
        FTPClient ftp_Con = new FTPClient();
        string now = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        string com_name = System.Windows.Forms.SystemInformation.ComputerName; //取電腦名稱
        static Mutex m;
        string download_Path = System.Windows.Forms.Application.StartupPath;
        string Fun = "";//下載不改檔名
        string PDF_Name = "";
        string autoOpen = "";//不下載,只作查詢用
        public string sop_path = "", SopName = "", version_old = "", version_new = "", filetempPath = "";
        string repeat, continous = "";        
        DataSet oDS = new DataSet();
        DataSet oDSop = new DataSet();

        DataSet dsSop = new DataSet();
        DataSet dsSopTF = new DataSet();
        public List<string> SOPURL = new List<string>();
        public string ftpuser = "", ftppassword = "", ftpServer = "";
        public string productid, mfc_sop, mfc_sop_desc, testing_sop, testing_sop_desc, packing_sop, packing_sop_desc, assy_sop, assy_sop_desc, spe_packing_sop, spe_packing_sop_desc;
        //------------------------------------------------------------------------------------------
        public string filename = "Setup.ini";
        SetupIniIP ini = new SetupIniIP();

        //WebService
        QueryService.QueryServiceSoapClient queryService = new QueryService.QueryServiceSoapClient("QueryServiceSoap");
        #endregion
        public class Production
        {
            public static string soppath;                     //esop初始目錄
            public static string production1;             //下拉選單1
            public static string production2;             //下拉選單2
            public static string sopfile;                        //量產SOP.試產SOP
            public static string deptItem;                   //製程別            
            public static string station;                       //站別
            public static string twicefail;                     //注意事項(不貳過)
            public static string pack_assy;                     //注意事項(包裝-組裝)
            public static string pack_proc;                     //注意事項(包裝-後製程)
            public static string qa_zone;                     //客戶特殊要求(QA-Zone)
            public static string dispens;                      //SMD點膠圖
            public static string producnotice;           //生產通知單
            public static string btn_producN;           //生產通知單按鈕
            public static string emo;                           //EMO發行文件
            public static string pcb;                               //PCB板料號
            public static string noproducnotice;
            public static ArrayList filepathlist = new ArrayList();            
        }
        public class MyItem
        {
            public string text;
            public string value;

            public MyItem(string text, string value)
            {
                this.text = text;
                this.value = value;
            }
            public override string ToString()
            {
                return text;
            }
        }
        public class SFIS
        {
            public string productid { get; set; }
            public string mfc_sop { get; set; }
            public string mfc_sop_desc { get; set; }
            public string testing_sop { get; set; }
            public string testing_sop_desc { get; set; }

            public string packing_sop { get; set; }
            public string packing_sop_desc { get; set; }
            public string assy_sop { get; set; }
            public string assy_sop_desc { get; set; }
            public string spe_packing_sop { get; set; }
            public string spe_packing_sop_desc { get; set; }
            public string other_sop { get; set; }
            public string other_sop_desc { get; set; }

            public string pe_extnotes { get; set; }
        }
        public RadForm1()
        {
            InitializeComponent();
        }
        private void RadForm1_Load(object sender, EventArgs e)
        {
            if (IsMyMutex("E-SOP"))
            {
                MessageBox.Show("程式正在執行中!!");
                Dispose();//關閉
            }
            version_old = ini.IniReadValue("Version", "version", filename);            
            version_new = selectVerSQL_new("E-SOP");

            int v_old = Convert.ToInt32(version_old.Replace(".", ""));
            int v_new = Convert.ToInt32(version_new.Replace(".", ""));

            lbl_ver.Text = "VER:V" + version_old;
            //判斷自動更新程式是否啟動
            if (v_old != v_new)
            {
                MessageBox.Show("有版本更新VER: V"+ version_new);
                autoupdate();
            }
            else
            {
                string defaultDept = "";
                repeat = ini.IniReadValue("REPEAT_SET", "SET", filename);
                autoOpen = ini.IniReadValue("AUTO_OPEN", "SET", filename);
                PDF_Name = ini.IniReadValue("PDF_NAME", "NAME", filename);
                continous = ini.IniReadValue("CONTINUOUS_SET", "SET", filename);
                Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", filename);
                defaultDept = ini.IniReadValue("Default_Type", "Type_Name", filename);
                ListBox.CheckForIllegalCrossThreadCalls = false;
                string count = ini.IniReadValue("DataBase", "DataBaseCount", filename);
                string sqlpath = @"select Sop_Path from E_SOP_MappingEversun_Table where Process='SopPath'";
                string sqlstr = @"select ID,Process,Groups from E_SOP_Process_Table";
                string sqlemo = @"select Sop_Path from E_SOP_MappingEversun_Table where Process='EMO'";
                DataSet ds = db.reDs(sqlstr);                
                DataSet dspath = db.reDs(sqlpath);
                DataSet dsemo = db.reDs(sqlemo);
                //取得電腦domain名稱,若無加入AD則為NULL
                string domain = (string)Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "Domain", null);
                //domain != null 代表加入AD使用\\miss08\esop公用區\
                Production.soppath = (!string.IsNullOrEmpty(domain)) ? @"\\miss08\esop公用區\" : dspath.Tables[0].Rows[0][0].ToString();
                Production.emo = (!string.IsNullOrEmpty(domain)) ? @"\\miss01\01廠務處管理資料\02-工程部資料\10-MEMO發文\MEMO\" : dsemo.Tables[0].Rows[0][0].ToString();
                rad_mass.Enabled = false;
                rad_trial.Enabled = false;
                foreach (DataRow row in ds.Tables[0].Rows)
                {                    
                    dpd_Process.Items.Add(new MyItem(row["Process"].ToString(), row["ID"].ToString()));                    
                    dpd_Process.SelectedIndex = 0;
                }
                
                //set DefaultDept
                if (!string.IsNullOrEmpty(defaultDept))
                {
                    DefaultDept(defaultDept);
                }
                //
                string tempFilePath = Application.StartupPath + "\\Temp";
                
                if (!Directory.Exists(tempFilePath))
                {
                    //新增資料夾
                    Directory.CreateDirectory(@tempFilePath);
                }
                
                DirectoryInfo faildi = new DirectoryInfo(Application.StartupPath + "\\" + "Temp" + "\\");
                int filecount = faildi.GetFiles().Length;
                if (filecount >= 50)
                {
                    foreach (var fi in faildi.GetFiles())
                    {
                        File.Delete(fi.FullName);
                    }                    
                }
                DeleteOldFiles();
                Btn_WO.Text = "工程編號找SOP";
                Btn_ProduNotice.Text = "生產通知單" + Environment.NewLine + "與 ECN 文件查詢";
            }
        }
        public void DefaultDept(string D)
        {
            switch (D)
            {
                case "QA":
                    dpd_Process.SelectedIndex = 1;
                    break;
                case "SMD":
                    dpd_Process.SelectedIndex = 2;
                    break;
                case "DIP":
                    dpd_Process.SelectedIndex = 3;
                    break;
                case "功能測試":
                    dpd_Process.SelectedIndex = 4;
                    break;
                case "系統組裝":
                    dpd_Process.SelectedIndex = 5;
                    break;
                case "包裝":
                    dpd_Process.SelectedIndex = 6;
                    break;
                case "Coating":
                    dpd_Process.SelectedIndex = 7;
                    break;
                default:
                    break;
            }
        }
        // 開啟檔案並手動更新最後訪問時間
        private void OpenFile(string filePath)
        {
            try
            {
                Process.Start(filePath);

                // 手動更新最後訪問時間
                File.SetLastAccessTime(filePath, DateTime.Now);
            }
            catch (Exception ex)
            {
                MessageBox.Show("無法開啟檔案：" + ex.Message);
            }
        }


        public class SetupIniIP
        { //api ini
            public string path;
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern int GetPrivateProfileString(string section,
            string key, string def, StringBuilder retVal,
            int size, string filePath);
            public void IniWriteValue(string Section, string Key, string Value, string inipath)
            {
                WritePrivateProfileString(Section, Key, Value, Application.StartupPath + "\\" + inipath);
            }
            public string IniReadValue(string Section, string Key, string inipath)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath);
                return temp.ToString();
            }
        }
        private void Txt_SN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                Btn_WO_Click(sender, e);
            }
        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            System_date_ID.Text = "Timer: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        }
        public void PCB_item() 
        {
            DataSet ds = db.reDs("select PCB_item from E_SOP_PCB_Table where Eng_SR='" + txt_EngSr.Text.Trim() + "'");
            if (ds.Tables[0].Rows.Count>0)
            {
                Production.pcb = ds.Tables[0].Rows[0][0].ToString();
            }
            
        }
        /// <summary>
        /// 2023 Q1
        /// 於\\miss01\全廠共用\31-SMD拋料率及不易維修紀錄
        /// 找到不易維修Excel與照片開新視窗列出
        /// </summary>
        public void SMD_Case(string engsr)
        {
            string excelPath = @"\\192.168.4.11\全廠共用\31-SMD拋料率及不易維修紀錄\02-不易維修零件機種\不易維修零件機種.xls";
            string picPath = @"\\miss01\全廠共用\31-SMD拋料率及不易維修紀錄\01-不易維修照片\";
            LoadExcel(excelPath);
            DataTable dt_E = dsData.Tables[0];
            var result = dt_E.AsEnumerable().Where(o => o.Field<string>("機種") == engsr.ToUpper()).CopyToDataTable(); 
            ViewForm viewForm = new ViewForm();
            viewForm.Engsr = engsr.ToUpper().Trim();
            viewForm.Tables = result;
            viewForm.setValue();
            viewForm.ShowDialog();
        }
        DataSet dsData = new DataSet();
        private void LoadExcel(string filename)
        {
            if (filename != "")
            {
                #region office 97-2003
                string excelString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                                     "Data Source=" + filename + ";" +
                                                     "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1\'";
                #endregion

                #region office 2007
                //string excelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename +
                //                 ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
                #endregion

                OleDbConnection cnn = new OleDbConnection(excelString);
                cnn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = cnn;
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                cmd.CommandText = "SELECT * FROM [Sheet1$]";
                adapter.SelectCommand = cmd;
                //DataSet dsData = new DataSet();
                adapter.Fill(dsData);
                cnn = null;
                cmd = null;
                adapter = null;
            }            
        }
        private void Btn_WO_Click(object sender, EventArgs e)
        {
            All_Clear();
            if (dpd_Process.SelectedItem.ToString()== "請選擇...")
            {
                addListItem("未選擇製程!!");
                return;
            }

            if (!txt_EngSr.ReadOnly && !string.IsNullOrEmpty(txt_EngSr.Text))
            {
                bool f = (!rad_trial.Checked && !rad_mass.Checked) ? true : false;
                if (f)
                {
                    listBox2.Items.Add("請選擇試產/量產");
                    return;
                }
            }
            if (txt_EngSr.ReadOnly || (string.IsNullOrEmpty(txt_EngSr.Text) && !string.IsNullOrEmpty(txt_WO.Text)))
            {
                string result=btnSN_Click();

                if (string.IsNullOrEmpty(result))
                {
                    listBox2.Items.Add("找不到SOP文件");
                    return;
                }
                var wipatts = JsonConvert.DeserializeObject<WipAtt>(result);
                //P 試產; M 量產
                rad_trial.Checked = wipatts.wipProcess == "P";
                rad_mass.Checked = wipatts.wipProcess != "P";
                txt_EngSr.Text = wipatts.itemNO.ToUpper().Trim();
            }
            else
            {
                txt_EngSr.Text = txt_EngSr.Text.ToUpper().Trim();
            }

            InputBox inbox = new InputBox();
            txt_WO.Text = txt_WO.Text.ToUpper().Trim();
            
            MyItem myProcess = (MyItem)this.dpd_Process.SelectedItem;
            #region vip
            string sqlvip = @" select * from E_SOP_Eversun_Vip_Customer where eng_sr ='"+ txt_EngSr.Text + "' ";
            DataSet dsvip = db.reDs(sqlvip);
            if (dsvip.Tables[0].Rows.Count>0)
            {
                label1.Text = dsvip.Tables[0].Rows[0][0].ToString();
                label2.Text = dsvip.Tables[0].Rows[0][3].ToString();
            }
            #endregion
            try
            {
                if (!string.IsNullOrEmpty(txt_EngSr.Text.Trim()))
                {
                    string txtWo = txt_EngSr.Text;
                    string FrontPath = txtWo.Substring(0, 2);
                    Judg(txtWo, FrontPath);
                    if (txtWo != null)
                    {
                        try
                        {
                            PCB_item();
                            //WebClient ESOPclient = new WebClient();
                            //生產注意事項
                            if (Production.deptItem=="包裝")
                            {
                                DontTwiceFault(txtWo,FrontPath);
                            }
                            else
                            {
                                //原不貳過程序;PCB路徑相同
                                DirectoryInfo twodi = new DirectoryInfo(Production.twicefail + "\\" + FrontPath);
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                //加入PCB程序一起開                                
                                System.Diagnostics.Process pro2 = new System.Diagnostics.Process();

                                string localpath;
                                if (Directory.Exists(twodi.FullName))
                                {
                                    foreach (var fi in twodi.GetFiles())
                                    {
                                        if (fi.Name.IndexOf(txtWo) >= 0)
                                        {
                                            String fileExtension = fi.Extension.ToLower();
                                            localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + txtWo + fileExtension;

                                            if (File.Exists(localpath))
                                            {
                                                File.Delete(localpath);
                                            }
                                            fi.CopyTo(localpath);
                                            #region //EXCELtoPDF
                                            //try
                                            //{
                                            //    // xlsx 檔案位置
                                            //    string sourcexlsx = localpath;
                                            //    // PDF 儲存位置
                                            //    //取得檔名
                                            //    string targetpdf = Application.StartupPath + "\\" + txtWo + ".pdf";
                                            //    //建立 xlsx application instance
                                            //    Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
                                            //    //開啟 xlsx 檔案
                                            //    var xlsxDocument = appExcel.Workbooks.Open(sourcexlsx);
                                            //    //匯出為 pdf
                                            //    xlsxDocument.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, targetpdf);
                                            //    //關閉 xlsx 檔
                                            //    xlsxDocument.Close();
                                            //    //結束 xlsx
                                            //    appExcel.Quit();

                                            //}
                                            //catch (Exception ee)
                                            //{
                                            //    MessageBox.Show(ee.ToString());
                                            //}

                                            //SOPForm1 Form = new SOPForm1();
                                            ////实例加一
                                            //SOPForm1.m_Static_InstanceCount++;
                                            ////show出来
                                            //SOPForm1.SOPName = Application.StartupPath + "\\" + txtWo + ".pdf";
                                            //Form.Show();
                                            #endregion
                                            
                                            pro.StartInfo.FileName = localpath;
                                            pro.Start();
                                            // 手動更新最後訪問時間
                                            File.SetLastAccessTime(localpath, DateTime.Now);

                                            //pro.WaitForExit(5000); //若excel開啟時再開一次excel會出錯
                                            inbox.MsgModel = "";
                                            inbox.MsgWip_No = txtWo;
                                            inbox.MsgRoute = Production.deptItem;
                                            inbox.ShowDialog(this);

                                        }

                                        //PCB
                                        if (!string.IsNullOrEmpty(Production.pcb) && fi.Name.IndexOf(Production.pcb) >= 0)
                                        {
                                            String fileExtension = fi.Extension.ToLower();
                                            localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + Production.pcb + fileExtension;

                                            if (File.Exists(localpath))
                                            {
                                                File.Delete(localpath);
                                            }
                                            fi.CopyTo(localpath);
                                            
                                            pro2.StartInfo.FileName = localpath;
                                            pro2.Start();
                                            // 手動更新最後訪問時間
                                            File.SetLastAccessTime(localpath, DateTime.Now);

                                            //pro.WaitForExit(5000); //若excel開啟時再開一次excel會出錯
                                            inbox.MsgModel = "";
                                            inbox.MsgWip_No = Production.pcb;
                                            inbox.MsgRoute = Production.deptItem;
                                            inbox.ShowDialog(this);

                                        }
                                    }
                                }
                            }                            
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        
                        //SOP文件
                        DirectoryInfo di = new DirectoryInfo(Production.sopfile + FrontPath + "\\" + txtWo);
                        DataSet data = StationKeyWord(myProcess.value,dpd_station.SelectedItem.ToString());//20210318 StationKeyWord search
                        
                        if (Directory.Exists(di.FullName))
                        {
                            int c = 0;                            
                            foreach (var fi in di.GetFiles())
                            {
                                if (dpd_station.SelectedItem.ToString() == "All") break;
                                for (int i = 0; i < data.Tables[0].Rows.Count; i++)
                                {
                                    Regex regex = new Regex(data.Tables[0].Rows[i][0].ToString());//查詢檔案名稱中有關鍵字的文件
                                    Match m = regex.Match(fi.ToString());
                                    if (m.Success == true)
                                    {
                                        listBox1.Items.Add(fi.Name);
                                        c++;
                                    }
                                }                                
                            }
                            if (c==0)
                            {
                                foreach (var fi in di.GetFiles())
                                {
                                    if (fi.Name.IndexOf("db") <= 0 )
                                    {
                                        listBox1.Items.Add(fi.Name);
                                    }
                                }
                            }
                            else
                            {
                                foreach (var fi in di.GetFiles())
                                {
                                    if (fi.ToString().IndexOf("封面") > 0 || fi.ToString().IndexOf("設定表") > 0)
                                    {
                                        listBox1.Items.Add(fi.Name);
                                    }
                                }
                            }
                        }
                        else
                        {
                            listBox2.Items.Add("沒有SOP文件資料");
                        }
                        //QA_ZOEN
                        try
                        {
                            DirectoryInfo qadi = new DirectoryInfo(Production.qa_zone + "\\" + FrontPath);
                            System.Diagnostics.Process pro = new System.Diagnostics.Process();
                            string localpath;
                            if (Directory.Exists(qadi.FullName))
                            {
                                foreach (var fi in qadi.GetFiles())
                                {
                                    if (fi.Name.IndexOf(txtWo) >= 0)
                                    {
                                        String fileExtension = fi.Extension.ToLower();
                                        localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + txtWo + fileExtension;

                                        if (File.Exists(localpath))
                                        {
                                            File.Delete(localpath);
                                        }
                                        fi.CopyTo(localpath);
                                        
                                        pro.StartInfo.FileName = localpath;
                                        pro.Start();
                                        // 手動更新最後訪問時間
                                        File.SetLastAccessTime(localpath, DateTime.Now);
                                    }
                                }
                            }
                        }
                        catch (Exception ee)
                        {
                            MessageBox.Show(ee.Message);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("請輸入工單編號");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("找不到資料");
            }                 
        }
        public static string StrRight(string param, int length)
        {
            string result = param.Substring(param.Length - length, length);
            return result;
        }
        public string openFile_copyFile(string path)
        {
            FileInfo fi = new FileInfo(path);
            //取得副檔名
            String fileExtension = fi.Extension.ToLower();            
            //組合新路徑copy用
            filetempPath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + fi.Name;
            //刪除檔案後複製
            try
            {
                if (File.Exists(filetempPath))
                {
                    File.SetAttributes(filetempPath, System.IO.FileAttributes.Normal);
                    File.Delete(filetempPath);
                }
                fi.CopyTo(filetempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("請先關閉" + fi.Name + "再開啟");
            }
            
            return filetempPath;
        }
        private void listBox1_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {            

        }
        private DataSet StationKeyWord(string key,string value)
        {
            string str = dpd_Process.SelectedItem.ToString();
            string sqlstr = @"select StationKeyWord from E_SOP_Station_Table where ProcessID='" + key + "' and Station='" + value + "' ";
            
            if (str == "SMD" || str == "DIP")
            {
                sqlstr += @" or Station='總檢圖' ";
            }
            else if (str == "QA")
            {
                sqlstr = @"select StationKeyWord from E_SOP_Station_Table where Station='" + value + "' ";
                if (value=="All")
                {
                    sqlstr = @"select StationKeyWord from E_SOP_Station_Table";
                }                
            }
            DataSet ds= db.reDs(sqlstr);                       
            return ds;
        }
        private void DontTwiceFault(string txtWo,string f_Path)
        {
            InputBox inbox = new InputBox();
            //組裝
            DirectoryInfo assy = new DirectoryInfo(Production.pack_assy + "\\" + f_Path);//組裝
            DirectoryInfo a_proc = new DirectoryInfo(Production.pack_proc + "\\" + f_Path);//後製程
            System.Diagnostics.Process pro = new System.Diagnostics.Process();
            System.Diagnostics.Process pro2= new System.Diagnostics.Process();//PCB
            string localpath;
            if (Directory.Exists(assy.FullName))
            {
                foreach (var fi in assy.GetFiles())
                {
                    if (fi.Name.IndexOf(txtWo) >= 0)
                    {
                        String fileExtension = fi.Extension.ToLower();
                        localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + txtWo + fileExtension;

                        if (File.Exists(localpath))
                        {
                            File.Delete(localpath);
                        }
                        fi.CopyTo(localpath);
                        
                        pro.StartInfo.FileName = localpath;
                        pro.Start();
                        // 手動更新最後訪問時間
                        File.SetLastAccessTime(localpath, DateTime.Now);

                        //pro.WaitForExit(5000); //若excel開啟時再開一次excel會出錯
                        inbox.MsgModel = "";
                        inbox.MsgWip_No = txtWo;
                        inbox.MsgRoute = Production.deptItem;
                        inbox.ShowDialog(this);

                    }

                    //PCB
                    if (!string.IsNullOrEmpty(Production.pcb) && fi.Name.IndexOf(Production.pcb) >= 0)
                    {
                        String fileExtension = fi.Extension.ToLower();
                        localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + Production.pcb + fileExtension;

                        if (File.Exists(localpath))
                        {
                            File.Delete(localpath);
                        }
                        fi.CopyTo(localpath);

                        pro2.StartInfo.FileName = localpath;
                        pro2.Start();
                        // 手動更新最後訪問時間
                        File.SetLastAccessTime(localpath, DateTime.Now);

                        //pro.WaitForExit(5000); //若excel開啟時再開一次excel會出錯
                        inbox.MsgModel = "";
                        inbox.MsgWip_No = Production.pcb;
                        inbox.MsgRoute = Production.deptItem;
                        inbox.ShowDialog(this);
                    }
                }
            }
            //後製程
            if (Directory.Exists(a_proc.FullName))
            {
                foreach (var fi in a_proc.GetFiles())
                {
                    if (fi.Name.IndexOf(txtWo) >= 0)
                    {
                        String fileExtension = fi.Extension.ToLower();
                        localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + txtWo + fileExtension;

                        if (File.Exists(localpath))
                        {
                            File.Delete(localpath);
                        }
                        fi.CopyTo(localpath);
                        
                        pro.StartInfo.FileName = localpath;
                        pro.Start();
                        // 手動更新最後訪問時間
                        File.SetLastAccessTime(localpath, DateTime.Now);

                        //pro.WaitForExit(5000); //若excel開啟時再開一次excel會出錯
                        inbox.MsgModel = "";
                        inbox.MsgWip_No = txtWo;
                        inbox.MsgRoute = Production.deptItem;
                        inbox.ShowDialog(this);

                    }
                    
                    //PCB
                    if (!string.IsNullOrEmpty(Production.pcb) && fi.Name.IndexOf(Production.pcb) >= 0)
                    {
                        String fileExtension = fi.Extension.ToLower();
                        localpath = Application.StartupPath + "\\" + "Temp" + "\\" + Production.deptItem + "-" + Production.pcb + fileExtension;

                        if (File.Exists(localpath))
                        {
                            File.Delete(localpath);
                        }
                        fi.CopyTo(localpath);

                        pro2.StartInfo.FileName = localpath;                        
                        pro2.Start();
                        // 手動更新最後訪問時間
                        File.SetLastAccessTime(localpath, DateTime.Now);

                        //pro.WaitForExit(5000); //若excel開啟時再開一次excel會出錯
                        inbox.MsgModel = "";
                        inbox.MsgWip_No = Production.pcb;
                        inbox.MsgRoute = Production.deptItem;
                        inbox.ShowDialog(this);

                    }
                }
            }
        }
        public static bool IsFileInUse(string fileName)
        {
            bool inUse = true;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None);
                inUse = false;
            }
            catch
            {
                // 檔案正在使用中，捕捉異常但不做處理
            }
            finally
            {
                fs?.Close();  // 使用 ?. 確保不為 null 時才調用 Close
            }

            // true表示正在使用, false表示沒有使用
            return inUse;
        }

        public static void WipeFile(string fileName, int timesToWrite)
        {
            try
            {
                if (System.IO.File.Exists(fileName))
                {
                    // 設定檔案的屬性為一般，防止檔案是唯讀的
                    System.IO.File.SetAttributes(fileName, System.IO.FileAttributes.Normal);

                    // 計算檔案儲存的 sectors 數目
                    double sectors = Math.Ceiling(new System.IO.FileInfo(fileName).Length / 512.0);

                    // 建立虛擬暫存
                    byte[] dummyBuffer = new byte[512];

                    // 建立隨機數生成器
                    System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();

                    // 打開檔案流進行覆蓋寫入
                    using (FileStream inputStream = new FileStream(fileName, FileMode.Open, FileAccess.Write, FileShare.None))
                    {
                        for (int currentPass = 0; currentPass < timesToWrite; currentPass++)
                        {
                            inputStream.Position = 0;
                            for (int sectorsWritten = 0; sectorsWritten < sectors; sectorsWritten++)
                            {
                                rng.GetBytes(dummyBuffer); // 隨機生成數據
                                inputStream.Write(dummyBuffer, 0, dummyBuffer.Length); // 寫入數據到檔案流
                            }
                        }
                        inputStream.SetLength(0); // 清空檔案內容
                    }

                    // 重置檔案的日期資訊
                    DateTime dt = new DateTime(2037, 1, 1, 0, 0, 0);
                    System.IO.File.SetCreationTime(fileName, dt);
                    System.IO.File.SetLastAccessTime(fileName, dt);
                    System.IO.File.SetLastWriteTime(fileName, dt);

                    // 刪除檔案
                    System.IO.File.Delete(fileName);
                }
            }
            catch (Exception ex)
            {
                // 可選的錯誤處理，例如記錄錯誤資訊
                Console.WriteLine($"Error wiping file {fileName}: {ex.Message}");
            }
        }

        public static void DeleteFilesInFolder(string folderPath, TimeSpan fileExpirationDuration)
        {
            // 只刪除資料夾中的檔案，不遞迴進入子資料夾
            foreach (string file in Directory.GetFiles(folderPath)) // 只針對檔案
            {
                try
                {
                    if (System.IO.File.Exists(file))
                    {
                        // 獲取檔案的最後訪問時間
                        DateTime lastAccessTime = File.GetLastAccessTime(file);

                        // 如果檔案超過設定的時間未被訪問，則刪除
                        if (DateTime.Now - lastAccessTime > fileExpirationDuration)
                        {
                            FileInfo fi = new FileInfo(file);
                            if ((fi.Attributes & FileAttributes.ReadOnly) != 0) // 檢查是否唯讀
                            {
                                fi.Attributes = FileAttributes.Normal; // 如果是，則設為一般屬性
                            }

                            // 檢查檔案是否正在使用
                            if (IsFileInUse(file))
                            {
                                WipeFile(file, 20);  // 如果正在使用，則進行覆蓋寫入
                            }
                            else
                            {
                                // 檔案未使用，直接刪除
                                System.IO.File.Delete(file);
                            }
                            Console.WriteLine($"檔案已刪除: {file}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"無法刪除檔案: {file}, 錯誤: {ex.Message}");
                }
            }
        }

        private void DeleteOldFiles()
        {
            // 設定過期的時間（例如：檔案超過一天未使用則刪除）
            TimeSpan fileExpirationDuration = TimeSpan.FromDays(1);

            string folderPath = Path.Combine(Application.StartupPath, "Temp");

            // 刪除資料夾中的過期檔案
            DeleteFilesInFolder(folderPath, fileExpirationDuration);
        }



        private string QueryDataSet(string txtWO)
        {
            string txt_SR = string.Empty;
            string sql = "select WO_NO,ENG_SR from SAP_DATA_TO_ESOP where WO_NO='"+ txtWO + "' ";//SAP
            DataSet ds = db.reDs(sql);
            
            if (ds.Tables[0].Rows.Count > 0)//SAP
            {
                string []temparry= ds.Tables[0].Rows[0]["ENG_SR"].ToString().Trim().Split('-');
                txt_SR = temparry[0].Trim();
            }//20211125改SAP優先,若無在找SFIS
            else
            {
                DataSet ds2 = queryService.QueryWO_Info(txtWO);//SFIS
                if (ds2.Tables[0].Rows.Count > 0)//SFIS
                {
                    string[] temparry = ds2.Tables[0].Rows[0]["ENG_SR"].ToString().Trim().Split('-');
                    txt_SR = temparry[0].Trim();
                }
            }
            
            if (txt_SR.IndexOf("SERVICE")>=0) 
            {
                txt_SR = string.Empty;
            }
            return txt_SR;
        }
        private void dpd_Process_SelectedIndexChanged(object sender, EventArgs e)
        {
            dpd_station.Items.Clear();
            string str = dpd_Process.SelectedItem.ToString();
            switch (str)
            {
                case "QA":
                    string sqlStation = @"select distinct station from E_SOP_Station_Table ";
                    DataSet ds = db.reDs(sqlStation);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        dpd_station.Items.Add("All");
                        foreach (DataRow row in ds.Tables[0].Rows)
                        {
                            dpd_station.Items.Add(row["Station"].ToString());
                            dpd_station.SelectedIndex = 0;
                        }
                    }
                    rad_mass.Enabled = true;
                    rad_trial.Enabled = true;
                    txt_EngSr.ReadOnly = false;
                    txt_EngSr.BackColor = Color.White;
                    break;
                default:
                    rad_mass.Enabled = false;
                    rad_trial.Enabled = false;
                    rad_mass.Checked = false;
                    rad_trial.Checked = false;
                    txt_EngSr.ReadOnly = true;
                    txt_EngSr.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
                    break;
            }
            if (dpd_Process.SelectedItem.ToString()!="請選擇...")
            {
                MyItem myProcess = (MyItem)this.dpd_Process.SelectedItem;
                string sqlStation = @"select distinct station from E_SOP_Station_Table where ProcessID= '" + myProcess.value + "'";
                DataSet ds = db.reDs(sqlStation);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        dpd_station.Items.Add(row["Station"].ToString());
                        dpd_station.SelectedIndex = 0;
                    }
                }
            }
            
            Production.deptItem = dpd_Process.SelectedItem.ToString();
        }

        public void Judg(string wono,string frontpath)
        {
            //Production.sopfile = (dssop.Tables[0].Rows.Count > 0) ? Production.soppath + dssop.Tables[0].Rows[0][0].ToString() : "";//(條件式) ?成立 :不成立  
            //同時確認量產/試產工程編號是否存在檔案;同時存在以量產路徑為主
            string sqlstr = @"select sop_path from E_SOP_MappingEversun_Table ";
            //SOP FilePath
            string sqlsop = sqlstr+@" where process= 'SOP'";
            DataSet dssop = db.reDs(sqlsop);
            //Trial=試產 ; Mass=量產            
            string Mass = Production.soppath + dssop.Tables[0].Rows[0][0].ToString();
            string Trial = Production.soppath + dssop.Tables[0].Rows[1][0].ToString();
            string condition = string.Empty;
            
            Production.sopfile = (rad_mass.Checked) ? Mass : Trial;
            //DirectoryInfo di = new DirectoryInfo(Mass + frontpath + "\\" + wono);
            //Production.sopfile = (Directory.Exists(di.FullName)) ? Mass : Trial;
            
            condition = (Production.sopfile.IndexOf("量產") > 0) ? "量產" : "試產";
            //QA-Zone
            //string sqlqa = sqlstr + @" where process= 'QA-Zone' ";// and process='"+ condition + "'";
            //DataSet dsqa = db.reDs(sqlqa);
            //Production.qa_zone = (dsqa.Tables[0].Rows.Count > 0) ? Production.soppath + dsqa.Tables[0].Rows[0][0].ToString() : "";//(條件式) ?成立 :不成立  

            #region //注意事項(不貳過)
            //注意事項; Precautions = "生產注意事項";            
            string sqltwice = sqlstr + @" where process= 'TwiceFail' ";// and process='" + condition + "'";
            DataSet dstwice = db.reDs(sqltwice);
            MyItem myProcess = (MyItem)this.dpd_Process.SelectedItem;
            
            string twiceGroup = @"select Groups from E_SOP_Process_Table where Process = '" + myProcess.text + "'";
            DataSet dsGroup= db.reDs(twiceGroup);
            Production.twicefail = (dstwice.Tables[0].Rows.Count > 0) ?
                Production.soppath + dstwice.Tables[0].Rows[0][0].ToString() + dsGroup.Tables[0].Rows[0][0].ToString() + @"\" : "";//(條件式) ?成立 :不成立 

            //↓↓process→包裝 例外判斷
            if (myProcess.text=="包裝")
            {
                //SOP目錄+不貳過目錄+(包裝_組裝)
                Production.pack_proc = Production.soppath + dstwice.Tables[0].Rows[0][0].ToString() + @"組裝\";
                //SOP目錄+不貳過目錄+(包裝_後製程)
                Production.pack_assy = Production.soppath + dstwice.Tables[0].Rows[0][0].ToString() + @"後製程\";
            }            
            #endregion
        }

        private void Btn_ProduNotice_Click(object sender, EventArgs e)
        {
            All_Clear();
            Production.filepathlist.Clear(); 
            try
            {
                if (txt_WO.Text.ToUpper().Trim().Length != 12)
                {
                    listBox2.Items.Add("工單號碼不對!!");
                    return;
                }
                //SOP 生產通知單
                string sqlstr = @"select sop_path from E_SOP_MappingEversun_Table where process= 'ProduNotice'";
                DataSet dssop = db.reDs(sqlstr);
                Production.producnotice = dssop.Tables[0].Rows[0][0].ToString();
                txt_WO.Text = txt_WO.Text.ToUpper().Trim();
                txt_EngSr.Text = txt_EngSr.Text.ToUpper().Trim();
                btnSN_Click();
                if (!string.IsNullOrEmpty(txt_EngSr.Text.Trim()))
                {
                    string txtWo = txt_EngSr.Text;
                    string FrontPath = txtWo.Substring(0, 2);

                    //txt_WO=工程編號,既然有工程編號就可以篩選出前兩碼客戶別,因此利用txtWo加速查詢
                    DirectoryInfo di = new DirectoryInfo(Production.producnotice + FrontPath + "\\" + txtWo);
                    //DirectoryInfo di = new DirectoryInfo(Production.producnotice + FrontPath + "\\");
                    if (Directory.Exists(di.FullName))
                    {
                        Production.btn_producN = "btn_pro_sender";
                        #region 完整路徑搜尋
                        //foreach (var fi in di.GetFiles())
                        //{
                        //    //if ((fi.Name.Trim()).Contains(txt_SN.Text) && fi.Name.IndexOf("db") <= 0)  //字串包含""空字串,contains會誤判
                        //    //if ((txt_SN.Text).Substring(0, 7) == fi.Name.Substring(0, 7) && fi.Name.IndexOf("db") <= 0) //多單號合併檔會找不到
                        //    if (fi.Name.IndexOf(txt_SN.Text)>=0 && fi.Name.IndexOf("db") <= 0)
                        //    {
                        //        listBox1.Items.Add(fi.Name);
                        //    }
                        //}
                        #endregion
                        //FindFile(di.FullName);//工程編號前2碼作為路徑依據
                        foreach (var fi in di.GetFiles())
                        {
                            if (fi.Name.IndexOf(txt_WO.Text) >= 0 && fi.Name.IndexOf("db") <= 0) 
                            { 
                                listBox1.Items.Add(fi.Name);
                            }
                        }
                        if (listBox1.Items.Count==0)
                        {
                            listBox1.Items.Add("沒有生產通知單!!");
                            Production.noproducnotice = "NoReport";
                        }
                    }
                    else
                    {
                        listBox1.Items.Add("沒有生產通知單!!");
                        Production.noproducnotice = "NoReport";
                    }
                }
                else//如果生產通知單工單號碼找不到工程編號就
                {
                    Production.btn_producN = "noEngSearch";
                    FindFile(Production.producnotice);
                }                
                FindFile(Production.emo);
            }
            catch (Exception ee)
            {
                MessageBox.Show("請輸入工單號碼!!");
            }
        }
        #region 列出指定目錄下的所有目錄並搜尋特定檔案        
        public void FindFile(string dirPath)//引數dirPath為指定的目錄
        {
            //在指定目錄及子目錄下查詢檔案,在listBox1中列出子目錄及檔案
            DirectoryInfo Dir = new DirectoryInfo(dirPath);
            try
            {
                foreach (DirectoryInfo d in Dir.GetDirectories())//查詢子目錄
                {
                    FindFile(Dir + d.ToString() + "\\");
                    //listBox1.Items.Add(Dir + d.ToString() + "\\");//listBox1中填加目錄名
                }
                foreach (FileInfo f in Dir.GetFiles())//查詢檔案
                {
                    if (f.Name.IndexOf(txt_WO.Text) >= 0)
                    {
                        listBox1.Items.Add(f.ToString());//listBox1中填加檔名
                       Production.filepathlist.Add(Dir.FullName + f.ToString());
                    }
                }
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
        }
        #endregion
        public string btnSN_Click()
        {
            #region MyRegion
            /*
             try
            {
                if (!string.IsNullOrEmpty(txt_WO.Text.Trim()))
                {
                    txt_SR = QueryDataSet(txt_WO.Text);
                    string FrontPath = txt_SR.Substring(0, 2);
                    if (txt_SR != null)
                    {
                        txt_EngSr.Text = txt_SR;
                        txt_EngSr.Focus();                        
                    }
                    else
                    {
                        addListItem("找不到工單編號資料");
                    }
                }
            }
            catch
            {
                addListItem("找不到工單編號資料");
            }
             */
            #endregion

            var result = ApiRoute.GetMethod($"api/WipAtts/{txt_WO.Text}");
            if (result == "error" || result == "無法連線WebAPI")
            {
                MessageBox.Show(result);
                
            }

            return result;
        }
        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            string special= listBox1.SelectedItem.ToString();

            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0) //排除memo.service file
            {
                filename = listBox1.SelectedItem.ToString();
                path = Production.producnotice;
                //Production.producnotice;
                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                folder_F = strrtn.Substring(0, 2);
            }
            else if (Production.btn_producN == "noEngSearch")
            {
                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
            }
            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
            {
                //if (Production.noproducnotice == "NoReport")
                //{
                //    path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                //}
                //else
                //{
                //    path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                //}
                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
            }
            else
            {
                filename = listBox1.SelectedItem.ToString();
                path = Production.sopfile;
                folder_F = filename.Substring(0, 2);
                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
            }
            if (listBox1.SelectedItem.ToString()=="沒有生產通知單!!")
            {
                return;
            }
            string SOPpath = (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) ? path : path + folder_F + @"\" + strrtn + @"\" + filename;            
            SOPpath = openFile_copyFile(SOPpath.TrimEnd('\\'));
            try
            {
                // 手動更新最後訪問時間
                File.SetLastAccessTime(SOPpath, DateTime.Now);
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"存取被拒：無法設定最後訪問時間。\n錯誤詳情: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"發生錯誤：{ex.Message}");
            }
            using (System.Diagnostics.Process prc = new System.Diagnostics.Process())
            {
                prc.StartInfo.FileName = SOPpath;
                prc.Start();
            }
        }

        private void txt_SN_TextChanged(object sender, EventArgs e)
        {
            txt_EngSr.Text = "";
        }

        private void listBox3_DrawItem(object sender, DrawItemEventArgs e)
        {
            ListBox list = (ListBox)sender;

            // Draw the background of the ListBox control for each item.
            e.DrawBackground();
            // Define the default color of the brush as black.
            Brush myBrush = Brushes.Black;

            // Determine the color of the brush to draw each item based 
            // on the index of the item to draw.
            
            switch (e.Index)
            {
                case 0:
                    myBrush = Brushes.Red;
                    break;
                case 1:
                    myBrush = Brushes.Orange;
                    break;
                case 2:
                    myBrush = Brushes.Purple;
                    break;
            }

            // Draw the current item text based on the current Font 
            // and the custom brush settings.
            e.Graphics.DrawString(list.Items[e.Index].ToString(),
                e.Font, myBrush, e.Bounds, StringFormat.GenericDefault);
            // If the ListBox has focus, draw a focus rectangle around the selected item.
            e.DrawFocusRectangle();
        }


        //private void Btn_Iso_Click(object sender, EventArgs e)
        //{
        //    All_Clear();
        //    if (txt_Iso.Text.IndexOf("QO3") == 0)
        //    {
        //        ISOZONE_ISOSERCH(txt_Iso.Text.Trim());
        //    }
        //    else
        //    {
        //        SearchTempSOP(txt_Iso.Text.Trim());
        //    }
        //    if (autoOpen == "OFF")
        //    {

        //        //OPENPDF
        //        try
        //        {

        //            for (int i = 0; i < listBox1.Items.Count; i++)
        //            {
        //                if (SOPURL[i].ToString().IndexOf("documents_maintain_files") >= 0)
        //                {
        //                    doPdf(i + 1, SOPURL[i].ToString());
        //                }
        //                else
        //                {
        //                    doPdf_TF(i + 1, SOPURL[i].ToString());
        //                }
        //            }

        //        }
        //        catch (Exception ex)
        //        {
        //            addListItem(ex.Message);
        //            MessageBox.Show("找不到檔案!!");
        //        }

        //    }

        //}
        #region SOP Open Button
        private void Btn_Sop1_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            //test
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            //Getftp("E-SOP");
            if (Fun == "ON")
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        if (SOPForm1.m_Static_InstanceCount < 1)
                        {
                            string special = listBox1.SelectedItem.ToString();
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);                                
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }


                            //string SOPpath = ISOZONE_SOPNAMESERCH(listBox1.SelectedItem.ToString());
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            //string SOPpath = path + folder_F + @"\" + strrtn + @"\" + filename;
                            String fileExtension = Path.GetExtension(SOPpath);

                            //if (!string.IsNullOrEmpty(SOPpath))
                            if(fileExtension == ".pdf")
                            {
                                //Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SOPpath);
                                Uri uri = new Uri(SOPpath);

                                //string Filename = SOPpath.Replace("documents_maintain_files/", "");
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + Production.deptItem + "-" + filename);
                                
                                SOPForm1 Form = new SOPForm1();
                                //实例加一
                                SOPForm1.m_Static_InstanceCount++;
                                SOPForm1.SOPName = Production.deptItem + "-" + filename;
                                Form.Text = filename;
                                //show出来
                                Form.Show();

                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();                                
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                            //else
                            //{
                            //    string Filename = listBox1.SelectedItem.ToString();                                                                
                            //    WebFTP.Downloadfile(Filename + ".pdf", download_Path + "\\" + Filename, "E-SOP", dpd_Process.SelectedItem.ToString());                                
                            //    SOPForm1 Form = new SOPForm1();
                            //    //实例加一
                            //    SOPForm1.m_Static_InstanceCount++;
                            //    SOPForm1.SOPName = Filename;
                            //    //show出来
                            //    Form.Show();
                            //}
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP1");
                        }

                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }


                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            else
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm1.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            //string SOPpath = path + folder_F + @"\" + strrtn + @"\" + filename;
                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + "SOP1.Pdf");

                                SOPForm1 Form = new SOPForm1();
                                //实例加一
                                SOPForm1.m_Static_InstanceCount++;
                                SOPForm1.SOPName = "SOP1.Pdf";
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();                                
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP1");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }

                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        
        private void Btn_Sop2_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            //Getftp("E-SOP");
            if (Fun == "ON")
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    //PDF_Name = ini.IniReadValue("PDF_NAME", "NAME", filename);

                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm2.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            //string SOPpath = path + folder_F + @"\" + strrtn + @"\" + filename;
                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                //Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SOPpath);
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + Production.deptItem + "-" + filename);

                                SOPForm2 Form = new SOPForm2();
                                //实例加一
                                SOPForm2.m_Static_InstanceCount++;
                                SOPForm2.SOPName = Production.deptItem + "-" + filename;
                                Form.Text = filename;
                                //show出来
                                Form.Show();

                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();                                
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }

                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP2");
                        }

                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }


                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            else
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {

                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm2.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            //string SOPpath = path + folder_F + @"\" + strrtn + @"\" + filename;
                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + "SOP2.Pdf");

                                SOPForm2 Form = new SOPForm2();
                                //实例加一
                                SOPForm2.m_Static_InstanceCount++;
                                SOPForm2.SOPName = "SOP2.Pdf";
                                //show出来
                                Form.Show();

                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();                                
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP2");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }

                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }

        }
        
        private void Btn_Sop3_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            //Getftp("E-SOP");
            if (Fun == "ON")
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm3.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + Production.deptItem + "-" + filename);

                                SOPForm3 Form = new SOPForm3();
                                //实例加一
                                SOPForm3.m_Static_InstanceCount++;
                                SOPForm3.SOPName = Production.deptItem + "-" + filename;
                                Form.Text = filename;
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP3");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            else
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {

                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm3.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + "SOP3.Pdf");

                                SOPForm3 Form = new SOPForm3();
                                //实例加一
                                SOPForm3.m_Static_InstanceCount++;
                                SOPForm3.SOPName = "SOP3.Pdf";
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP3");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }

                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        private void Btn_Sop4_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            //Getftp("E-SOP");
            if (Fun == "ON")
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm4.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + Production.deptItem + "-" + filename);

                                SOPForm4 Form = new SOPForm4();
                                //实例加一
                                SOPForm4.m_Static_InstanceCount++;
                                SOPForm4.SOPName = Production.deptItem + "-" + filename;
                                Form.Text = filename;
                                //show出来
                                Form.Show();

                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP4");
                        }

                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            else
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm4.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + "SOP4.Pdf");

                                SOPForm4 Form = new SOPForm4();
                                //实例加一
                                SOPForm4.m_Static_InstanceCount++;
                                SOPForm4.SOPName = "SOP4.Pdf";
                                //show出来
                                Form.Show();

                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP4");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }

                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        private void Btn_Sop5_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            //Getftp("E-SOP");
            if (Fun == "ON")
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm5.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);
                                
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + Production.deptItem + "-" + filename);

                                SOPForm5 Form = new SOPForm5();
                                //实例加一
                                SOPForm5.m_Static_InstanceCount++;
                                SOPForm5.SOPName = Production.deptItem + "-" + filename;
                                Form.Text = filename;
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP5");
                        }

                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            else
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm5.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + "SOP5.Pdf");

                                SOPForm5 Form = new SOPForm5();
                                //实例加一
                                SOPForm5.m_Static_InstanceCount++;
                                SOPForm5.SOPName = "SOP5.Pdf";
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP5");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }

                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        private void Btn_Sop6_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string folder_F = string.Empty;
            string strrtn = string.Empty;
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            //Getftp("E-SOP");

            if (Fun == "ON")
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm6.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + Production.deptItem + "-" + filename);

                                SOPForm6 Form = new SOPForm6();
                                //实例加一
                                SOPForm6.m_Static_InstanceCount++;
                                SOPForm6.SOPName = Production.deptItem + "-" + filename;
                                Form.Text = filename;
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;                                
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP6");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            else
            {
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        WebClient ESOPclient = new WebClient();
                        string special = listBox1.SelectedItem.ToString();
                        if (SOPForm6.m_Static_InstanceCount < 1)
                        {
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.producnotice;
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                folder_F = strrtn.Substring(0, 2);
                            }
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true ) //包含memo.service file
                            {
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            else
                            {
                                filename = listBox1.SelectedItem.ToString();
                                path = Production.sopfile;
                                folder_F = filename.Substring(0, 2);
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

                            String fileExtension = Path.GetExtension(SOPpath);

                            if (fileExtension == ".pdf")
                            {
                                Uri uri = new Uri(SOPpath);

                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/" + "Temp" + "/" + "SOP6.Pdf");

                                SOPForm6 Form = new SOPForm6();
                                //实例加一
                                SOPForm6.m_Static_InstanceCount++;
                                SOPForm6.SOPName = "SOP6.Pdf";
                                //show出来
                                Form.Show();
                            }
                            else
                            {
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        else
                        {
                            MessageBox.Show("請先關閉SOP6");
                        }
                    }
                    catch (Exception ex)
                    {
                        //showinfo.Clear();
                        //showinfo.SelectionColor = Color.Red;
                        //showinfo.AppendText(ex.Message);
                        listBox2.Items.Clear();
                        addListItem(ex.Message);
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }

                }
                else
                {
                    //showinfo.Clear();
                    //showinfo.SelectionColor = Color.Red;
                    //showinfo.AppendText("尚未查詢,無法下載!!");
                    listBox2.Items.Clear();
                    addListItem("尚未查詢,無法下載!!");
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        #endregion
        private void Btn_Clear_Click(object sender, EventArgs e)
        {
            All_Clear();
            txt_EngSr.Text = "";
            txt_WO.Text = "";
            //txt_Iso.Text = "";
        }
        private void Btn_Exit_Click(object sender, EventArgs e)
        {
            Close();
        }
        #region //doPdf_TF
        //public void doPdf_TF(int SopQty, string Sopname)
        //{//下載


        //    //Getftp("E-SOP");
        //    ftp_Con.RemoteHost = ftpServer;
        //    ftp_Con.RemoteUser = ftpuser;
        //    ftp_Con.RemotePass = ftppassword;
        //    string filepath = "";
        //    if(dpd_Process.SelectedItem.ToString()=="T1" || dpd_Process.SelectedItem.ToString() == "T2")
        //    {
        //        filepath = "T2";
        //    }
        //    else if (dpd_Process.SelectedItem.ToString() == "Packing")
        //    {
        //        filepath = "Packing";
        //    }
        //    else if (dpd_Process.SelectedItem.ToString() == "ASSY1" || dpd_Process.SelectedItem.ToString() == "ASSY2" ||
        //                     dpd_Process.SelectedItem.ToString() == "ASSY3" || dpd_Process.SelectedItem.ToString() == "ASSY4" ||
        //                     dpd_Process.SelectedItem.ToString() == "ASSY5" || dpd_Process.SelectedItem.ToString() == "ASSY6" ||
        //                     dpd_Process.SelectedItem.ToString() == "ASSY")
        //    {
        //        filepath = "ASSY";
        //    }
        //    else if (dpd_Process.SelectedItem.ToString() == "ASSY-T1" || dpd_Process.SelectedItem.ToString() == "ASSY-T2")
        //    {
        //        filepath = "ASSY-T2";
        //    }
        //    else if (dpd_Process.SelectedItem.ToString() == "Packing2" || dpd_Process.SelectedItem.ToString() == "Packing2-1" ||
        //                    dpd_Process.SelectedItem.ToString() == "Packing2-2")
        //    {
        //        filepath = "Packing2";
        //    }
        //    else if (dpd_Process.SelectedItem.ToString() == "Rework" )
        //    {
        //        filepath = "Rework";
        //    }

        //    if (SopQty == 1)
        //    {
        //        if (SOPForm1.m_Static_InstanceCount < 1)
        //        {
        //            if (WebFTP.Downloadfile(Sopname + ".pdf", download_Path + "\\" + "SOP" + SopQty.ToString() + ".pdf", "E-SOP", filepath) == true)
        //            {
        //                SOPForm1 Form = new SOPForm1();
        //                //实例加一
        //                SOPForm1.SOPName = "SOP1.pdf";
        //                SOPForm1.m_Static_InstanceCount++;
        //                //show出来
        //                Form.Show();
        //            }
        //            else
        //            {
        //                addListItem("無相關SOP資料!!");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("請先關閉SOP1");
        //        }
        //    }
        //    else if (SopQty == 2)
        //    {
        //        if (SOPForm2.m_Static_InstanceCount < 1)
        //        {
        //            if (WebFTP.Downloadfile(Sopname, download_Path + "\\" + "SOP" + SopQty.ToString() + ".pdf", "E-SOP", filepath) == true)
        //            {
        //                SOPForm2 Form = new SOPForm2();
        //                //实例加一
        //                SOPForm2.SOPName = "SOP2.pdf";
        //                SOPForm2.m_Static_InstanceCount++;
        //                //show出来
        //                Form.Show();
        //            }
        //            else
        //            {
        //                addListItem("無相關SOP資料!!");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("請先關閉SOP2");
        //        }
        //    }
        //    else if (SopQty == 3)
        //    {
        //        if (SOPForm3.m_Static_InstanceCount < 1)
        //        {
        //            if (WebFTP.Downloadfile(Sopname, download_Path + "\\" + "SOP" + SopQty.ToString() + ".pdf", "E-SOP", filepath) == true)
        //            {
        //                SOPForm3 Form = new SOPForm3();
        //                //实例加一
        //                SOPForm3.SOPName = "SOP3.pdf";
        //                SOPForm3.m_Static_InstanceCount++;
        //                //show出来
        //                Form.Show();
        //            }
        //            else
        //            {
        //                addListItem("無相關SOP資料!!");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("請先關閉SOP3");
        //        }
        //    }
        //    else if (SopQty == 4)
        //    {
        //        if (SOPForm4.m_Static_InstanceCount < 1)
        //        {
        //            if (WebFTP.Downloadfile(Sopname, download_Path + "\\" + "SOP" + SopQty.ToString() + ".pdf", "E-SOP", filepath) == true)
        //            {
        //                SOPForm4 Form = new SOPForm4();
        //                //实例加一
        //                SOPForm4.SOPName = "SOP4.pdf";
        //                SOPForm4.m_Static_InstanceCount++;
        //                //show出来
        //                Form.Show();
        //            }
        //            else
        //            {
        //                addListItem("無相關SOP資料!!");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("請先關閉SOP4");
        //        }
        //    }
        //    else if (SopQty == 5)
        //    {
        //        if (SOPForm5.m_Static_InstanceCount < 1)
        //        {
        //            if (WebFTP.Downloadfile(Sopname, download_Path + "\\" + "SOP" + SopQty.ToString() + ".pdf", "E-SOP", filepath) == true)
        //            {
        //                SOPForm5 Form = new SOPForm5();
        //                //实例加一
        //                SOPForm5.SOPName = "SOP5.pdf";
        //                SOPForm5.m_Static_InstanceCount++;
        //                //show出来
        //                Form.Show();
        //            }
        //            else
        //            {
        //                addListItem("無相關SOP資料!!");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("請先關閉SOP5");
        //        }
        //    }
        //    else if (SopQty == 6)
        //    {
        //        if (SOPForm6.m_Static_InstanceCount < 1)
        //        {
        //            if (WebFTP.Downloadfile(Sopname, download_Path + "\\" + "SOP" + SopQty.ToString() + ".pdf", "E-SOP", filepath) == true)
        //            {
        //                SOPForm6 Form = new SOPForm6();
        //                //实例加一
        //                SOPForm6.SOPName = "SOP6.pdf";
        //                SOPForm6.m_Static_InstanceCount++;
        //                //show出来
        //                Form.Show();
        //            }
        //            else
        //            {
        //                addListItem("無相關SOP資料!!");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("請先關閉SOP6");
        //        }
        //    }



        //    //addListItem("開啟暫行SOP...");


        //    ftp_Con.DisConnect();

        //}
        #endregion

        public void doPdf(int SopQty,string SopWeb)
        {
            WebClient ESOPclient = new WebClient();
            

            if (SopQty == 1)
            {
                if (SOPForm1.m_Static_InstanceCount < 1)
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb);
                    
                    string Filename = SopWeb.Replace("documents_maintain_files/", "");
                    ESOPclient.DownloadFile(uri,Application.StartupPath +"/SOP1.pdf");
                    
                    SOPForm1 Form = new SOPForm1();
                    //实例加一
                    SOPForm1.SOPName = "SOP1.pdf";
                    SOPForm1.m_Static_InstanceCount++;
                    //show出来
                    Form.Show();
                }
                else
                {
                    MessageBox.Show("請先關閉SOP1");
                }
            }
            else if(SopQty == 2)
            {
                if (SOPForm2.m_Static_InstanceCount < 1)
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb);

                    string Filename = SopWeb.Replace("documents_maintain_files/", "");
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP2.pdf");

                    SOPForm2 Form = new SOPForm2();
                    //实例加一
                    SOPForm2.SOPName = "SOP2.pdf";
                    SOPForm2.m_Static_InstanceCount++;
                    //show出来
                    Form.Show();
                }
                else
                {
                    MessageBox.Show("請先關閉SOP2");
                }
            }
            else if (SopQty == 3)
            {
                if (SOPForm3.m_Static_InstanceCount < 1)
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb);

                    string Filename = SopWeb.Replace("documents_maintain_files/", "");
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP3.pdf");

                    SOPForm3 Form = new SOPForm3();
                    //实例加一
                    SOPForm3.SOPName = "SOP3.pdf";
                    SOPForm3.m_Static_InstanceCount++;
                    //show出来
                    Form.Show();
                }
                else
                {
                    MessageBox.Show("請先關閉SOP3");
                }
            }
            else if (SopQty == 4)
            {
                if (SOPForm4.m_Static_InstanceCount < 1)
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb);

                    string Filename = SopWeb.Replace("documents_maintain_files/", "");
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP4.pdf");

                    SOPForm4 Form = new SOPForm4();
                    //实例加一
                    SOPForm4.SOPName = "SOP4.pdf";
                    SOPForm4.m_Static_InstanceCount++;
                    //show出来
                    Form.Show();
                }
                else
                {
                    MessageBox.Show("請先關閉SOP4");
                }
            }
            else if (SopQty == 5)
            {
                if (SOPForm5.m_Static_InstanceCount < 1)
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb);

                    string Filename = SopWeb.Replace("documents_maintain_files/", "");
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP5.pdf");

                    SOPForm5 Form = new SOPForm5();
                    //实例加一
                    SOPForm5.SOPName = "SOP5.pdf";
                    SOPForm5.m_Static_InstanceCount++;
                    //show出来
                    Form.Show();
                }
                else
                {
                    MessageBox.Show("請先關閉SOP5");
                }
            }
            else if (SopQty == 6)
            {
                if (SOPForm6.m_Static_InstanceCount < 1)
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb);

                    string Filename = SopWeb.Replace("documents_maintain_files/", "");
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP6.pdf");

                    SOPForm6 Form = new SOPForm6();
                    //实例加一
                    SOPForm6.SOPName = "SOP6.pdf";
                    SOPForm6.m_Static_InstanceCount++;
                    //show出来
                    Form.Show();
                }
                else
                {
                    MessageBox.Show("請先關閉SOP6");
                }
            }
        }
        private bool IsMyMutex(string prgname)
        {
            bool IsExist;
            m = new Mutex(true, prgname, out IsExist);
            GC.Collect();
            if (IsExist)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        #region //ISOZONE_ISOSERCH
        //private void ISOZONE_ISOSERCH(string ISO)
        //{
        //    try
        //    {
        //        string Sopname = "";
        //        SOPURL.Clear();
        //        string sqlCmd = "SELECT * FROM [isozone_db].[dbo].[doc_esop_view] WHERE [document_no] = '" + ISO + "'  ";
        //        DataSet ds = db_IsoZoen.reDs(sqlCmd);
        //        if (ds.Tables[0].Rows.Count > 0)
        //        {
        //            Sopname = ds.Tables[0].Rows[0]["document_name"].ToString().Trim();
        //            listBox1.Items.Add(Sopname);
        //            txt_Ver1.Text = ds.Tables[0].Rows[0]["document_version"].ToString().Trim();
        //            SOPURL.Add(ds.Tables[0].Rows[0]["file_path"].ToString().Trim());
        //            string insSql;

        //            insSql = " INSERT INTO E-SOP_Mapping_Table (Search_Time,Model,WO,Sop_Name,Iso_Number,Process,Ver) VALUES("
        //                                   + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
        //                                   + "'" + txt_Model.Text + "',"
        //                                   + "'" + txt_WO.Text + "',"
        //                                   + "'" + Sopname + "',"
        //                                   + "'" + ds.Tables[0].Rows[0]["document_no"].ToString().Trim() + "',"
        //                                   + "N'" + dpd_Process.SelectedItem.ToString() + "',"
        //                                   + "'" + ds.Tables[0].Rows[0]["document_version"].ToString().Trim() + "')";
        //            if (db.Exsql(insSql) == true)
        //            {

        //            }
        //            else
        //            {

        //                addListItem("Mapping_Table資料庫寫入失敗");
        //            }
        //        }
        //        else
        //        {
        //            //Sopname = "";
        //            //listBox1.Items.Add(Sopname);
        //            addListItem("無相關SOP資料!!");

        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        addListItem("無相關SOP資料!!");
        //    }
        //}
        #endregion

        private string ISOZONE_SOPNAMESERCH(string SOPNAME)
        {
            string SopPath = "";
            try
            {
                
                string sqlCmd = "SELECT * FROM [isozone_db].[dbo].[doc_esop_view] WHERE [document_name] = N'" + SOPNAME + "'  ";
                DataSet ds = db_IsoZoen.reDs(sqlCmd);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    SopPath = ds.Tables[0].Rows[0]["file_path"].ToString().Trim();
                    txt_Ver1.Text = ds.Tables[0].Rows[0]["document_version"].ToString().Trim();
                    return SopPath;
                }
                else
                {
                    
                    
                    return SopPath;
                }
            }
            catch
            {
                return SopPath;
            }
        }
        private void All_Clear()
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            
            txt_Ver1.Text = "";
            txt_Model.Text = "";
            label1.Text = "";
            label2.Text = "";
            Production.btn_producN = "";
        }
        private void addListItem(string value)
        {
            this.listBox2.ForeColor = Color.Red;
            this.listBox2.Items.Add(value);
        }
        private void Getftp(string Ftp_Server_name)//ftp資訊
        {


            string sqlCmd = "SELECT [Ftp_Server_OA_Ip],[Ftp_Username],[Ftp_Password],[Ftp_Server_name] FROM i_Program_FtpServer_Table where [Ftp_Server_name] ='" + Ftp_Server_name + "' ";
            DataSet ds = db.reDs(sqlCmd);
            if (ds.Tables[0].Rows.Count != 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ftpServer = ds.Tables[0].Rows[0]["Ftp_Server_OA_Ip"].ToString().Trim();
                    ftpuser = ds.Tables[0].Rows[0]["Ftp_Username"].ToString().Trim();
                    ftppassword = ds.Tables[0].Rows[0]["Ftp_Password"].ToString().Trim();
                }
            }
        }
        #region //SearchTempSOP
        //private void SearchTempSOP( string SOP_Name)//搜尋暫時sop
        //{
        //    string sqlCmd = "SELECT [Sop_Name] FROM E-SOP_Temp_Table where [Sop_Name] ='" + SOP_Name.Trim() + "'  and Sign_off_Date IS NOT NULL";
        //    DataSet ds = db.reDs(sqlCmd);


        //    if (ds.Tables[0].Rows.Count > 0)
        //    {
        //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
        //        {
        //            SopName = ds.Tables[0].Rows[i]["SOP_NAME"].ToString();
        //            listBox1.Items.Add(SopName);
        //            SOPURL.Add(SopName+".pdf");
        //            string insSql;
        //            insSql = " INSERT INTO E-SOP_Mapping_Table (Search_Time,Model,WO,Sop_Name,Iso_Number,Process) VALUES("
        //                                   + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
        //                                   + "'" + txt_Model.Text + "',"
        //                                   + "'" + txt_WO.Text + "',"
        //                                   + "'" + SopName + "',"
        //                                   + "'Temp',"
        //                                   + "N'" + dpd_Process.SelectedItem.ToString() +  "')";
        //            if (db.Exsql(insSql) == true)
        //            {

        //            }
        //            else
        //            {

        //                addListItem("Mapping_Table資料庫寫入失敗");
        //            }

        //        }
        //    }
        //    else
        //    {
        //        addListItem("無相關SOP資料!!");
        //        MessageBox.Show("無相關SOP資料!!");
        //    }
        //}
        #endregion
        private string selectVerSQL_new(string tool)//Version Check new
        {
            string sqlCmd = "";
            bool result = false;
            try
            {
                sqlCmd = "select *  FROM TE_Program_Table where [Program_Name] ='" + tool + "'";
                DataSet ds = db.reDs(sqlCmd);
                if (ds.Tables[0].Rows.Count != 0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        if (ds.Tables[0].Columns[i].ToString() == "Version")
                        {
                            version_new = ds.Tables[0].Rows[0][i].ToString();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("更新異常");
            }
            return version_new;
        }
        public void autoupdate()//自動更新
        {
            //寫入目前版本與程式名後執行更新

            Process p = new Process();
            p.StartInfo.FileName = System.Windows.Forms.Application.StartupPath + "\\AutoUpdate.exe";
            p.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath; //檔案所在的目錄
            p.Start();
            this.Close();            

        }
    }
    
}
