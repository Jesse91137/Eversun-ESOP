using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using System.Threading;

namespace E_SOP
{
    public partial class WebBrowser : Telerik.WinControls.UI.RadForm
    {
        public static string Route = "";
        public static string Wip_No = "";
        public static string Model = "";
        public WebBrowser()
        {
            InitializeComponent();
        }

        private void WebBrowser_Load(object sender, EventArgs e)
        {
          
            webBrowser1.Navigate("http://qazone.avalue.com.tw/qazone/sfislinktopp.aspx?QA_MFID=YS00&QA_PRDID=" + Model+"&"+"QA_ROUTEID="+ Route);
        }

       

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
            {

                if(webBrowser1.DocumentText.IndexOf("查無相關資料") >0)
                {
                   
                    //Thread.Sleep(100);
                    DialogResult = DialogResult.OK;
                    Close();
                    
                }
                else
                {
                    lbl_msg.Visible = false;
                }

            }
        }

        private void RadButton1_Click(object sender, EventArgs e)
        {
            try
            {
                insertSQLScan();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch
            {
                DialogResult = DialogResult.Cancel;
            }
           
        }
        public void insertSQLScan() //建立資料庫
        {
            try
            {
                string insSql;
                insSql = "";
                //Dobule_Image_scanon();

                //insSql = compareinsSql + compareinsupdate + " INSERT INTO [scanvirus_testprglist] (time,status,_96level,model,series,version,Bits,Language,pename,penote,status_fix,ext_status,ext_item,cp_final,os,function_name,station_no,Factory,Avirachk_log,Avira_ver,Sophoschk_log,Scan_check,_96model,Linux_Sophoschk_Log,Sophos_ver) VALUES("

                insSql = " INSERT INTO [iFactory].[E-SOP].[E-SOP_Product_Information_Check_Table] (Record_Time,Model,Name,Wip_No,Process) VALUES("
                                       + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                       + "'" + Model + "',"
                                       + "N'" + txt_Check_Name.Text.Trim() + "',"
                                       + "N'" + Wip_No + "',"
                                       + "N'" + Route + "')";




                if (db.Exsql(insSql) == true)
                {
                   
                }
                else
                {
                    MessageBox.Show("Windows資料庫上傳失敗");
                }
            }
            catch (Exception aa)
            {
               
                MessageBox.Show("Windows資料庫上傳失敗");
            }
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            System_date_ID.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
    }
}
