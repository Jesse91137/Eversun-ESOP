using System;
using System.Windows.Forms;

namespace E_SOP
{
    /// <summary>
    /// WebBrowser 視窗表單，負責顯示網頁並執行資料庫記錄。
    /// </summary>
    public partial class WebBrowser : Telerik.WinControls.UI.RadForm
    {
        /// <summary>
        /// 製程路徑代碼
        /// </summary>
        public static string Route = "";

        /// <summary>
        /// 工單號碼
        /// </summary>
        public static string Wip_No = "";

        /// <summary>
        /// 機種名稱
        /// </summary>
        public static string Model = "";

        /// <summary>
        /// 建構函式，初始化元件。
        /// </summary>
        public WebBrowser()
        {
            InitializeComponent(); // 初始化表單元件
        }

        /// <summary>
        /// 表單載入事件，導向指定網址。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void WebBrowser_Load(object sender, EventArgs e)
        {
            // 導向 QA Zone 指定查詢網址
            webBrowser1.Navigate("http://qazone.avalue.com.tw/qazone/sfislinktopp.aspx?QA_MFID=YS00&QA_PRDID=" + Model + "&" + "QA_ROUTEID=" + Route);
        }

        /// <summary>
        /// 網頁載入完成事件，判斷是否查無資料。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // 檢查網頁是否載入完成
            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
            {
                // 判斷網頁內容是否包含「查無相關資料」
                if (webBrowser1.DocumentText.IndexOf("查無相關資料") > 0)
                {
                    // 關閉視窗並回傳 OK
                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    // 隱藏訊息標籤
                    lbl_msg.Visible = false;
                }
            }
        }

        /// <summary>
        /// 按鈕點擊事件，執行資料庫記錄並關閉視窗。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void RadButton1_Click(object sender, EventArgs e)
        {
            try
            {
                insertSQLScan(); // 執行資料庫記錄
                DialogResult = DialogResult.OK; // 設定回傳結果為 OK
                Close(); // 關閉視窗
            }
            catch
            {
                DialogResult = DialogResult.Cancel; // 設定回傳結果為 Cancel
            }
        }

        /// <summary>
        /// 新增一筆產品資訊檢查記錄到資料庫。
        /// </summary>
        public void insertSQLScan()
        {
            try
            {
                string insSql;
                insSql = "";

                // 組合 SQL 新增語法
                insSql = " INSERT INTO [iFactory].[E-SOP].[E-SOP_Product_Information_Check_Table] (Record_Time,Model,Name,Wip_No,Process) VALUES("
                        + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                        + "'" + Model + "',"
                        + "N'" + txt_Check_Name.Text.Trim() + "',"
                        + "N'" + Wip_No + "',"
                        + "N'" + Route + "')";

                // 執行 SQL 新增
                if (db.Exsql(insSql) == true)
                {
                    // 新增成功不做事
                }
                else
                {
                    // 新增失敗顯示訊息
                    MessageBox.Show("Windows資料庫上傳失敗");
                }
            }
            catch (Exception)
            {
                // 發生例外顯示訊息
                MessageBox.Show("Windows資料庫上傳失敗");
            }
        }

        /// <summary>
        /// 計時器事件，更新系統時間顯示。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void Timer1_Tick(object sender, EventArgs e)
        {
            // 顯示目前系統時間
            System_date_ID.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
    }
}
