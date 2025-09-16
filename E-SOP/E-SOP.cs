using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace E_SOP
{
    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        //宣告變數
        #region

        /// <summary>
        /// 主視窗共用的 ListBox，供其他類別或靜態存取主視窗顯示之訊息用。
        /// </summary>
        public static ListBox MainListBox;

        /// <summary>
        /// 用來儲存接受到的 Socket client 連線物件。
        /// </summary>
        public Socket newclient;

        /// <summary>
        /// 表示目前是否已經連線的旗標。
        /// true 表示已連線，false 表示未連線。
        /// </summary>
        public bool Connected;

        /// <summary>
        /// 背景執行緒參考，通常用於處理接收或監聽等長時間工作。
        /// </summary>
        public Thread myThread;

        /// <summary>
        /// 代表一個委派: 用來執行接受字串並在 UI 或其他地方回呼的委派定義。
        /// 方法簽名: void Method(string str)
        /// </summary>
        public delegate void MyInvoke(string str);

        /// <summary>
        /// FTP 客戶端實例，用於與 FTP 伺服器進行檔案上傳/下載操作。
        /// </summary>
        FTPClient ftp_Con = new FTPClient();

        /// <summary>
        /// 用於顯示或記錄目前時間的字串 (格式: yyyy-MM-dd HH:mm)。
        /// 在建立時會以系統現在時間初始化。
        /// </summary>
        string now = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm");

        /// <summary>
        /// 目前電腦名稱，從系統資訊擷取 (ComputerName)。
        /// </summary>
        string com_name = System.Windows.Forms.SystemInformation.ComputerName; //取電腦名稱

        /// <summary>
        /// 用於單一執行個體檢查的 Mutex。
        /// 若程式需避免重複執行，會以此物件做為判斷依據。
        /// </summary>
        static Mutex m;

        /// <summary>
        /// 預設下載路徑，使用應用程式啟動目錄 (Application.StartupPath)。
        /// </summary>
        string download_Path = System.Windows.Forms.Application.StartupPath;

        /// <summary>
        /// 功能參數: 若為特殊設定 (如 "ON")，表示下載時不改檔名等自定行為。
        /// 由 Setup.ini 的 CHANGE_FILE.FOUNCTION 讀取。
        /// </summary>
        string Fun = "";//下載不改檔名

        /// <summary>
        /// 目前 PDF 名稱暫存欄位，供下載或顯示 PDF 時使用。
        /// </summary>
        string PDF_Name = "";

        /// <summary>
        /// 自動開啟設定的旗標或模式，若為特定值表示僅查詢不下載。
        /// 由 Setup.ini 的 AUTO_OPEN 讀取並決定行為。
        /// </summary>
        string autoOpen = "";//不下載,只作查詢用

        /// <summary>
        /// SOP 相關路徑或名稱暫存 (範例: 網路 SOP 路徑、SOP 檔名等)。
        /// </summary>
        public string sop_path = "", SopName = "", version_old = "", version_new = "", filetempPath = "";

        /// <summary>
        /// 其他設定值或旗標：repeat 用於重複設定，continous 用於連續處理相關的設定。
        /// </summary>
        string repeat, continous = "";

        /// <summary>
        /// 臨時 DataSet 物件，供內部查詢或串接資料使用。
        /// </summary>
        DataSet oDS = new DataSet();

        /// <summary>
        /// 另一本地 DataSet，用於儲存 SOP 或相關資料查詢結果。
        /// </summary>
        DataSet oDSop = new DataSet();

        /// <summary>
        /// 用於儲存 SOP 查詢結果的 DataSet。
        /// </summary>
        DataSet dsSop = new DataSet();

        /// <summary>
        /// 用於儲存 SOP TwiceFail 或其他分類資料的 DataSet。
        /// </summary>
        DataSet dsSopTF = new DataSet();

        /// <summary>
        /// 裝 Excel 資料
        /// </summary>
        DataSet dsData = new DataSet();

        /// <summary>
        /// 儲存多個 SOP URL 或路徑的清單，常用於批次處理或檢視。
        /// </summary>
        public List<string> SOPURL = new List<string>();

        /// <summary>
        /// FTP 帳戶與伺服器相關資訊欄位 (由 Getftp 讀入)。
        /// ftpuser: 帳號, ftppassword: 密碼, ftpServer: 主機位址。
        /// </summary>
        public string ftpuser = "", ftppassword = "", ftpServer = "";

        /// <summary>
        /// SFIS 或系統回傳的機種與各種 SOP 路徑/說明欄位。
        /// 以下欄位對應 SFIS 的欄位名稱，可能在程式中由 WebService 或 API 填入：
        /// productid: 機種編號
        /// mfc_sop: 製造 SOP 路徑
        /// mfc_sop_desc: 製造 SOP 描述
        /// testing_sop: 測試 SOP 路徑
        /// testing_sop_desc: 測試 SOP 描述
        /// packing_sop: 包裝 SOP 路徑
        /// packing_sop_desc: 包裝 SOP 描述
        /// assy_sop: 組裝 SOP 路徑
        /// assy_sop_desc: 組裝 SOP 描述
        /// spe_packing_sop: 特殊包裝 SOP 路徑
        /// spe_packing_sop_desc: 特殊包裝 SOP 描述
        /// </summary>
        public string productid
            , mfc_sop
            , mfc_sop_desc
            , testing_sop
            , testing_sop_desc
            , packing_sop
            , packing_sop_desc
            , assy_sop
            , assy_sop_desc
            , spe_packing_sop
            , spe_packing_sop_desc;
        //------------------------------------------------------------------------------------------
        /// <summary>
        /// ini 檔案
        /// </summary>
        public string filename = "Setup.ini";
        SetupIniIP ini = new SetupIniIP();

        // 定時刪除過期檔案的計時器（背景執行）
        private System.Threading.Timer _deleteOldFilesTimer;
        // 預設間隔：1 小時
        private TimeSpan _deleteOldFilesInterval = TimeSpan.FromHours(1);

        //WebService
        QueryService.QueryServiceSoapClient queryService = new QueryService.QueryServiceSoapClient("QueryServiceSoap");

        #endregion

        /// <summary>
        /// 嘗試刪除整個資料夾，會重試多次並處理被鎖定或唯讀的檔案。
        /// 若遇到正在使用的檔案，將使用 WipeFile 做覆寫刪除保底處理。
        /// 此方法適合在背景執行，避免阻塞 UI。
        /// </summary>
        /// <param name="directoryPath"></param>
        private static void TryDeleteDirectoryWithRetries(string directoryPath)
        {
            // 檢查目錄是否存在
            if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath)) return;

            // 最大重試次數與等待時間
            const int maxAttempts = 5;
            const int delayMs = 1000; // 每次重試等待時間

            // 嘗試多次刪除
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    // 先嘗試遞迴刪除所有檔案
                    foreach (string file in Directory.GetFiles(directoryPath))
                    {
                        try
                        {
                            if (File.Exists(file))
                            {
                                FileInfo fi = new FileInfo(file);
                                if ((fi.Attributes & FileAttributes.ReadOnly) != 0)
                                {
                                    fi.Attributes = FileAttributes.Normal;
                                }

                                if (IsFileInUse(file))
                                {
                                    // 若檔案被鎖定，嘗試使用 WipeFile 覆寫再刪除
                                    WipeFile(file, 3);
                                }
                                else
                                {
                                    File.Delete(file);
                                }
                            }
                        }
                        catch (Exception exFile)
                        {
                            Console.WriteLine($"TryDeleteDirectoryWithRetries - 刪除檔案失敗: {file}, {exFile.Message}");
                        }
                    }

                    // 刪除子目錄（若有）
                    foreach (string dir in Directory.GetDirectories(directoryPath))
                    {
                        try
                        {
                            TryDeleteDirectoryWithRetries(dir);
                        }
                        catch (Exception exDir)
                        {
                            Console.WriteLine($"TryDeleteDirectoryWithRetries - 刪除子資料夾失敗: {dir}, {exDir.Message}");
                        }
                    }

                    // 嘗試刪除空的目錄
                    Directory.Delete(directoryPath, false);
                    Console.WriteLine($"已刪除資料夾: {directoryPath}");
                    break; // 成功，跳出重試迴圈
                }
                catch (IOException ioEx)
                {
                    Console.WriteLine($"TryDeleteDirectoryWithRetries - IO 錯誤 (嘗試 {attempt}): {ioEx.Message}");
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    Console.WriteLine($"TryDeleteDirectoryWithRetries - 權限錯誤 (嘗試 {attempt}): {uaEx.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"TryDeleteDirectoryWithRetries - 未知錯誤 (嘗試 {attempt}): {ex.Message}");
                }

                // 若尚未成功，等待後重試
                Thread.Sleep(delayMs);
            }
        }

        /// <summary>
        /// 建構函式，初始化 RadForm1 物件。
        /// </summary>
        public RadForm1()
        {
            InitializeComponent(); // 初始化元件
        }

        /// <summary>
        /// 主視窗載入事件。初始化各項設定、判斷程式是否重複執行、版本更新、資料夾建立、預設部門設定等。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void RadForm1_Load(object sender, EventArgs e)
        {

            // 判斷程式是否重複執行（Mutex）
            if (IsMyMutex("E-SOP"))
            {
                MessageBox.Show("程式正在執行中!!");
                Dispose(); // 關閉程式
            }
            // 讀取舊版本號
            version_old = ini.IniReadValue("Version", "version", filename);
            // 取得新版本號
            version_new = selectVerSQL_new("E-SOP");

            // 轉換版本號為整數以便比較
            int v_old = Convert.ToInt32(version_old.Replace(".", ""));
            int v_new = Convert.ToInt32(version_new.Replace(".", ""));

            // 顯示目前版本號於標籤
            lbl_ver.Text = "VER:V" + version_old;

            // 判斷是否需要自動更新
            if (v_old != v_new)
            {
                MessageBox.Show("有版本更新VER: V" + version_new);
                autoupdate(); // 執行自動更新
            }
            else
            {
                string defaultDept = "";
                // 讀取各項設定值
                repeat = ini.IniReadValue("REPEAT_SET", "SET", filename);
                autoOpen = ini.IniReadValue("AUTO_OPEN", "SET", filename);
                PDF_Name = ini.IniReadValue("PDF_NAME", "NAME", filename);
                continous = ini.IniReadValue("CONTINUOUS_SET", "SET", filename);
                Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", filename);
                defaultDept = ini.IniReadValue("Default_Type", "Type_Name", filename);
                // 關閉 ListBox 跨執行緒檢查
                ListBox.CheckForIllegalCrossThreadCalls = false;
                // 讀取資料庫設定
                string count = ini.IniReadValue("DataBase", "DataBaseCount", filename);

                /*
                 * SOP文件夾位置
                 */
                #region SOP文件夾位置
                /* \\192.168.4.53\esop公用區\ */
                string sqlpath = @"select Sop_Path from E_SOP_MappingEversun_Table where Process='SopPath'";
                DataSet dspath = db.reDs(sqlpath);

                /* \\192.168.4.11\01廠務處管理資料\02-工程部資料\10-MEMO發文\MEMO\ */
                string sqlemo = @"select Sop_Path from E_SOP_MappingEversun_Table where Process='EMO'";
                DataSet dsemo = db.reDs(sqlemo);

                // 取得電腦 domain 名稱，判斷是否加入 AD
                string domain = (string)Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "Domain", null);
                // 根據 domain 設定 SOP 路徑
                Production.soppath = (!string.IsNullOrEmpty(domain)) ? @"\\miss08\esop公用區\" : dspath.Tables[0].Rows[0][0].ToString();
                Production.emo = (!string.IsNullOrEmpty(domain)) ? @"\\miss01\01廠務處管理資料\02-工程部資料\10-MEMO發文\MEMO\" : dsemo.Tables[0].Rows[0][0].ToString();
                #endregion

                // 關閉量產/試產選項
                rad_mass.Enabled = false;
                rad_trial.Enabled = false;

                /* 載入製程下拉選單 */
                string sqlstr = @"select ID,Process,Groups from E_SOP_Process_Table";
                DataSet ds = db.reDs(sqlstr);
                // 載入製程下拉選單
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    dpd_Process.Items.Add(new MyItem(row["Process"].ToString(), row["ID"].ToString()));
                    dpd_Process.SelectedIndex = 0;
                }

                // 設定預設部門
                if (!string.IsNullOrEmpty(defaultDept))
                {
                    DefaultDept(defaultDept);
                }
                // 檢查 Temp 資料夾是否存在，若無則建立
                string tempFilePath = Application.StartupPath + "\\Temp";

                if (!Directory.Exists(tempFilePath))
                {
                    // 新增資料夾
                    Directory.CreateDirectory(@tempFilePath);
                }

                // 檢查 Temp 資料夾是否存在再處理
                string tempDir = Path.Combine(Application.StartupPath, "Temp");
                if (Directory.Exists(tempDir))
                {
                    DirectoryInfo faildi = new DirectoryInfo(tempDir);
                    FileInfo[] files = faildi.GetFiles();
                    int filecount = files.Length;
                    // 若超過閾值，快速切換資料夾名稱並在背景刪除舊資料夾，以提高啟動與回應速度
                    if (filecount >= 50)
                    {
                        try
                        {
                            // 以唯一名稱重命名舊 Temp 資料夾
                            string backupTemp = Path.Combine(Application.StartupPath, "Temp_Cleanup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N"));
                            Directory.Move(tempDir, backupTemp);

                            // 立即建立新的 Temp 資料夾以供程式繼續使用
                            Directory.CreateDirectory(tempDir);

                            // 在背景執行刪除舊資料夾，避免阻塞 UI
                            Thread cleanupThread = new Thread(() => TryDeleteDirectoryWithRetries(backupTemp));
                            cleanupThread.IsBackground = true;
                            cleanupThread.Start();
                        }
                        catch (Exception ex)
                        {
                            // 若重命名失敗，回退到逐一嘗試刪除的保守策略
                            Console.WriteLine($"Temp 資料夾快速切換失敗，回退逐檔刪除: {ex.Message}");
                            foreach (var fi in files)
                            {
                                try
                                {
                                    // 如果檔案正在使用，跳過刪除
                                    if (IsFileInUse(fi.FullName))
                                    {
                                        Console.WriteLine($"檔案正在使用，略過: {fi.FullName}");
                                        continue;
                                    }

                                    if ((fi.Attributes & FileAttributes.ReadOnly) != 0)
                                    {
                                        fi.Attributes = FileAttributes.Normal;
                                    }

                                    fi.Delete();
                                }
                                catch (IOException)
                                {
                                    Console.WriteLine($"IO 錯誤，刪除失敗，略過: {fi.FullName}");
                                    continue;
                                }
                                catch (UnauthorizedAccessException)
                                {
                                    Console.WriteLine($"權限不足，無法刪除，略過: {fi.FullName}");
                                    continue;
                                }
                                catch (Exception ex2)
                                {
                                    Console.WriteLine($"刪除檔案失敗: {fi.FullName}，錯誤: {ex2.Message}");
                                    continue;
                                }
                            }
                        }
                    }
                }
                // 啟動定時刪除過期檔案（會先執行一次，之後週期性執行）
                StartDeleteOldFilesTimer();
                // 設定按鈕文字
                Btn_WO.Text = "工程編號找SOP";
                Btn_ProduNotice.Text = "生產通知單" + Environment.NewLine + "與 ECN 文件查詢";
            }
        }

        /// <summary>
        /// 設定預設製程部門，根據傳入的部門名稱 D，自動選擇對應的下拉選單索引。
        /// </summary>
        /// <param name="D">部門名稱（如 QA、SMD、DIP、功能測試、系統組裝、包裝、Coating）</param>
        public void DefaultDept(string D)
        {
            // 根據部門名稱設定 dpd_Process 的選取索引
            switch (D)
            {
                case "QA":
                    dpd_Process.SelectedIndex = 1; // QA 對應索引 1
                    break;
                case "SMD":
                    dpd_Process.SelectedIndex = 2; // SMD 對應索引 2
                    break;
                case "DIP":
                    dpd_Process.SelectedIndex = 3; // DIP 對應索引 3
                    break;
                case "功能測試":
                    dpd_Process.SelectedIndex = 4; // 功能測試 對應索引 4
                    break;
                case "系統組裝":
                    dpd_Process.SelectedIndex = 5; // 系統組裝 對應索引 5
                    break;
                case "包裝":
                    dpd_Process.SelectedIndex = 6; // 包裝 對應索引 6
                    break;
                case "Coating":
                    dpd_Process.SelectedIndex = 7; // Coating 對應索引 7
                    break;
                default:
                    break; // 其他部門不做任何設定
            }
        }

        /// <summary>
        /// 打開指定路徑的檔案，並委派給共用工具 FileUtils 進行實際開啟處理。
        /// </summary>
        /// <param name="filePath">
        /// 要開啟的檔案完整路徑。可接受本機路徑或可由系統資源存取的 UNC/URI 路徑字串。
        /// 傳入前務必確保路徑不為 null 或空字串，並且應具有適當的存取權限。
        /// </param>
        /// <remarks>
        /// 此方法僅作為單一入口（facade），將實際的檔案開啟責任轉發到 FileUtils.OpenFile()，
        /// 以維持單一實作並利於日後統一處理（例如：統一更新時間戳記、例外處理或日誌）。
        /// 若需擴充日誌或錯誤處理，請在此層加入 try/catch 並呼叫共用的錯誤記錄方法。
        /// </remarks>
        private void OpenFile(string filePath)
        {
            FileUtils.OpenFile(filePath);
        }

        /// <summary>
        /// 處理 txt_SN 的 KeyPress 事件，當按下 Enter 鍵（KeyChar=13）時，
        /// 會觸發 Btn_WO_Click 方法，執行工程編號查詢流程。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">KeyPressEventArgs，包含按鍵資訊。</param>
        private void Txt_SN_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 如果按下的鍵是 Enter（KeyChar=13）
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                // 執行工程編號查詢流程
                Btn_WO_Click(sender, e);
            }
        }

        /// <summary>
        /// Timer1_Tick 事件處理函式。
        /// 於每次計時器觸發時，更新 System_date_ID 控制項的文字內容，顯示目前的日期與時間。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Timer1_Tick(object sender, EventArgs e)
        {
            // 更新 System_date_ID 的文字，顯示目前時間
            System_date_ID.Text = "Timer: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        }

        /// <summary>
        /// 依據工程編號查詢 PCB 料號(生產履程-檔案名稱)，並記錄到 Production.pcb。
        /// 有一支應用程式叫[E_SOP_PCB_Table]，人工執行更新資料表[E_SOP_PCB_Table]。目前由[工程部-黃寶雅]維護。
        /// </summary>
        public void PCB_item()
        {
            // 從 E_SOP_PCB_Table 依據 txt_EngSr.Text 查詢 PCB_item 欄位
            DataSet ds = db.reDs("select PCB_item from E_SOP_PCB_Table where Eng_SR='" + txt_EngSr.Text.Trim() + "'");
            // 如果查詢結果有資料
            if (ds.Tables[0].Rows.Count > 0)
            {
                // 取得查詢到的 PCB_item 並記錄到 Production.pcb
                Production.pcb = ds.Tables[0].Rows[0][0].ToString();
            }
        }

        /// <summary>
        /// 2023 Q1 於\\miss01\全廠共用\31-SMD拋料率及不易維修紀錄，找到不易維修Excel與照片開新視窗列出
        /// 在 SMD 拋料率與不易維修紀錄中，依機種 (engsr) 搜尋不易維修零件清單並顯示結果視窗。每一行皆有中文註解。
        /// </summary>
        /// <param name="engsr">要查詢的機種編號，方法會以大寫比對資料表欄位「機種」。</param>
        public void SMD_Case(string engsr)
        {
            // Excel 檔案所在網路路徑（含檔名）。
            string excelPath = @"\\192.168.4.11\全廠共用\31-SMD拋料率及不易維修紀錄\02-不易維修零件機種\不易維修零件機種.xls";
            // 照片資料夾路徑（目前未在方法內使用，但保留以備擴充）。
            string picPath = @"\\miss01\全廠共用\31-SMD拋料率及不易維修紀錄\01-不易維修照片\";

            // 讀取指定的 Excel 檔案並將資料載入成員變數 dsData（呼叫既有方法）。
            LoadExcel(excelPath); // 讀取 Excel 並把結果放到 this.dsData。

            // 取得讀取後的第一個 DataTable（假設資料存在於 Tables[0]）。
            DataTable dt_E = dsData.Tables[0]; // 將 dsData 的第一個表格指派給 dt_E 以便後續搜尋。

            // 確保輸入的 engsr 不為 null/空並進行處理
            if (string.IsNullOrWhiteSpace(engsr)) // 如果傳入機種為空，直接顯示訊息並結束。
            {
                MessageBox.Show("請提供要查詢的機種編號"); // 提示使用者輸入
                return; // 結束方法
            }

            try
            {
                // 使用 LINQ 依照欄位「機種」做過濾，並比對大寫後的值；若無符合列 CopyToDataTable 會拋 InvalidOperationException。
                DataTable result = dt_E.AsEnumerable() // 取得資料列的可查詢序列
                    .Where(o => (o.Field<string>("機種") ?? string.Empty).Equals(engsr.ToUpper(), StringComparison.Ordinal)) // 篩選機種欄位等於傳入 engsr（大寫比對）
                    .CopyToDataTable(); // 將篩選結果轉回 DataTable（若無列會拋例外）

                // 建立 ViewForm 視窗實例，準備顯示篩選結果
                ViewForm viewForm = new ViewForm(); // 新增一個顯示視窗的物件

                // 設定視窗的工程編號屬性（轉大寫並去除前後空白）
                viewForm.Engsr = engsr.ToUpper().Trim(); // 將工程編號傳入 ViewForm

                // 將查詢結果的 DataTable 指派給視窗（用於在視窗中顯示）
                viewForm.Tables = result; // 傳入結果資料表到 ViewForm 的 Tables 屬性

                // 呼叫 ViewForm 的 setValue()，將資料設定到表單內部 UI 或欄位
                viewForm.setValue(); // 將傳入資料應用到 ViewForm 的內部欄位與控制項

                // 以模態視窗顯示結果，使用者關閉後才會回到此程序
                viewForm.ShowDialog(); // 顯示視窗並等待使用者關閉
            }
            catch (InvalidOperationException)
            {
                // 當 CopyToDataTable 在沒有資料列時會拋出 InvalidOperationException，處理此情況顯示友善提示
                MessageBox.Show($"查無機種 {engsr} 的不易維修資料"); // 通知使用者查無資料
            }
            catch (Exception ex)
            {
                // 捕捉其他例外並記錄或顯示錯誤訊息
                MessageBox.Show("讀取不易維修資料時發生錯誤: " + ex.Message); // 顯示錯誤內容供排查
            }
        }

        /// <summary>
        /// 讀取指定 Excel 檔案並將資料載入 dsData。
        /// 支援 Office 97-2003 格式（.xls），如需 Office 2007 以上請參考註解。
        /// </summary>
        /// <param name="filename">Excel 檔案完整路徑</param>
        private void LoadExcel(string filename)
        {
            if (filename != "")
            {
                #region office 97-2003
                // 建立連線字串，使用 Microsoft.Jet.OLEDB.4.0 讀取 Excel 97-2003 格式
                string excelString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                                     "Data Source=" + filename + ";" +
                                                     "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1\'";
                #endregion

                #region office 2007
                //string excelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename +
                //                 ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
                #endregion

                // 建立 OleDbConnection 物件並開啟連線
                OleDbConnection cnn = new OleDbConnection(excelString);
                cnn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = cnn;
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                cmd.CommandText = "SELECT * FROM [Sheet1$]";
                adapter.SelectCommand = cmd;
                // 將 Excel 資料載入 dsData
                adapter.Fill(dsData);
                // 釋放資源
                cnn = null;
                cmd = null;
                adapter = null;
            }
        }

        /// <summary>
        /// 工程編號找SOP Enter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_WO_Click(object sender, EventArgs e)
        {
            All_Clear();
            // 製程
            if (dpd_Process.SelectedItem.ToString() == "請選擇...")
            {
                addListItem("未選擇製程!!");
                return;
            }

            // 工程編號
            if (!txt_EngSr.ReadOnly && !string.IsNullOrEmpty(txt_EngSr.Text))
            {
                // 試產/量產
                bool f = (!rad_trial.Checked && !rad_mass.Checked) ? true : false;
                if (f)
                {
                    listBox2.Items.Add("請選擇試產/量產");
                    return;
                }
            }

            if (txt_EngSr.ReadOnly || (string.IsNullOrEmpty(txt_EngSr.Text) && !string.IsNullOrEmpty(txt_WO.Text)))
            {
                string result = btnSN_Click();

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

            #region vip 最終客戶
            /* 應用程式[EndCustomer_TO_SOP_DB]，應該是[工程部]更新資料表[E_SOP_Eversun_Vip_Customer]*/
            string sqlvip = @" select * from E_SOP_Eversun_Vip_Customer where eng_sr ='" + txt_EngSr.Text + "' ";
            DataSet dsvip = db.reDs(sqlvip);
            if (dsvip.Tables[0].Rows.Count > 0)
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
                            /* 有一支應用程式叫[E_SOP_PCB_Table]，人工執行更新資料表[E_SOP_PCB_Table]。目前由[工程部-黃寶雅]維護。*/
                            PCB_item();
                            //WebClient ESOPclient = new WebClient();
                            //生產注意事項
                            if (Production.deptItem == "包裝")
                            {
                                DontTwiceFault(txtWo, FrontPath);
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

                                            // 使用共用 OpenFile 方法開啟並統一更新時間戳記
                                            OpenFile(localpath);

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

                                            // 使用共用 OpenFile 方法開啟並統一更新時間戳記
                                            OpenFile(localpath);

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
                        DataSet data = StationKeyWord(myProcess.value, dpd_station.SelectedItem.ToString());//20210318 StationKeyWord search

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
                            if (c == 0)
                            {
                                foreach (var fi in di.GetFiles())
                                {
                                    if (fi.Name.IndexOf("db") <= 0)
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

                                        // 使用共用 OpenFile 方法開啟並統一更新時間戳記
                                        OpenFile(localpath);
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

        /// <summary>
        /// 複製檔案到臨時目錄並返回新路徑
        /// </summary>
        /// <param name="path">來源檔案路徑</param>
        /// <returns></returns>
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

        /// <summary>
        /// 依據製程與站別查詢 SOP 關鍵字資料集。
        /// </summary>
        /// <param name="key">製程代碼（ProcessID）。</param>
        /// <param name="value">站別名稱（Station）。</param>
        /// <returns>
        /// 回傳包含 StationKeyWord 欄位的 DataSet，
        /// 依據製程與站別條件查詢 E_SOP_Station_Table。
        /// 若製程為 SMD 或 DIP，則額外包含「總檢圖」站別。
        /// 若製程為 QA，則依站別條件查詢，若站別為 All 則查詢全部。
        /// </returns>
        private DataSet StationKeyWord(string key, string value)
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
                if (value == "All")
                {
                    sqlstr = @"select StationKeyWord from E_SOP_Station_Table";
                }
            }
            DataSet ds = db.reDs(sqlstr);
            return ds;
        }

        /// <summary>
        /// 處理「包裝」製程的「不貳過」注意事項檔案開啟流程。
        /// 依據工程編號(txtWo)與客戶前碼(f_Path)，分別搜尋「組裝」與「後製程」目錄下的相關檔案。
        /// 若找到符合檔名的檔案，則複製到 Temp 資料夾並開啟，並顯示 InputBox 視窗供使用者確認。
        /// 同時也會搜尋 PCB 相關檔案並執行相同流程。
        /// </summary>
        /// <param name="txtWo">工程編號（如機種編號）</param>
        /// <param name="f_Path">客戶前碼（工程編號前兩碼）</param>
        private void DontTwiceFault(string txtWo, string f_Path)
        {
            InputBox inbox = new InputBox();
            //組裝
            DirectoryInfo assy = new DirectoryInfo(Production.pack_assy + "\\" + f_Path);//組裝
            DirectoryInfo a_proc = new DirectoryInfo(Production.pack_proc + "\\" + f_Path);//後製程
            System.Diagnostics.Process pro = new System.Diagnostics.Process();
            System.Diagnostics.Process pro2 = new System.Diagnostics.Process();//PCB
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

                        // 使用共用 OpenFile 方法開啟並統一更新時間戳記
                        OpenFile(localpath);

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


        /// <summary>
        /// 查詢機種名稱
        /// </summary>
        /// <param name="txtWO">工單號碼</param>
        /// <returns></returns>
        private string QueryDataSet(string txtWO)
        {
            string txt_SR = string.Empty;
            string sql = "select WO_NO,ENG_SR from SAP_DATA_TO_ESOP where WO_NO='" + txtWO + "' ";//SAP
            DataSet ds = db.reDs(sql);

            if (ds.Tables[0].Rows.Count > 0)//SAP
            {
                string[] temparry = ds.Tables[0].Rows[0]["ENG_SR"].ToString().Trim().Split('-');
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

            if (txt_SR.IndexOf("SERVICE") >= 0)
            {
                txt_SR = string.Empty;
            }
            return txt_SR;
        }

        /// <summary>
        /// 製程下拉選單選擇變更事件。
        /// 根據選取的製程，動態載入站別選項，並設定相關控制項狀態。
        /// QA 製程會載入所有站別並啟用量產/試產選項，其他製程則關閉相關選項。
        /// 若選取非「請選擇...」，則依據製程代碼載入對應站別。
        /// 最後將選取的製程名稱記錄到 Production.deptItem。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
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
            if (dpd_Process.SelectedItem.ToString() != "請選擇...")
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

        /// <summary>
        /// 判斷量產/試產路徑，並依據目前選取的製程設定 SOP 相關目錄。
        /// 1. 依據資料庫查詢取得 SOP 路徑（量產/試產），並設定 Production.sopfile。
        /// 2. 設定生產注意事項（不貳過）目錄 Production.twicefail。
        /// 3. 若製程為「包裝」，額外設定組裝與後製程目錄（Production.pack_proc, Production.pack_assy）。
        /// </summary>
        /// <param name="wono">工程編號（如機種編號）</param>
        /// <param name="frontpath">工程編號前兩碼（用於路徑組合）</param>
        public void Judg(string wono, string frontpath)
        {
            //Production.sopfile = (dssop.Tables[0].Rows.Count > 0) ? Production.soppath + dssop.Tables[0].Rows[0][0].ToString() : "";//(條件式) ?成立 :不成立
            //同時確認量產/試產工程編號是否存在檔案;同時存在以量產路徑為主
            string sqlstr = @"select sop_path from E_SOP_MappingEversun_Table ";
            //SOP FilePath
            string sqlsop = sqlstr + @" where process= 'SOP'";
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
            DataSet dsGroup = db.reDs(twiceGroup);
            Production.twicefail = (dstwice.Tables[0].Rows.Count > 0) ?
                Production.soppath + dstwice.Tables[0].Rows[0][0].ToString() + dsGroup.Tables[0].Rows[0][0].ToString() + @"\" : "";//(條件式) ?成立 :不成立

            //↓↓process→包裝 例外判斷
            if (myProcess.text == "包裝")
            {
                //SOP目錄+不貳過目錄+(包裝_組裝)
                Production.pack_proc = Production.soppath + dstwice.Tables[0].Rows[0][0].ToString() + @"組裝\";
                //SOP目錄+不貳過目錄+(包裝_後製程)
                Production.pack_assy = Production.soppath + dstwice.Tables[0].Rows[0][0].ToString() + @"後製程\";
            }
            #endregion
        }

        /// <summary>
        /// 生產通知單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                        if (listBox1.Items.Count == 0)
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

        /// <summary>
        /// 遞迴搜尋指定目錄及其所有子目錄下，
        /// 檔名包含工單號碼(txt_WO.Text)的檔案，
        /// 並將檔案名稱加入 listBox1，同時記錄完整路徑到 Production.filepathlist。
        /// </summary>
        /// <param name="dirPath">要搜尋的根目錄路徑。</param>
        public void FindFile(string dirPath)
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

        /// <summary>
        /// 依工單號碼(txt_WO.Text)呼叫 WebAPI 查詢 WipAtts 資料。
        /// 若 API 回傳 "error" 或 "無法連線WebAPI" 則顯示錯誤訊息視窗。
        /// 回傳 API 查詢結果字串。
        /// </summary>
        /// <returns>API 回傳的 WipAtts 資料字串，若失敗則回傳錯誤訊息。</returns>
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

        /// <summary>
        /// 雙擊檔案名稱時開啟對應檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            string filename = string.Empty;   // 存放選定的檔案名稱
            string path = string.Empty;       // 存放選定的檔案路徑
            string folder_F = string.Empty;   // 存放選定的檔案所在目錄
            string strrtn = string.Empty;     // 存放選定的檔案所在目錄
            string special = listBox1.SelectedItem.ToString();

            // 情況1: 生產通知單且不是特殊文件(EM或SERVICE)
            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0) //排除memo.service file
            {
                filename = listBox1.SelectedItem.ToString();
                path = Production.producnotice; ;   // 使用生產通知單路徑
                //Production.producnotice;
                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                folder_F = strrtn.Substring(0, 2);  // 取工程編號前兩個字作為資料夾首碼
            }
            // 情況2: 無工程編號搜尋模式
            else if (Production.btn_producN == "noEngSearch")
            {
                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
            }
            // 情況3: 特殊文件(如EM開頭或包含SERVICE的檔案)
            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
            {

                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
            }
            // 情況4: 一般SOP文件
            else
            {
                filename = listBox1.SelectedItem.ToString();
                path = Production.sopfile;
                folder_F = filename.Substring(0, 2);
                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
            }
            if (listBox1.SelectedItem.ToString() == "沒有生產通知單!!")
            {
                return;
            }

            #region OPEN ESOP
            // 組合完整的檔案路徑
            // 對於特殊文件直接使用path，對於一般文件則組合完整路徑
            string SOPpath = (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) ? path : path + folder_F + @"\" + strrtn + @"\" + filename;

            // 複製檔案到臨時目錄並返回新路徑
            SOPpath = openFile_copyFile(SOPpath.TrimEnd('\\'));

            // 使用系統預設應用程式開啟檔案
            // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
            OpenFile(SOPpath);
            #endregion

            // 顯示相對應的[AMES-維修資料統計] By 20250513 Jesse
            #region 維修資料統計
            // 維修統計資料集
            DataSet dsRepair = null;
            // 根據 責任單位 的值查詢維修統計資料
            string responsibilityUnit = string.Empty;
            // 機種名稱
            string deviceName = txt_EngSr.Text;
            // 檔案名稱
            string strfilename = string.Empty;
            // 站別名稱
            string strStationName = string.Empty;
            try
            {
                if (string.IsNullOrEmpty(deviceName))
                {
                    deviceName = QueryDataSet(txt_WO.Text);
                }

                if (!string.IsNullOrEmpty(deviceName))
                {
                    // 因查找維修統計 20250514 By Jesse
                    // 去除檔案名稱中的 _ 部分
                    strfilename = filename.Replace(deviceName + "_", "");
                    // 檢查檔案名稱中是否包含點
                    int dotIndex = strfilename.IndexOf('.');

                    if (dotIndex != -1) // 檢查字串中是否包含點
                    {
                        strfilename = strfilename.Substring(0, dotIndex);
                    }

                    // 根據檔案名稱判斷責任單位
                    responsibilityUnit = DetermineResponsibilityUnit(strfilename);

                    // 根據檔案名稱判斷站別名稱 By 20250911 Jesse
                    // 如果檔名或處理後的檔名包含 DIP，直接使用 strStationName 作為查詢條件
                    if ((!string.IsNullOrEmpty(filename) && filename.IndexOf("DIP", StringComparison.OrdinalIgnoreCase) >= 0)
                        || (!string.IsNullOrEmpty(strfilename) && strfilename.IndexOf("DIP", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        strStationName = "DIP";
                    }
                    if (!string.IsNullOrWhiteSpace(responsibilityUnit) || !string.IsNullOrWhiteSpace(strStationName))
                    {
                        // 查詢資料庫
                        dsRepair = QueryRepairRecords(deviceName, responsibilityUnit, strStationName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查詢維修紀錄時發生錯誤：{ex.Message}");
            }
            #endregion

            // 先顯示維修統計視窗（如有資料）
            if (dsRepair != null && dsRepair.Tables.Count > 0 && dsRepair.Tables[0].Rows.Count > 0 && !string.IsNullOrWhiteSpace(deviceName))
            {
                ShowRepairRecordsForm(dsRepair, deviceName);
            }

        }

        #region 維修資料統計 20250514 By Jesse
        /// <summary>
        /// 根據檔案名稱判斷對應的責任單位
        /// </summary>
        /// <param name="filename">檔案名稱</param>
        /// <returns></returns>
        private string DetermineResponsibilityUnit(string filename)
        {
            // 只用後置萬用字元，提升索引利用率
            string likePattern = filename + "%";
            string sql = @"SELECT TOP 1 ResponsibilityUnit 
                           FROM [dbo].[E_SOP_ResponsibilityMapping] 
                           WHERE IsActive = 1 AND [E-SOP] LIKE @FileName 
                           ORDER BY LEN([E-SOP]) DESC";

            Dictionary<string, object> parameters = new Dictionary<string, object>
            {
                { "@FileName", likePattern }
            };

            object result = db.ExecuteScalar(sql, parameters);

            // 如果查詢結果為空，則返回空字串
            return result?.ToString() ?? string.Empty;

        }

        /// <summary>
        /// 查詢維修統計資料
        /// </summary>
        /// <param name="deviceName">機種名稱</param>
        /// <param name="responsibilityUnit">責任單位</param>
        /// <param name="strStationName">站別名稱</param>
        /// <returns></returns>
        private DataSet QueryRepairRecords(string deviceName, string responsibilityUnit, string strStationName)
        {
            // 增加判斷站別是DIP By 20250911 Jesse
            // A.StationName /*站別*/

            // 若沒有提供機種名稱，直接回傳空的 DataSet，避免不必要的 DB 查詢
            if (string.IsNullOrWhiteSpace(deviceName))
            {
                return new DataSet();
            }

            string sql = @"SELECT DISTINCT A.FailureDate,/*不良日期*/
	                       A.ResponsibilityUnit,/*責任單位*/
                           A.RepairPartNumber,/*維修料號*/
                           A.RepairLocation,/*維修位置*/
                           A.FailureDescription,/*不良描述*/
                           A.Notes/*備註*/
                           FROM [dbo].[E_SOP_RepairRecords] A
                           WHERE A.DeviceName = @DeviceName ";

            // 只在有值時加入對應的條件與參數，避免傳遞空字串到 DB
            var parameters = new Dictionary<string, object>
            {
                { "@DeviceName", deviceName }
            };

            if (!string.IsNullOrWhiteSpace(responsibilityUnit) && !string.IsNullOrWhiteSpace(strStationName))
            {
                sql += " AND (A.ResponsibilityUnit = @ResponsibilityUnit OR A.StationName LIKE @StationName) ";
                parameters.Add("@ResponsibilityUnit", responsibilityUnit);
                parameters.Add("@StationName", "%" + strStationName + "%");
            }
            else
            if (!string.IsNullOrWhiteSpace(strStationName))
            {
                // 使用模糊比對，並在參數中帶上通配符以利索引使用
                sql += " AND A.StationName LIKE @StationName ";
                parameters.Add("@StationName", "%" + strStationName + "%");
            }
            else
            if (!string.IsNullOrWhiteSpace(responsibilityUnit))
            {
                sql += " AND A.ResponsibilityUnit = @ResponsibilityUnit ";
                parameters.Add("@ResponsibilityUnit", responsibilityUnit);
            }

            sql += " Order By A.FailureDate ";

            return db.reDsWithParams(sql, parameters);
        }

        /// <summary>
        /// 顯示維修紀錄視窗
        /// </summary>
        private void ShowRepairRecordsForm(DataSet dsRepair, string deviceName)
        {
            try
            {
                // 建立新視窗來顯示查詢結果
                Form repairForm = new Form();
                repairForm.Text = "維修統計 ";
                repairForm.Width = 800;
                repairForm.Height = 600;
                repairForm.StartPosition = FormStartPosition.CenterScreen;

                // 直接創建帶中文標題的新 DataTable
                DataTable filteredTable = new DataTable();
                // A.FailureDate,/*不良日期*/
                // A.ResponsibilityUnit/*責任單位*/
                // A.RepairPartNumber/*維修料號*/
                // A.RepairLocation,/*維修位置*/
                // A.FailureDescription,/*不良描述*/
                // A.Notes/*備註*/  */

                // 先添加欄位，使用中文名稱
                filteredTable.Columns.Add("不良日期", typeof(string));
                filteredTable.Columns.Add("責任單位", typeof(string));
                filteredTable.Columns.Add("維修料號", typeof(string));
                filteredTable.Columns.Add("維修位置", typeof(string));
                filteredTable.Columns.Add("不良描述", typeof(string));
                filteredTable.Columns.Add("備註", typeof(string));

                // 將原始資料複製到新表格中
                foreach (DataRow row in dsRepair.Tables[0].Rows)
                {
                    DataRow newRow = filteredTable.NewRow();
                    #region 對應原始欄位和新欄位
                    // 對應原始欄位和新欄位
                    // A.FailureDate,/*不良日期*/
                    if (dsRepair.Tables[0].Columns.Contains("FailureDate"))
                    {
                        newRow["不良日期"] = row["FailureDate"];
                    }
                    // A.ResponsibilityUnit/*責任單位*/
                    if (dsRepair.Tables[0].Columns.Contains("ResponsibilityUnit"))
                    {
                        newRow["責任單位"] = row["ResponsibilityUnit"];
                    }

                    // A.RepairPartNumber/*維修料號*/
                    if (dsRepair.Tables[0].Columns.Contains("RepairPartNumber"))
                    {
                        newRow["維修料號"] = row["RepairPartNumber"];
                    }

                    // A.RepairLocation,/*維修位置*/
                    if (dsRepair.Tables[0].Columns.Contains("RepairLocation"))
                    {
                        newRow["維修位置"] = row["RepairLocation"];
                    }

                    // A.FailureDescription,/*不良描述*/
                    if (dsRepair.Tables[0].Columns.Contains("FailureDescription"))
                    {
                        newRow["不良描述"] = row["FailureDescription"];
                    }

                    // A.Notes/*備註*/  */
                    if (dsRepair.Tables[0].Columns.Contains("Notes"))
                    {
                        newRow["備註"] = row["Notes"];
                    }
                    #endregion
                    filteredTable.Rows.Add(newRow);
                }

                // 創建 DataGridView 並設定基本屬性
                DataGridView dgv = new DataGridView();
                dgv.Dock = DockStyle.Fill;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.AllowUserToAddRows = false;
                dgv.ReadOnly = true;
                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgv.RowHeadersVisible = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                // 設定資料來源
                dgv.DataSource = filteredTable; // 使用過濾後的表格

                // 添加標題 Label
                Label titleLabel = new Label();
                titleLabel.Text = "機種名稱 : " + deviceName;
                titleLabel.Dock = DockStyle.Top;
                titleLabel.Font = new Font(titleLabel.Font.FontFamily, 14, FontStyle.Bold);
                titleLabel.Height = 30;
                titleLabel.TextAlign = ContentAlignment.MiddleCenter;

                // 添加控制項到表單
                repairForm.Controls.Add(dgv);
                repairForm.Controls.Add(titleLabel);

                // 調整 DataGridView 位置，使其在標題下方
                dgv.Location = new Point(0, titleLabel.Height);
                dgv.Height = repairForm.ClientSize.Height - titleLabel.Height;

                // 設定為最上層
                repairForm.TopMost = true;
                // 顯示表單
                repairForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("顯示維修紀錄時發生錯誤: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        /// <summary>
        /// 當 txt_SN 文字內容變更時觸發的事件處理函式。
        /// 主要用途為清空 txt_EngSr 的內容。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void txt_SN_TextChanged(object sender, EventArgs e)
        {
            // 當 txt_SN 的內容變更時，清空 txt_EngSr 的內容
            txt_EngSr.Text = "";
        }

        /// <summary>
        /// ListBox3 的 DrawItem 事件處理函式。
        /// 根據項目索引繪製不同顏色的項目文字。
        /// </summary>
        /// <param name="sender">觸發事件的 ListBox 控制項。</param>
        /// <param name="e">DrawItemEventArgs，包含繪製資訊。</param>
        private void listBox3_DrawItem(object sender, DrawItemEventArgs e)
        {
            ListBox list = (ListBox)sender; // 取得觸發事件的 ListBox 控制項

            e.DrawBackground(); // 繪製每個項目的背景

            Brush myBrush = Brushes.Black; // 預設畫筆顏色為黑色

            // 根據項目索引決定畫筆顏色
            switch (e.Index)
            {
                case 0:
                    myBrush = Brushes.Red; // 第一行用紅色
                    break;
                case 1:
                    myBrush = Brushes.Orange; // 第二行用橘色
                    break;
                case 2:
                    myBrush = Brushes.Purple; // 第三行用紫色
                    break;
            }

            // 使用指定的字型與顏色繪製目前項目的文字
            e.Graphics.DrawString(list.Items[e.Index].ToString(),
                e.Font, myBrush, e.Bounds, StringFormat.GenericDefault);

            e.DrawFocusRectangle(); // 如果 ListBox 有焦點，則繪製選取框
        }

        #region SOP Open Button
        /// <summary>
        /// Btn_Sop1_Click 事件處理函式。
        /// 處理 SOP1 按鈕點擊時的邏輯，根據設定檔開關與選取項目，
        /// 下載或開啟對應的 SOP 檔案，並依副檔名決定開啟方式。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Sop1_Click(object sender, EventArgs e)
        {
            // 宣告檔名變數
            string filename = string.Empty;
            // 宣告路徑變數
            string path = string.Empty;
            // 宣告資料夾前綴變數
            string folder_F = string.Empty;
            // 宣告工程編號變數
            string strrtn = string.Empty;
            // 讀取設定檔中的功能開關
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            // 判斷功能開關是否為 ON
            if (Fun == "ON")
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 檢查 SOPForm1 實例數量是否小於 1
                        if (SOPForm1.m_Static_InstanceCount < 1)
                        {
                            // 取得選取項目文字
                            string special = listBox1.SelectedItem.ToString();
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true)
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }

                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);

                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP1.pdf");
                                // 建立 SOPForm1 視窗
                                SOPForm1 Form = new SOPForm1();
                                // 實例數加一
                                SOPForm1.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm1.SOPName = "SOP1.pdf";
                                // 設定視窗標題
                                Form.Text = filename;
                                // 顯示 SOPForm1 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;
                                pro.Start();
                                // 手動更新最後訪問時間

                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        // SOPForm1 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP1");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            // 功能開關不是 ON
            else
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm1 實例數量是否小於 1
                        if (SOPForm1.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true)
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }

                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);

                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP1.pdf");
                                // 建立 SOPForm1 視窗
                                SOPForm1 Form = new SOPForm1();
                                // 實例數加一
                                SOPForm1.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm1.SOPName = "SOP1.pdf";
                                // 設定視窗標題
                                Form.Text = filename;
                                // 顯示 SOPForm1 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm1 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP1");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }

        /// <summary>
        /// Btn_Sop2_Click 事件處理函式。
        /// 處理 SOP2 按鈕點擊時的邏輯，根據設定檔開關與選取項目，
        /// 下載或開啟對應的 SOP 檔案，並依副檔名決定開啟方式。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Sop2_Click(object sender, EventArgs e)
        {
            // 宣告檔名變數
            string filename = string.Empty;
            // 宣告路徑變數
            string path = string.Empty;
            // 宣告資料夾前綴變數
            string folder_F = string.Empty;
            // 宣告工程編號變數
            string strrtn = string.Empty;
            // 讀取設定檔中的功能開關
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            // 判斷功能開關是否為 ON
            if (Fun == "ON")
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm2 實例數量是否小於 1
                        if (SOPForm2.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true)
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);

                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP2.pdf");

                                // 建立 SOPForm2 視窗
                                SOPForm2 Form = new SOPForm2();
                                // 實例數加一
                                SOPForm2.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm2.SOPName = "SOP2.pdf";
                                // 設定視窗標題
                                Form.Text = filename;
                                // 顯示 SOPForm2 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm2 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP2");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            // 功能開關不是 ON
            else
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm2 實例數量是否小於 1
                        if (SOPForm2.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true)
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);

                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP2.pdf");
                                // 建立 SOPForm2 視窗
                                SOPForm2 Form = new SOPForm2();
                                // 實例數加一
                                SOPForm2.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm2.SOPName = "SOP2.pdf";
                                // 顯示 SOPForm2 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm2 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP2");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }

        /// <summary>
        /// Btn_Sop3_Click 事件處理函式。
        /// 處理 SOP3 按鈕點擊時的邏輯，根據設定檔開關與選取項目，
        /// 下載或開啟對應的 SOP 檔案，並依副檔名決定開啟方式。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Sop3_Click(object sender, EventArgs e)
        {
            // 宣告檔名變數
            string filename = string.Empty;
            // 宣告路徑變數
            string path = string.Empty;
            // 宣告資料夾前綴變數
            string folder_F = string.Empty;
            // 宣告工程編號變數
            string strrtn = string.Empty;
            // 讀取設定檔中的功能開關
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            // 判斷功能開關是否為 ON
            if (Fun == "ON")
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm3 實例數量是否小於 1
                        if (SOPForm3.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true)
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);

                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP3.pdf");

                                // 建立 SOPForm3 視窗
                                SOPForm3 Form = new SOPForm3();
                                // 實例數加一
                                SOPForm3.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm3.SOPName = "SOP3.pdf";
                                // 設定視窗標題
                                Form.Text = filename;
                                // 顯示 SOPForm3 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm3 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP3");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        /// <summary>
        /// Btn_Sop4_Click 事件處理函式。
        /// 處理 SOP4 按鈕點擊時的邏輯，根據設定檔開關與選取項目，
        /// 下載或開啟對應的 SOP 檔案，並依副檔名決定開啟方式。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Sop4_Click(object sender, EventArgs e)
        {
            // 宣告檔名變數
            string filename = string.Empty;
            // 宣告路徑變數
            string path = string.Empty;
            // 宣告資料夾前綴變數
            string folder_F = string.Empty;
            // 宣告工程編號變數
            string strrtn = string.Empty;
            // 讀取設定檔中的功能開關
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            // 判斷功能開關是否為 ON
            if (Fun == "ON")
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm4 實例數量是否小於 1
                        if (SOPForm4.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);
                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP4.pdf");
                                // 建立 SOPForm4 視窗
                                SOPForm4 Form = new SOPForm4();
                                // 實例數加一
                                SOPForm4.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm4.SOPName = "SOP4.pdf";
                                // 設定視窗標題
                                Form.Text = filename;
                                // 顯示 SOPForm4 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm4 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP4");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            // 功能開關不是 ON
            else
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm4 實例數量是否小於 1
                        if (SOPForm4.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);
                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP4.pdf");
                                // 建立 SOPForm4 視窗
                                SOPForm4 Form = new SOPForm4();
                                // 實例數加一
                                SOPForm4.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm4.SOPName = "SOP4.pdf";
                                // 顯示 SOPForm4 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm4 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP4");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        /// <summary>
        /// Btn_Sop5_Click 事件處理函式。
        /// 處理 SOP5 按鈕點擊時的邏輯，根據設定檔開關與選取項目，
        /// 下載或開啟對應的 SOP 檔案，並依副檔名決定開啟方式。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Sop5_Click(object sender, EventArgs e)
        {
            // 宣告檔名變數
            string filename = string.Empty;
            // 宣告路徑變數
            string path = string.Empty;
            // 宣告資料夾前綴變數
            string folder_F = string.Empty;
            // 宣告工程編號變數
            string strrtn = string.Empty;
            // 讀取設定檔中的功能開關
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            // 判斷功能開關是否為 ON
            if (Fun == "ON")
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm5 實例數量是否小於 1
                        if (SOPForm5.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);
                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP5.pdf");
                                // 建立 SOPForm5 視窗
                                SOPForm5 Form = new SOPForm5();
                                // 實例數加一
                                SOPForm5.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm5.SOPName = "SOP5.pdf";
                                // 設定視窗標題
                                Form.Text = filename;
                                // 顯示 SOPForm5 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                // 使用共用 OpenFile 方法來開啟檔案並統一處理時間更新
                                OpenFile(path + folder_F + @"\" + strrtn + @"\" + filename);
                            }
                        }
                        // SOPForm5 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP5");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
            // 功能開關不是 ON
            else
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm5 實例數量是否小於 1
                        if (SOPForm5.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);
                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP5.pdf");
                                // 建立 SOPForm5 視窗
                                SOPForm5 Form = new SOPForm5();
                                // 實例數加一
                                SOPForm5.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm5.SOPName = "SOP5.pdf";
                                // 顯示 SOPForm5 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        // SOPForm5 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP5");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        /// <summary>
        /// Btn_Sop6_Click 事件處理函式。
        /// 處理 SOP6 按鈕點擊時的邏輯，根據設定檔開關與選取項目，
        /// 下載或開啟對應的 SOP 檔案，並依副檔名決定開啟方式。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Sop6_Click(object sender, EventArgs e)
        {
            // 宣告檔名變數
            string filename = string.Empty;
            // 宣告路徑變數
            string path = string.Empty;
            // 宣告資料夾前綴變數
            string folder_F = string.Empty;
            // 宣告工程編號變數
            string strrtn = string.Empty;
            // 讀取設定檔中的功能開關
            Fun = ini.IniReadValue("CHANGE_FILE", "FOUNCTION", "Setup.ini");
            // 判斷功能開關是否為 ON
            if (Fun == "ON")
            {
                // 檢查 listBox1 有項目且有選取項目
                if ((listBox1.Items.Count != 0) && (listBox1.SelectedItem != null))
                {
                    try
                    {
                        // 建立 WebClient 物件
                        WebClient ESOPclient = new WebClient();
                        // 取得選取項目文字
                        string special = listBox1.SelectedItem.ToString();
                        // 檢查 SOPForm6 實例數量是否小於 1
                        if (SOPForm6.m_Static_InstanceCount < 1)
                        {
                            // 判斷是否為生產通知單且不是特殊文件
                            if (Production.btn_producN == "btn_pro_sender" && special.IndexOf("EM") < 0 && special.IndexOf("SERVICE") < 0)
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為生產通知單路徑
                                path = Production.producnotice;
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                                // 取工程編號前兩碼
                                folder_F = strrtn.Substring(0, 2);
                            }
                            // 無工程編號搜尋模式
                            else if (Production.btn_producN == "noEngSearch")
                            {
                                // 直接取檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex].ToString();
                            }
                            // 特殊文件（EM 或 SERVICE）
                            else if (special.Substring(0, 2) == "EM" || special.Contains("SERVICE") == true) //包含memo.service file
                            {
                                // 取特殊檔案路徑
                                path = Production.filepathlist[listBox1.SelectedIndex - 1].ToString();
                            }
                            // 一般 SOP 文件
                            else
                            {
                                // 設定檔名
                                filename = listBox1.SelectedItem.ToString();
                                // 設定路徑為 SOP 路徑
                                path = Production.sopfile;
                                // 取檔名前兩碼
                                folder_F = filename.Substring(0, 2);
                                // 取得工程編號
                                strrtn = (txt_EngSr.Text == "") ? QueryDataSet(txt_WO.Text) : txt_EngSr.Text;
                            }
                            // 組合 SOP 檔案路徑
                            string SOPpath = (special.Substring(0, 2) == "EM") ? path : path + folder_F + @"\" + strrtn + @"\" + filename;
                            // 取得副檔名
                            String fileExtension = Path.GetExtension(SOPpath);
                            // 如果是 PDF 檔案
                            if (fileExtension == ".pdf")
                            {
                                // 建立 URI 物件
                                Uri uri = new Uri(SOPpath);
                                // 下載 PDF 檔案到 Temp 資料夾
                                ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP6.pdf");
                                // 建立 SOPForm6 視窗
                                SOPForm6 Form = new SOPForm6();
                                // 實例數加一
                                SOPForm6.m_Static_InstanceCount++;
                                // 設定 SOP 名稱
                                SOPForm6.SOPName = "SOP6.pdf";
                                // 顯示 SOPForm6 視窗
                                Form.Show();
                            }
                            // 如果不是 PDF 檔案
                            else
                            {
                                // 建立新程序開啟檔案
                                System.Diagnostics.Process pro = new System.Diagnostics.Process();
                                pro.StartInfo.FileName = path + folder_F + @"\" + strrtn + @"\" + filename;
                                pro.Start();
                                // 手動更新最後訪問時間
                                File.SetLastAccessTime(path + folder_F + @"\" + strrtn + @"\" + filename, DateTime.Now);
                            }
                        }
                        // SOPForm6 已開啟，提示使用者先關閉
                        else
                        {
                            MessageBox.Show("請先關閉SOP6");
                        }
                    }
                    // 捕捉例外錯誤
                    catch (Exception ex)
                    {
                        // 清空 listBox2
                        listBox2.Items.Clear();
                        // 新增錯誤訊息到 listBox2
                        addListItem(ex.Message);
                        // 顯示錯誤訊息
                        MessageBox.Show("關閉開啟中的SOP,完成後請重新操作!!");
                    }
                }
                // 沒有查詢或沒有選取項目
                else
                {
                    // 清空 listBox2
                    listBox2.Items.Clear();
                    // 新增提示訊息到 listBox2
                    addListItem("尚未查詢,無法下載!!");
                    // 顯示提示訊息
                    MessageBox.Show("尚未查詢,無法下載!!");
                }
            }
        }
        #endregion

        /// <summary>
        /// Btn_Clear_Click 事件處理函式。
        /// 按下「清除」按鈕時，會清空所有顯示資料與輸入欄位。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Clear_Click(object sender, EventArgs e)
        {
            All_Clear(); // 清空所有顯示與資料
            txt_EngSr.Text = ""; // 清空工程編號欄位
            txt_WO.Text = ""; // 清空工單號碼欄位
            //txt_Iso.Text = ""; // (註解) 清空ISO欄位
        }

        /// <summary>
        /// 處理「離開」按鈕點擊事件，關閉目前的主視窗。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Exit_Click(object sender, EventArgs e)
        {
            // 關閉目前視窗
            Close();
        }

        /// <summary>
        /// 判斷指定名稱的 Mutex 是否已存在（用於防止程式重複執行）。
        /// </summary>
        /// <param name="prgname">要檢查的 Mutex 名稱（通常為程式名稱）。</param>
        /// <returns>
        /// 若已存在同名 Mutex（即程式已在執行中）則回傳 true，否則回傳 false。
        /// </returns>
        private bool IsMyMutex(string prgname)
        {
            bool IsExist; // 宣告布林變數用來判斷 Mutex 是否已存在
            m = new Mutex(true, prgname, out IsExist); // 建立新的 Mutex，並取得是否已存在的狀態
            GC.Collect(); // 強制執行垃圾回收，釋放資源
            if (IsExist) // 如果 Mutex 是新建立的（表示沒有重複執行）
            {
                return false; // 回傳 false，表示沒有重複執行
            }
            else // 如果 Mutex 已存在（表示程式已經在執行中）
            {
                return true; // 回傳 true，表示有重複執行
            }
        }

        /// <summary>
        /// 清空
        /// </summary>
        private void All_Clear()
        {
            listBox1.Items.Clear(); // 清空 listBox1 的所有項目
            listBox2.Items.Clear(); // 清空 listBox2 的所有項目

            txt_Ver1.Text = ""; // 清空 txt_Ver1 的文字內容
            txt_Model.Text = ""; // 清空 txt_Model 的文字內容
            label1.Text = ""; // 清空 label1 的文字內容
            label2.Text = ""; // 清空 label2 的文字內容
            Production.btn_producN = ""; // 將 Production.btn_producN 設為空字串
        }
        /// <summary>
        /// 將指定的訊息加入 listBox2，並將文字顏色設為紅色。
        /// </summary>
        /// <param name="value">要顯示於 listBox2 的訊息內容。</param>
        private void addListItem(string value)
        {
            // 設定 listBox2 的文字顏色為紅色
            this.listBox2.ForeColor = Color.Red;
            // 將傳入的 value 加入 listBox2 的項目中
            this.listBox2.Items.Add(value);
        }

        #region 固定刪除過期檔案的背景計時器
        /// <summary>
        /// 啟動一個背景計時器，週期性呼叫 DeleteOldFiles()。
        /// 會先立即在背景執行一次 DeleteOldFiles()，之後以 _deleteOldFilesInterval 週期執行。
        /// </summary>
        private void StartDeleteOldFilesTimer()
        {
            try
            {
                // 先立刻在背景執行一次（避免阻塞 UI）
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try { DeleteOldFiles(); } catch (Exception ex) { Console.WriteLine("DeleteOldFiles initial run error: " + ex.Message); }
                });

                // 建立 System.Threading.Timer，state 為 null
                _deleteOldFilesTimer = new System.Threading.Timer(_ =>
                {
                    try { DeleteOldFiles(); }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Periodic DeleteOldFiles error: " + ex.Message);
                    }
                }, null, _deleteOldFilesInterval, _deleteOldFilesInterval);
            }
            catch (Exception ex)
            {
                Console.WriteLine("StartDeleteOldFilesTimer error: " + ex.Message);
            }
        }

        /// <summary>
        /// 停止並釋放背景計時器
        /// </summary>
        private void StopDeleteOldFilesTimer()
        {
            try
            {
                if (_deleteOldFilesTimer != null)
                {
                    _deleteOldFilesTimer.Change(Timeout.Infinite, Timeout.Infinite);
                    _deleteOldFilesTimer.Dispose();
                    _deleteOldFilesTimer = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("StopDeleteOldFilesTimer error: " + ex.Message);
            }
        }
        #endregion

        #region 表單關閉時，停止並釋放背景計時器
        /// <summary>
        /// 表單關閉時，停止並釋放背景計時器以避免背景執行緒繼續執行。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 停止定時任務
            StopDeleteOldFilesTimer();

            base.OnFormClosing(e);
        }
        #endregion

        #region 刪除資料夾中的過期檔案
        /// <summary>
        /// 刪除過期檔案
        /// </summary>
        private void DeleteOldFiles()
        {
            // 設定過期的時間（例如：檔案超過一天未使用則刪除）
            TimeSpan fileExpirationDuration = TimeSpan.FromDays(1);

            string folderPath = Path.Combine(Application.StartupPath, "Temp");

            // 刪除資料夾中的過期檔案
            DeleteFilesInFolder(folderPath, fileExpirationDuration);
        }

        /// <summary>
        /// 刪除資料夾中的過期檔案
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="fileExpirationDuration"></param>
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

        /// <summary>
        /// 檢查檔案是否正在使用
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 刪除檔案
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="timesToWrite"></param>
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

        #endregion

        #region 版本自動更新
        #region .ini 讀寫功能
        /// <summary>
        /// 提供 INI 檔案的讀寫功能，使用 Windows API 操作設定檔。
        /// </summary>
        public class SetupIniIP
        {
            /// <summary>
            /// INI 檔案路徑。
            /// </summary>
            public string path;

            /// <summary>
            /// 透過 Windows API 寫入 INI 檔案指定區段與鍵值。
            /// </summary>
            /// <param name="section">區段名稱</param>
            /// <param name="key">鍵名稱</param>
            /// <param name="val">要寫入的值</param>
            /// <param name="filePath">INI 檔案路徑</param>
            /// <returns>API 回傳結果</returns>
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern long WritePrivateProfileString(string section,
                string key, string val, string filePath);

            /// <summary>
            /// 透過 Windows API 讀取 INI 檔案指定區段與鍵值。
            /// </summary>
            /// <param name="section">區段名稱</param>
            /// <param name="key">鍵名稱</param>
            /// <param name="def">預設值</param>
            /// <param name="retVal">回傳值緩衝區</param>
            /// <param name="size">緩衝區大小</param>
            /// <param name="filePath">INI 檔案路徑</param>
            /// <returns>API 回傳結果</returns>
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern int GetPrivateProfileString(string section,
                string key, string def, StringBuilder retVal,
                int size, string filePath);

            /// <summary>
            /// 寫入 INI 檔案指定區段與鍵值。
            /// </summary>
            /// <param name="Section">區段名稱</param>
            /// <param name="Key">鍵名稱</param>
            /// <param name="Value">要寫入的值</param>
            /// <param name="inipath">INI 檔案名稱</param>
            public void IniWriteValue(string Section, string Key, string Value, string inipath)
            {
                // 寫入指定區段與鍵值到 INI 檔案
                WritePrivateProfileString(Section, Key, Value, Application.StartupPath + "\\" + inipath);
            }

            /// <summary>
            /// 讀取 INI 檔案指定區段與鍵值。
            /// </summary>
            /// <param name="Section">區段名稱</param>
            /// <param name="Key">鍵名稱</param>
            /// <param name="inipath">INI 檔案名稱</param>
            /// <returns>讀取到的值</returns>
            public string IniReadValue(string Section, string Key, string inipath)
            {
                // 建立緩衝區儲存讀取結果
                StringBuilder temp = new StringBuilder(255);
                // 讀取指定區段與鍵值
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath);
                // 回傳讀取到的值
                return temp.ToString();
            }
        }
        #endregion
        #region 取得新版本資訊
        /// <summary>
        /// 取得新版本資訊。
        /// 從 TE_Program_Table 依據指定程式名稱查詢最新版本號碼。
        /// </summary>
        /// <param name="tool">程式名稱（如 "E-SOP"）。</param>
        /// <returns>查詢到的新版本號碼，若查無則回傳原本的 version_new。</returns>
        private string selectVerSQL_new(string tool)
        {
            string sqlCmd = ""; // 宣告 SQL 指令字串

            try
            {
                // 組合 SQL 查詢語句，查詢指定程式名稱的所有欄位
                sqlCmd = "select *  FROM TE_Program_Table where [Program_Name] ='" + tool + "'";
                // 執行 SQL 查詢，取得資料集
                DataSet ds = db.reDs(sqlCmd);
                // 如果查詢結果有資料
                if (ds.Tables[0].Rows.Count != 0)
                {
                    // 逐一檢查欄位名稱
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        // 如果欄位名稱為 "Version"
                        if (ds.Tables[0].Columns[i].ToString() == "Version")
                        {
                            // 取得新版本號碼
                            version_new = ds.Tables[0].Rows[0][i].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 發生例外時在主控台顯示錯誤訊息
                Console.WriteLine($"更新異常時發生錯誤：{ex.Message}");
                // 顯示錯誤訊息視窗
                MessageBox.Show("更新異常");
            }
            // 回傳新版本號碼
            return version_new;
        }
        #endregion

        /// <summary>
        /// Autoupdates this instance.
        /// </summary>
        public void autoupdate()//自動更新
        {
            // 建立一個新的 Process 物件
            Process p = new Process();
            // 設定要執行的檔案名稱為 AutoUpdate.exe，路徑為應用程式啟動目錄
            p.StartInfo.FileName = System.Windows.Forms.Application.StartupPath + "\\AutoUpdate.exe";
            // 設定工作目錄為應用程式啟動目錄
            p.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath; //檔案所在的目錄
            // 啟動自動更新程式
            p.Start();
            // 關閉目前主程式
            this.Close();
        }
        #endregion

        #region 無使用
        //public static string StrRight(string param, int length)
        //{
        //    string result = param.Substring(param.Length - length, length);
        //    return result;
        //}

        /// <summary>
        /// 下載 PDF 並開啟對應 SOPForm 視窗，依照 SopQty 決定視窗與檔名。
        /// </summary>
        /// <param name="SopQty">SOP 視窗編號（1~6，對應 SOPForm1~SOPForm6）</param>
        /// <param name="SopWeb">PDF 檔案在 Web 上的相對路徑</param>
        public void doPdf(int SopQty, string SopWeb)
        {
            WebClient ESOPclient = new WebClient(); // 建立 WebClient 物件

            if (SopQty == 1) // 若為 SOP1
            {
                if (SOPForm1.m_Static_InstanceCount < 1) // 檢查 SOPForm1 是否已開啟
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb); // 組合下載網址
                    string Filename = SopWeb.Replace("documents_maintain_files/", ""); // 取得檔名（未使用）
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP1.pdf"); // 下載 PDF 檔案到指定路徑

                    SOPForm1 Form = new SOPForm1(); // 建立 SOPForm1 視窗
                    SOPForm1.SOPName = "SOP1.pdf"; // 設定 SOP 名稱
                    SOPForm1.m_Static_InstanceCount++; // 實例數加一
                    Form.Show(); // 顯示 SOPForm1 視窗
                }
                else
                {
                    MessageBox.Show("請先關閉SOP1"); // 已開啟則提示
                }
            }
            else if (SopQty == 2) // 若為 SOP2
            {
                if (SOPForm2.m_Static_InstanceCount < 1) // 檢查 SOPForm2 是否已開啟
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb); // 組合下載網址
                    string Filename = SopWeb.Replace("documents_maintain_files/", ""); // 取得檔名（未使用）
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP2.pdf"); // 下載 PDF 檔案到指定路徑

                    SOPForm2 Form = new SOPForm2(); // 建立 SOPForm2 視窗
                    SOPForm2.SOPName = "SOP2.pdf"; // 設定 SOP 名稱
                    SOPForm2.m_Static_InstanceCount++; // 實例數加一
                    Form.Show(); // 顯示 SOPForm2 視窗
                }
                else
                {
                    MessageBox.Show("請先關閉SOP2"); // 已開啟則提示
                }
            }
            else if (SopQty == 3) // 若為 SOP3
            {
                if (SOPForm3.m_Static_InstanceCount < 1) // 檢查 SOPForm3 是否已開啟
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb); // 組合下載網址
                    string Filename = SopWeb.Replace("documents_maintain_files/", ""); // 取得檔名（未使用）
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP3.pdf"); // 下載 PDF 檔案到指定路徑

                    SOPForm3 Form = new SOPForm3(); // 建立 SOPForm3 視窗
                    SOPForm3.SOPName = "SOP3.pdf"; // 設定 SOP 名稱
                    SOPForm3.m_Static_InstanceCount++; // 實例數加一
                    Form.Show(); // 顯示 SOPForm3 視窗
                }
                else
                {
                    MessageBox.Show("請先關閉SOP3"); // 已開啟則提示
                }
            }
            else if (SopQty == 4) // 若為 SOP4
            {
                if (SOPForm4.m_Static_InstanceCount < 1) // 檢查 SOPForm4 是否已開啟
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb); // 組合下載網址
                    string Filename = SopWeb.Replace("documents_maintain_files/", ""); // 取得檔名（未使用）
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP4.pdf"); // 下載 PDF 檔案到指定路徑

                    SOPForm4 Form = new SOPForm4(); // 建立 SOPForm4 視窗
                    SOPForm4.SOPName = "SOP4.pdf"; // 設定 SOP 名稱
                    SOPForm4.m_Static_InstanceCount++; // 實例數加一
                    Form.Show(); // 顯示 SOPForm4 視窗
                }
                else
                {
                    MessageBox.Show("請先關閉SOP4"); // 已開啟則提示
                }
            }
            else if (SopQty == 5) // 若為 SOP5
            {
                if (SOPForm5.m_Static_InstanceCount < 1) // 檢查 SOPForm5 是否已開啟
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb); // 組合下載網址
                    string Filename = SopWeb.Replace("documents_maintain_files/", ""); // 取得檔名（未使用）
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP5.pdf"); // 下載 PDF 檔案到指定路徑

                    SOPForm5 Form = new SOPForm5(); // 建立 SOPForm5 視窗
                    SOPForm5.SOPName = "SOP5.pdf"; // 設定 SOP 名稱
                    SOPForm5.m_Static_InstanceCount++; // 實例數加一
                    Form.Show(); // 顯示 SOPForm5 視窗
                }
                else
                {
                    MessageBox.Show("請先關閉SOP5"); // 已開啟則提示
                }
            }
            else if (SopQty == 6) // 若為 SOP6
            {
                if (SOPForm6.m_Static_InstanceCount < 1) // 檢查 SOPForm6 是否已開啟
                {
                    Uri uri = new Uri("http://qazone.avalue.com.tw/isozone/" + SopWeb); // 組合下載網址
                    string Filename = SopWeb.Replace("documents_maintain_files/", ""); // 取得檔名（未使用）
                    ESOPclient.DownloadFile(uri, Application.StartupPath + "/SOP6.pdf"); // 下載 PDF 檔案到指定路徑

                    SOPForm6 Form = new SOPForm6(); // 建立 SOPForm6 視窗
                    SOPForm6.SOPName = "SOP6.pdf"; // 設定 SOP 名稱
                    SOPForm6.m_Static_InstanceCount++; // 實例數加一
                    Form.Show(); // 顯示 SOPForm6 視窗
                }
                else
                {
                    MessageBox.Show("請先關閉SOP6"); // 已開啟則提示
                }
            }
        }

        /// <summary>
        /// SFIS 資料結構，對應 SFIS 系統回傳的各種 SOP 與描述欄位。
        /// </summary>
        public class SFIS
        {
            /// <summary>
            /// 機種編號
            /// </summary>
            public string productid { get; set; }

            /// <summary>
            /// 製造 SOP 檔案路徑
            /// </summary>
            public string mfc_sop { get; set; }

            /// <summary>
            /// 製造 SOP 說明
            /// </summary>
            public string mfc_sop_desc { get; set; }

            /// <summary>
            /// 測試 SOP 檔案路徑
            /// </summary>
            public string testing_sop { get; set; }

            /// <summary>
            /// 測試 SOP 說明
            /// </summary>
            public string testing_sop_desc { get; set; }

            /// <summary>
            /// 包裝 SOP 檔案路徑
            /// </summary>
            public string packing_sop { get; set; }

            /// <summary>
            /// 包裝 SOP 說明
            /// </summary>
            public string packing_sop_desc { get; set; }

            /// <summary>
            /// 組裝 SOP 檔案路徑
            /// </summary>
            public string assy_sop { get; set; }

            /// <summary>
            /// 組裝 SOP 說明
            /// </summary>
            public string assy_sop_desc { get; set; }

            /// <summary>
            /// 特殊包裝 SOP 檔案路徑
            /// </summary>
            public string spe_packing_sop { get; set; }

            /// <summary>
            /// 特殊包裝 SOP 說明
            /// </summary>
            public string spe_packing_sop_desc { get; set; }

            /// <summary>
            /// 其他 SOP 檔案路徑
            /// </summary>
            public string other_sop { get; set; }

            /// <summary>
            /// 其他 SOP 說明
            /// </summary>
            public string other_sop_desc { get; set; }

            /// <summary>
            /// PE 額外注意事項
            /// </summary>
            public string pe_extnotes { get; set; }
        }

        /// <summary>
        /// 依據 SOP 名稱查詢 ISOZONE 資料庫，回傳檔案路徑。
        /// </summary>
        /// <param name="SOPNAME">SOP 文件名稱。</param>
        /// <returns>查詢到的 SOP 檔案路徑，若查無資料則回傳空字串。</returns>
        private string ISOZONE_SOPNAMESERCH(string SOPNAME)
        {
            string SopPath = ""; // 儲存查詢到的 SOP 路徑
            try
            {
                // 組合 SQL 查詢語句，根據文件名稱查詢
                string sqlCmd = "SELECT * FROM [isozone_db].[dbo].[doc_esop_view] WHERE [document_name] = N'" + SOPNAME + "'  ";
                // 執行查詢，取得資料集
                DataSet ds = db_IsoZoen.reDs(sqlCmd);
                // 如果查詢有結果
                if (ds.Tables[0].Rows.Count > 0)
                {
                    // 取得檔案路徑並去除前後空白
                    SopPath = ds.Tables[0].Rows[0]["file_path"].ToString().Trim();
                    // 顯示文件版本於 txt_Ver1
                    txt_Ver1.Text = ds.Tables[0].Rows[0]["document_version"].ToString().Trim();
                    // 回傳檔案路徑
                    return SopPath;
                }
                else
                {
                    // 查無資料則回傳空字串
                    return SopPath;
                }
            }
            catch
            {
                // 發生例外時回傳空字串
                return SopPath;
            }
        }

        /// <summary>
        /// 取得指定 FTP 伺服器的連線資訊，並將結果存入 ftpServer、ftpuser、ftppassword 欄位。
        /// </summary>
        /// <param name="Ftp_Server_name">FTP 伺服器名稱（資料庫查詢條件）。</param>
        private void Getftp(string Ftp_Server_name)
        {
            // 建立 SQL 查詢語句，根據 FTP 伺服器名稱查詢相關資訊
            string sqlCmd = "SELECT [Ftp_Server_OA_Ip],[Ftp_Username],[Ftp_Password],[Ftp_Server_name] FROM i_Program_FtpServer_Table where [Ftp_Server_name] ='" + Ftp_Server_name + "' ";
            // 執行 SQL 查詢，取得資料集
            DataSet ds = db.reDs(sqlCmd);
            // 如果查詢結果有資料
            if (ds.Tables[0].Rows.Count != 0)
            {
                // 逐筆處理查詢結果（雖然只取第一筆資料）
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    // 取得 FTP 伺服器 IP 並去除前後空白
                    ftpServer = ds.Tables[0].Rows[0]["Ftp_Server_OA_Ip"].ToString().Trim();
                    // 取得 FTP 使用者名稱並去除前後空白
                    ftpuser = ds.Tables[0].Rows[0]["Ftp_Username"].ToString().Trim();
                    // 取得 FTP 密碼並去除前後空白
                    ftppassword = ds.Tables[0].Rows[0]["Ftp_Password"].ToString().Trim();
                }
            }
        }
        #endregion
    }
}
