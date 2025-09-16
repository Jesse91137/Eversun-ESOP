using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;
using System.Data;


namespace E_SOP
{
    /// <summary>
    /// WebFTP 類別，提供 FTP 檔案上傳、下載及檔案大小查詢功能。
    /// </summary>
    public class WebFTP
    {
        /// <summary>
        /// FTP 伺服器位址。
        /// </summary>
        public static string FTP_Server;

        /// <summary>
        /// FTP 登入使用者名稱。
        /// </summary>
        public static string FTP_User;
        /// <summary>
        /// FTP 登入密碼。
        /// </summary>
        public static string FTP_Password;


        /// <summary>
        /// 使用 Windows API 進行使用者登入驗證。
        /// </summary>
        /// <param name="lpszUsername">登入使用者名稱。</param>
        /// <param name="lpszDomain">登入網域名稱。</param>
        /// <param name="lpszPassword">登入密碼。</param>
        /// <param name="dwLogonType">登入型態。</param>
        /// <param name="dwLogonProvider">登入提供者。</param>
        /// <param name="phToken">回傳的權杖指標。</param>
        /// <returns>登入成功回傳 true，否則回傳 false。</returns>
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        /// <summary>
        /// 緩衝區大小常數，單位為位元組 (byte)。
        /// </summary>
        const int BUFFER_SIZE = 4096;

        /// <summary>
        /// 登出，釋放指定的權杖資源。
        /// </summary>
        /// <param name="hToken">要關閉的權杖指標。</param>
        /// <returns>成功釋放則回傳 true，否則回傳 false。</returns>
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);

        /// <summary>
        /// 取得 FTP 伺服器上指定檔案的大小（位元組）。
        /// </summary>
        /// <param name="filename">欲查詢大小的檔案名稱。</param>
        /// <param name="Path">檔案所在路徑（未使用）。</param>
        /// <param name="NASserver_master">FTP 伺服器主機名稱（未使用）。</param>
        /// <returns>檔案大小（位元組），若失敗則回傳 0。</returns>
        /// <remarks>
        /// 此方法會連線至 FTP 伺服器，並取得指定檔案的大小。若發生例外，會於主控台顯示錯誤訊息。
        /// </remarks>
        public static long GetFileSize(string filename, string Path, string NASserver_master)
        {
            FtpWebRequest reqFTP; // 宣告 FTP 請求物件
            long fileSize = 0; // 檔案大小初始值
            try
            {
                // 建立 FTP 請求，指定 FTP 伺服器位址與檔案路徑
                reqFTP = (FtpWebRequest)FtpWebRequest.Create("ftp://" + "172.22.2.82" + "/" + "ESOP/PDF/" + filename);
                reqFTP.Method = WebRequestMethods.Ftp.GetFileSize; // 設定請求方法為取得檔案大小
                reqFTP.UseBinary = true; // 設定使用二進位模式
                reqFTP.KeepAlive = true; // 設定保持連線
                reqFTP.Credentials = new NetworkCredential(FTP_User, FTP_Password); // 設定 FTP 登入帳號密碼
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse(); // 取得 FTP 回應
                Stream ftpStream = response.GetResponseStream(); // 取得回應串流
                fileSize = response.ContentLength; // 取得檔案大小（位元組）
                ftpStream.Close(); // 關閉串流
                response.Close(); // 關閉回應
            }
            catch (Exception ex)
            {
                // 發生例外時，於主控台顯示錯誤訊息
                Console.WriteLine($"取的檔案大小時發生錯誤：{ex.Message}");
            }
            return fileSize; // 回傳檔案大小
        }
        /// <summary>
        /// 檔案上傳
        /// </summary>
        /// <param name="updatefilename">本機檔案完整路徑（欲上傳的檔案）。</param>
        /// <param name="filename">上傳至 FTP 的檔案名稱。</param>
        /// <param name="file">FTP 目標資料夾名稱。</param>
        /// <param name="NASserver_master">FTP 伺服器主機名稱（用於查詢 FTP 連線資訊）。</param>
        /// <param name="Bar1">進度條控制元件，顯示上傳進度。</param>
        /// <returns>上傳成功回傳 true，失敗回傳 false。</returns>
        public static bool UploadFile(string updatefilename, string filename, string file, string NASserver_master, ProgressBar Bar1)
        {
            bool result = true; // 上傳結果

            try
            {
                Getftp(NASserver_master); // 取得 FTP 連線資訊

                FileInfo finfo = new FileInfo(updatefilename); // 取得檔案資訊

                FtpWebResponse response = null; // FTP 回應物件
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + FTP_Server + "/" + file + "/" + filename); // 建立 FTP 上傳請求
                request.KeepAlive = true; // 保持連線
                request.UseBinary = true; // 使用二進位模式
                request.Credentials = new NetworkCredential(FTP_User, FTP_Password); // 設定 FTP 帳號密碼
                request.Method = WebRequestMethods.Ftp.UploadFile; // 設定上傳方法
                request.ContentLength = finfo.Length; // 指定上傳檔案大小
                response = request.GetResponse() as FtpWebResponse; // 取得 FTP 回應
                int buffLength = 2048; // 緩衝區大小
                byte[] buffer = new byte[buffLength]; // 緩衝區
                int contentLen; // 讀取長度
                FileStream fs = File.OpenRead(updatefilename); // 開啟本機檔案串流
                Stream ftpstream = request.GetRequestStream(); // 取得 FTP 串流
                contentLen = fs.Read(buffer, 0, buffer.Length); // 讀取檔案內容
                int allbye = (int)finfo.Length; // 檔案總大小
                Form.CheckForIllegalCrossThreadCalls = false; // 關閉跨執行緒檢查
                Bar1.Maximum = allbye; // 設定進度條最大值
                int startbye = 0; // 已上傳位元組數
                while (contentLen != 0) // 持續上傳直到檔案結束
                {
                    startbye = contentLen + startbye; // 累加已上傳位元組
                    ftpstream.Write(buffer, 0, contentLen); // 寫入 FTP 串流
                    // 更新進度
                    if (Bar1 != null)
                    {
                        Bar1.Value += contentLen; // 更新進度條
                    }
                    contentLen = fs.Read(buffer, 0, buffLength); // 讀取下一段檔案內容
                }
                fs.Close(); // 關閉本機檔案串流
                ftpstream.Close(); // 關閉 FTP 串流
                response.Close(); // 關閉 FTP 回應
            }
            catch (Exception ftp)
            {
                result = false; // 上傳失敗
                MessageBox.Show(ftp.Message); // 顯示錯誤訊息
            }

            return result; // 回傳結果
        }
        /// <summary>
        /// 下載檔案至本機指定路徑。
        /// </summary>
        /// <param name="DLfilename">FTP 來源檔案名稱。</param>
        /// <param name="LocalPath">本機儲存路徑。</param>
        /// <param name="NASserver_master">FTP 伺服器主機名稱（用於查詢 FTP 連線資訊）。</param>
        /// <param name="FilePath">FTP 目標資料夾路徑。</param>
        /// <returns>下載成功回傳 true，失敗回傳 false。</returns>
        public static bool Downloadfile(string DLfilename, string LocalPath, string NASserver_master, string FilePath)
        {
            bool result = true; // 下載結果

            try
            {
                Getftp(NASserver_master); // 取得 FTP 連線資訊
                string tempStoragePath = LocalPath; // 設定本機暫存路徑

                // 建立 FTP 下載請求
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create("ftp://" + FTP_Server + "/" + FilePath + "/" + DLfilename);
                NetworkCredential ftpCredential = new NetworkCredential(FTP_User, FTP_Password); // 設定 FTP 登入帳號密碼
                ftpRequest.Credentials = ftpCredential; // 指定憑證
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile; // 設定請求方法為下載檔案

                // 取得 FTP 回應
                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                // 取得回應串流
                Stream ftpStream = ftpResponse.GetResponseStream();
                using (FileStream fileStream = new FileStream(tempStoragePath, FileMode.Create))
                {
                    int bufferSize = 2048; // 緩衝區大小
                    int readCount; // 讀取長度
                    byte[] buffer = new byte[bufferSize]; // 緩衝區

                    readCount = ftpStream.Read(buffer, 0, bufferSize); // 讀取 FTP 串流內容
                    int allbye = (int)fileStream.Length; // 檔案總大小
                    Form.CheckForIllegalCrossThreadCalls = false; // 關閉跨執行緒檢查

                    while (readCount > 0) // 持續寫入直到檔案結束
                    {
                        fileStream.Write(buffer, 0, readCount); // 寫入本機檔案
                        readCount = ftpStream.Read(buffer, 0, bufferSize); // 讀取下一段內容
                    }
                }
                ftpStream.Close(); // 關閉 FTP 串流
                ftpResponse.Close(); // 關閉 FTP 回應

                return result; // 回傳結果
            }
            catch (Exception ex)
            {
                Console.WriteLine($"下載時發生錯誤：{ex.Message}"); // 顯示錯誤訊息
                result = false; // 下載失敗
                return result; // 回傳結果
            }
        }

        /// <summary>
        /// 以 Windows API 進行身分模擬並複製檔案到指定機器的共享資料夾。
        /// </summary>
        /// <param name="updatefilename">本機檔案完整路徑（欲上傳的檔案）。</param>
        /// <param name="filename">來源檔案名稱。</param>
        /// <param name="file">目標資料夾名稱。</param>
        /// <param name="NASserver_master">目標機器名稱（用於 Windows 驗證）。</param>
        /// <param name="Bar1">進度條控制元件（未使用）。</param>
        /// <returns>複製成功回傳 true，失敗回傳 false。</returns>
        public static bool SAMbar(string updatefilename, string filename, string file, string NASserver_master, ProgressBar Bar1)
        {
            bool result = true; // 複製結果
            try
            {
                string MachineName = NASserver_master; // 取得目標機器名稱
                string UserName = FTP_User; // 取得登入使用者名稱
                string Pw = FTP_Password; // 取得登入密碼
                string IPath = String.Format(@"\\{0}\abc", MachineName); // 建立網路路徑（未使用）
                const int LOGON32_PROVIDER_DEFAULT = 0; // 預設登入提供者
                const int LOGON32_LOGON_NEW_CREDENTIALS = 9; // 使用新認證登入型態
                IntPtr tokenHandle = new IntPtr(0); // 權杖指標初始化
                tokenHandle = IntPtr.Zero; // 權杖指標歸零

                // 呼叫 Windows API 進行身分驗證
                bool returnValue = LogonUser(UserName, MachineName, Pw,
                LOGON32_LOGON_NEW_CREDENTIALS,
                LOGON32_PROVIDER_DEFAULT,
                ref tokenHandle);

                // 進行身分模擬
                WindowsIdentity w = new WindowsIdentity(tokenHandle);
                w.Impersonate();
                if (false == returnValue)
                {
                    result = false; // 驗證失敗
                }
                FileInfo finfo = new FileInfo(filename); // 取得檔案資訊
                updatefilename = Path.GetFileNameWithoutExtension(filename) + Path.GetExtension(filename); // 取得檔案名稱（含副檔名）
                DirectoryInfo dir = new DirectoryInfo(file); // 取得目標資料夾資訊

                // 複製檔案到網路共享資料夾
                File.Copy(filename, @"\\" + MachineName + "\\" + file + updatefilename, true);
                // 檢查檔案是否複製成功
                if (File.Exists(@"\\" + MachineName + "\\" + file + updatefilename) == true)
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"時發生錯誤：{ex.Message}"); // 發生例外時顯示錯誤訊息
                result = false; // 複製失敗
            }
            return result; // 回傳結果
        }

        /// <summary>
        /// 取得 FTP 伺服器連線資訊。
        /// </summary>
        /// <param name="Ftp_Server_name">FTP 伺服器名稱。</param>
        /// <remarks>
        /// 從資料庫查詢指定 FTP 伺服器名稱的連線資訊，並設定至靜態欄位 FTP_Server、FTP_User、FTP_Password。
        /// </remarks>
        public static void Getftp(string Ftp_Server_name)//ftp資訊
        {
            // 建立 SQL 查詢語句，取得 FTP 伺服器 IP、使用者名稱、密碼、伺服器名稱
            string sqlCmd = "SELECT [Ftp_Server_OA_Ip],[Ftp_Username],[Ftp_Password],[Ftp_Server_name] FROM [iFactory].[i_Program].[i_Program_FtpServer_Table] where [Ftp_Server_name] ='" + Ftp_Server_name + "' ";
            // 執行 SQL 查詢，取得結果 DataSet
            DataSet ds = E_SOP.db.reDs(sqlCmd);
            // 檢查查詢結果是否有資料
            if (ds.Tables[0].Rows.Count != 0)
            {
                // 逐筆設定 FTP 連線資訊（實際只取第一筆）
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    // 設定 FTP 伺服器 IP
                    FTP_Server = ds.Tables[0].Rows[0]["Ftp_Server_OA_Ip"].ToString().Trim();
                    // 設定 FTP 使用者名稱
                    FTP_User = ds.Tables[0].Rows[0]["Ftp_Username"].ToString().Trim();
                    // 設定 FTP 密碼
                    FTP_Password = ds.Tables[0].Rows[0]["Ftp_Password"].ToString().Trim();
                }
            }
        }
    }
}
