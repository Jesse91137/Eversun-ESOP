using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Text;

namespace E_SOP
{
    /// <summary>
    /// 共用檔案工具方法。
    /// </summary>
    public static class FileUtils
    {

        /// <summary>
        /// 使用預設 shell 開啟指定的檔案或 URL；若為本機檔案則嘗試更新最後存取與最後寫入時間（含 UTC）。
        /// </summary>
        /// <param name="filePath">要開啟的本機路徑或 URL。</param>
        /// <remarks>
        /// - 方法會先嘗試解析傳入字串是否為 file:// URI，並在檔案存在時先呼叫 TrySetFileTimes 嘗試更新時間戳記以降低外部程式先鎖定檔案後無法更新的風險。
        /// - 若第一次更新時間失敗，會在背景非阻塞地重試數次；不會因更新時間失敗而阻止開啟檔案的動作（fail-safe）。
        /// - 若啟動外部程式發生錯誤，會以 MessageBox 提示使用者，但不會拋出例外給呼叫端。
        /// </remarks>
        public static void OpenFile(string filePath) // 開放給外部呼叫，用於啟動檔案或 URL
        {
            if (string.IsNullOrWhiteSpace(filePath)) // 若傳入為 null、空字串或僅空白則
            { // 開始 if 區塊
                return; // 直接回傳，不進行任何操作
            } // 結束 if 區塊

            // 先解析是否為本機檔案路徑
            string localPath = null; // 宣告 localPath 變數以儲存可能的本機路徑
            if (Uri.TryCreate(filePath, UriKind.Absolute, out var uri) && uri.IsFile) // 嘗試將輸入解析為絕對 URI，且檢查是否為檔案 URI
            { // 開始 if 區塊
                localPath = uri.LocalPath; // 若是檔案 URI，使用其 LocalPath 作為本機路徑
            } // 結束 if 區塊
            else // 否則
            { // 開始 else 區塊
                localPath = filePath; // 直接將傳入的字串當作本機路徑使用
            } // 結束 else 區塊

            // 如果是本機檔案，先嘗試在啟動前更新時間戳記，這樣通常能避免在外部程式啟動後被鎖住而無法更新的問題
            bool initialTimeSet = false; // 用來記錄第一次嘗試是否成功
            if (!string.IsNullOrEmpty(localPath) && File.Exists(localPath)) // 若 localPath 非空且該路徑存在檔案
            { // 開始 if 區塊
                initialTimeSet = TrySetFileTimes(localPath, DateTime.Now); // 呼叫 TrySetFileTimes 嘗試以當前本地時間更新時間戳記
                if (!initialTimeSet) // 若第一次嘗試失敗
                { // 開始 if 區塊
                  // 若第一次失敗，非阻塞地在背景重試數次（可能是外部程式短暫鎖定檔案）
                    _ = Task.Run(async () => // 啟動背景任務，不等待其完成（fire-and-forget）
                    { // 開始匿名 async lambda 區塊
                        const int maxRetries = 5; // 定義最大重試次數
                        const int delayMs = 1000; // 每次重試間隔（毫秒）
                        for (int i = 0; i < maxRetries; i++) // 以迴圈進行多次重試
                        { // 開始 for 迴圈區塊
                            try // 嘗試區塊：保護 Task.Delay 與 TrySetFileTimes 呼叫，避免未處理例外中斷背景任務
                            { // 開始 try 區塊
                                await Task.Delay(delayMs).ConfigureAwait(false); // 等待指定的毫秒數（非同步延遲）
                                if (TrySetFileTimes(localPath, DateTime.Now)) // 再次嘗試更新時間戳記
                                { // 開始 if 區塊
                                    Console.WriteLine($"背景更新檔案時間成功: {localPath}"); // 若成功則輸出成功訊息到 Console
                                    return; // 成功後結束背景任務
                                } // 結束 if 區塊
                            } // 結束 try 區塊
                            catch (Exception bgEx) // 捕捉背景任務內可能拋出的任何例外
                            { // 開始 catch 區塊
                                Console.WriteLine($"背景更新檔案時間發生例外: {localPath} => {bgEx.Message}"); // 將例外訊息寫到 Console 以便偵錯
                            } // 結束 catch 區塊
                        } // 結束 for 迴圈區塊
                        Console.WriteLine($"多次嘗試後仍無法更新檔案時間: {localPath}"); // 若所有重試仍失敗則輸出最終失敗訊息
                    }); // 結束匿名 async lambda 與 Task.Run 呼叫
                } // 結束 if (!initialTimeSet) 區塊
            } // 結束 if (File.Exists) 區塊

            try // 使用 shell 啟動檔案或 URL 的保護區塊，避免例外外洩
            { // 開始 try 區塊
              // 啟動檔案或 URL（不會因為更新時間失敗而停止啟動）
                var psi = new ProcessStartInfo(filePath) // 建立 ProcessStartInfo 並指定要啟動的檔案或 URL
                { // 開始物件初始化器
                    UseShellExecute = true // 設定 UseShellExecute = true 使用系統預設 shell 來開啟資源
                }; // 結束物件初始化器
                Process.Start(psi); // 使用 Process.Start 啟動對應的應用程式或瀏覽器
            } // 結束 try 區塊
            catch (Exception ex) // 捕捉啟動外部程式時發生的例外
            { // 開始 catch 區塊
              // 使用 UI 提示呼叫端仍可見，保留行為一致性
                System.Windows.Forms.MessageBox.Show("無法開啟檔案：" + ex.Message); // 以訊息方塊顯示錯誤訊息給使用者
            } // 結束 catch 區塊
        }

        /// <summary>
        /// 嘗試更新指定檔案的最後存取與最後寫入時間（含 UTC），並在失敗時回傳 false 而不拋出例外。
        /// </summary>
        /// <param name="path">要設定時間的檔案或目錄完整路徑（UTF-8 可讀路徑）。</param>
        /// <param name="localTime">以當地時間表示欲設定的時間；方法會將此時間轉換為 UTC 傳給原生或 managed API。</param>
        /// <returns>
        /// 傳回 bool：若成功以 managed 或 native 方法設定時間則為 true；若所有嘗試失敗或發生例外則為 false。
        /// </returns>
        /// <remarks>
        /// - 方法流程：先以 managed API (File.SetLastAccessTime / SetLastWriteTime) 嘗試設定時間；若遭遇授權或 IO 例外，會再呼叫原生備援方法 TrySetFileTimesNative 嘗試設定時間。
        /// - 此方法會將錯誤情況以 Console 與 AppendLog 記錄，但不會向上拋出例外，維持對呼叫端的非侵入性（fail-safe）行為。
        /// - 呼叫端可依回傳值與日誌判斷是否需要進一步處理（例如提示使用者或重試）。
        /// </remarks>
        private static bool TrySetFileTimes(string path, DateTime localTime)
        {
            try // 嘗試以 managed API 設定檔案時間，若成功則直接回傳 true
            {
                DateTime utc = localTime.ToUniversalTime(); // 將傳入的本地時間轉換為 UTC（供後續設定 UTC 時間戳使用）
                File.SetLastAccessTime(path, localTime); // 使用 managed API 設定最後存取時間（以 local time）
                File.SetLastWriteTime(path, localTime); // 使用 managed API 設定最後寫入時間（以 local time）
                File.SetLastAccessTimeUtc(path, utc); // 使用 managed API 設定最後存取時間的 UTC 值
                File.SetLastWriteTimeUtc(path, utc); // 使用 managed API 設定最後寫入時間的 UTC 值
                AppendLog($"Managed SetFileTimes 成功: {path}"); // 記錄成功狀態到日誌，便於診斷與追蹤
                return true; // managed 設定成功，回傳 true
            }
            catch (UnauthorizedAccessException uaEx) // 若發生權限不足（無法存取檔案）則進入此區塊處理
            {
                Console.WriteLine($"無法存取檔案以更新時間: {path} => {uaEx.Message}"); // 將錯誤訊息輸出到 Console 以便開發時看到
                AppendLog($"Managed SetFileTimes UnauthorizedAccess: {path} => {uaEx.Message}"); // 將詳細錯誤寫入日誌供後續分析
                // 嘗試使用原生備援方法設定時間（可能在某些權限情況或檔案鎖定情境下有效）
                if (TrySetFileTimesNative(path, localTime)) // 若原生備援成功，記錄並回傳 true
                {
                    AppendLog($"Native SetFileTime 成功: {path}"); // 記錄 native 成功
                    return true; // 回傳成功
                }
                return false; // managed 與 native 都失敗，回傳 false
            }
            catch (IOException ioEx) // 若發生 IO 相關錯誤（例如檔案被佔用），則進入此處理區塊
            {
                Console.WriteLine($"IO 錯誤，無法更新檔案時間: {path} => {ioEx.Message}"); // 顯示 IO 錯誤訊息於 Console
                AppendLog($"Managed SetFileTimes IOEx: {path} => {ioEx.Message}"); // 將 IO 錯誤記錄到日誌
                if (TrySetFileTimesNative(path, localTime)) // 再嘗試原生備援方法
                {
                    AppendLog($"Native SetFileTime 成功: {path}"); // 記錄 native 成功
                    return true; // 回傳成功
                }
                return false; // 仍然失敗則回傳 false
            }
            catch (Exception ex) // 捕捉其他任何未預期的例外，確保方法不會拋出例外給呼叫端
            {
                Console.WriteLine($"更新檔案時間失敗: {path} => {ex.Message}"); // 將錯誤訊息輸出到 Console 協助偵錯
                AppendLog($"Managed SetFileTimes UnknownEx: {path} => {ex.Message}"); // 記錄未知例外到日誌
                if (TrySetFileTimesNative(path, localTime)) // 最後嘗試原生備援方法一次
                {
                    AppendLog($"Native SetFileTime 成功: {path}"); // 記錄 native 成功
                    return true; // 若成功回傳 true
                }
                return false; // 若仍失敗則回傳 false，維持非侵入性行為
            }
        }

        #region
        [StructLayout(LayoutKind.Sequential)] // 指示欄位以宣告順序排列於記憶體中（與 Windows FILETIME 二進位格式相容） // 中文註解
        /// <summary>
        /// 表示 Windows 原生 FILETIME 結構的對應 managed 結構，
        /// 用於與原生 API (例如 SetFileTime) 互動時傳遞時間值的低／高 32 位元部分。
        /// </summary>
        /// <remarks>
        /// - 此結構對應 Windows API 中的 FILETIME（64-bit 時間值拆成低 32-bit 與高 32-bit）。
        /// - 在呼叫 SetFileTime 或其他需要 FILETIME 的 P/Invoke 時，可以以此結構傳遞時間值的低/高部份。
        /// - 為了與原生資料結構完全對齊，使用 Sequential layout 並維持欄位順序與型別（uint）。
        /// - 注意：此結構不包含任何方法或屬性，僅用作資料攜帶；組合 64-bit 時間須由呼叫端自行處理（例如透過 ToFileTimeUtc 取得 long 再拆分）。
        /// </remarks>
        private struct FILETIME // 封閉命名空間內的 private 結構，用於 P/Invoke 呼叫時的資料映射 // 中文註解
        {
            /// <summary>
            /// FILETIME 的低 32 位元（unsigned 32-bit），代表 64-bit 時間值的低位元組部分。
            /// </summary>
            public uint dwLowDateTime; // 低 32 位元：對應 Windows FILETIME 的 dwLowDateTime 欄位（無號整數） // 中文註解

            /// <summary>
            /// FILETIME 的高 32 位元（unsigned 32-bit），代表 64-bit 時間值的高位元組部分。
            /// </summary>
            public uint dwHighDateTime; // 高 32 位元：對應 Windows FILETIME 的 dwHighDateTime 欄位（無號整數） // 中文註解
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] // P/Invoke 屬性：指定載入 kernel32.dll、使用 Unicode 字串、並於失敗時設置 Win32 錯誤碼 // 每行中文註解
        /// <summary>  // XML 文件註解（方法用途）
        /// 以原生 Windows API CreateFileW 開啟或取得檔案／目錄的句柄。  // 中文說明方法用途
        /// </summary>  // 結束 summary
        /// <param name="lpFileName">要開啟的檔案或目錄路徑（Unicode 字串）。</param>  // 參數說明
        /// <param name="dwDesiredAccess">所需的存取權限（例如 FILE_WRITE_ATTRIBUTES）。</param>  // 參數說明
        /// <param name="dwShareMode">共用模式（例如允許讀/寫/刪除共用）。</param>  // 參數說明
        /// <param name="lpSecurityAttributes">安全性屬性指標，通常為 IntPtr.Zero（預設）。</param>  // 參數說明
        /// <param name="dwCreationDisposition">檔案建立或開啟行為（例如 OPEN_EXISTING）。</param>  // 參數說明
        /// <param name="dwFlagsAndAttributes">檔案旗標與屬性（例如 FILE_FLAG_BACKUP_SEMANTICS）。</param>  // 參數說明
        /// <param name="hTemplateFile">範本檔案句柄（通常為 IntPtr.Zero）。</param>  // 參數說明
        /// <returns>成功時回傳有效的檔案句柄（非 IntPtr.Zero 與非 INVALID_HANDLE_VALUE）；失敗時回傳 IntPtr.Zero 或 INVALID_HANDLE_VALUE，並可透過 Marshal.GetLastWin32Error() 取得錯誤碼。</returns>  // 回傳值說明
        /// <remarks>  // 補充說明
        /// - 呼叫端需檢查回傳句柄是否為 IntPtr.Zero 或 new IntPtr(-1) 以判定失敗。  // 注意事項
        /// - 如釋放句柄請務必呼叫 CloseHandle 以避免資源遺漏。  // 注意事項
        /// </remarks>  // 結束 remarks
        private static extern IntPtr CreateFileW(  // 宣告 CreateFileW 的 P/Invoke 簽章，回傳 Win32 句柄（IntPtr） // 每行中文註解
            string lpFileName,                      // 要開啟的檔案路徑（Unicode） // 每行中文註解
            uint dwDesiredAccess,                   // 要求的存取權限（例如 FILE_WRITE_ATTRIBUTES） // 每行中文註解
            uint dwShareMode,                       // 指定共用模式（允許其他程序共用存取） // 每行中文註解
            IntPtr lpSecurityAttributes,            // 安全性屬性，通常為 IntPtr.Zero // 每行中文註解
            uint dwCreationDisposition,             // 建立或開啟方式（如 OPEN_EXISTING） // 每行中文註解
            uint dwFlagsAndAttributes,              // 檔案旗標或屬性（如 FILE_FLAG_BACKUP_SEMANTICS） // 每行中文註解
            IntPtr hTemplateFile);                  // 範本檔案句柄，常為 IntPtr.Zero // 每行中文註解

        [DllImport("kernel32.dll", SetLastError = true)]
        /// <summary>  // XML 文件註解（方法用途）
        /// 使用 SetFileTime 對已開啟的檔案句柄設定建立時間、最後存取時間與最後寫入時間（此處僅使用最後存取與最後寫入時間的參考簽章）。  // 中文說明方法用途
        /// </summary>  // 結束 summary
        /// <param name="hFile">由 CreateFileW 或其他 API 回傳的檔案句柄。</param>  // 參數說明
        /// <param name="lpCreationTime">指向 FILETIME 的指標以設定建立時間，若不需更改可傳入 IntPtr.Zero。</param>  // 參數說明
        /// <param name="lpLastAccessTime">參考型別 FILETIME，用以設定最後存取時間。</param>  // 參數說明
        /// <param name="lpLastWriteTime">參考型別 FILETIME，用以設定最後寫入時間。</param>  // 參數說明
        /// <returns>成功回傳 true；失敗回傳 false，並可透過 Marshal.GetLastWin32Error() 取得錯誤碼。</returns>  // 回傳值說明
        /// <remarks>呼叫前應確保傳入的句柄有效，呼叫後仍需關閉句柄以釋放資源。</remarks>  // 補充說明
        private static extern bool SetFileTime(  // 宣告 SetFileTime 的 P/Invoke 簽章，回傳是否成功 // 每行中文註解
            IntPtr hFile,                          // 已開啟的檔案或目錄句柄 // 每行中文註解
            IntPtr lpCreationTime,                 // 建立時間指標，若不需修改傳入 IntPtr.Zero // 每行中文註解
            ref FILETIME lpLastAccessTime,         // 參考型別 FILETIME：最後存取時間 // 每行中文註解
            ref FILETIME lpLastWriteTime);         // 參考型別 FILETIME：最後寫入時間 // 每行中文註解

        [DllImport("kernel32.dll", SetLastError = true)]
        /// <summary>  // XML 文件註解（方法用途）
        /// 關閉由原生 API 開啟的物件句柄（例如 CreateFileW 回傳的句柄），釋放系統資源。  // 中文說明方法用途
        /// </summary>  // 結束 summary
        /// <param name="hObject">欲關閉的物件句柄。</param>  // 參數說明
        /// <returns>成功回傳 true；失敗回傳 false，可透過 Marshal.GetLastWin32Error() 取得詳細錯誤碼。</returns>  // 回傳值說明
        private static extern bool CloseHandle(IntPtr hObject); // 宣告 CloseHandle 的 P/Invoke 簽章以關閉句柄 // 每行中文註解

        /// <summary>  // XML 文件註解（常數用途）
        /// 要求寫入檔案屬性（FILE_WRITE_ATTRIBUTES），用於修改檔案時間等屬性時的存取權限。  // 中文說明常數用途
        /// </summary>  // 結束 summary
        private const uint FILE_WRITE_ATTRIBUTES = 0x0100; // FILE_WRITE_ATTRIBUTES 權限：允許修改檔案屬性（例如時間戳記） // 每行中文註解
        /// <summary>  // XML 文件註解（常數用途）
        /// OPEN_EXISTING：只開啟已存在的檔案，若檔案不存在則開啟失敗。  // 中文說明常數用途
        /// </summary>  // 結束 summary
        private const uint OPEN_EXISTING = 3; // OPEN_EXISTING 常數值：只開啟已存在的資源 // 每行中文註解
        /// <summary>  // XML 文件註解（常數用途）
        /// FILE_FLAG_BACKUP_SEMANTICS：允許以 CreateFile 開啟目錄，且對備份/還原操作友好。  // 中文說明常數用途
        /// </summary>  // 結束 summary
        private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000; // FLAG：讓 CreateFile 能開啟目錄（必要於目錄時間修改） // 每行中文註解
        /// <summary>  // XML 文件註解（常數用途）
        /// FILE_SHARE_READ：允許其他程序以讀取方式共用檔案。  // 中文說明常數用途
        /// </summary>  // 結束 summary
        private const uint FILE_SHARE_READ = 0x00000001; // 共用讀取權限常數 // 每行中文註解
        /// <summary>  // XML 文件註解（常數用途）
        /// FILE_SHARE_WRITE：允許其他程序以寫入方式共用檔案。  // 中文說明常數用途
        /// </summary>  // 結束 summary
        private const uint FILE_SHARE_WRITE = 0x00000002; // 共用寫入權限常數 // 每行中文註解
        /// <summary>  // XML 文件註解（常數用途）
        /// FILE_SHARE_DELETE：允許其他程序刪除該檔案或目錄（在某些情況下需授權以避免鎖定）。  // 中文說明常數用途
        /// </summary>  // 結束 summary
        private const uint FILE_SHARE_DELETE = 0x00000004; // 共用刪除權限常數 // 每行中文註解

        /// <summary>
        /// 使用原生 Windows API 嘗試更新指定檔案的最後存取與最後寫入時間（備援方法）。
        /// </summary>
        /// <param name="path">欲更新時間的檔案完整路徑。</param>
        /// <param name="localTime">以本地時間表示欲設定的時間（會轉成 UTC 傳給 SetFileTime）。</param>
        /// <returns>
        /// 若成功設定時間則回傳 true；若失敗或發生例外則回傳 false。此方法會在失敗時以 AppendLog 記錄診斷訊息，但不會拋出例外給呼叫端，維持非侵入性行為（fail-safe）。
        /// </returns>
        /// <remarks>
        /// - 此方法使用 CreateFileW 以 FILE_WRITE_ATTRIBUTES 權限開啟檔案（或目錄）句柄，允許變更檔案時間戳記；若檔案為目錄則需搭配 FILE_FLAG_BACKUP_SEMANTICS。
        /// - 若 CreateFileW 或 SetFileTime 發生錯誤，方法會使用 Marshal.GetLastWin32Error() 讀取最後的 Win32 錯誤碼並記錄到日誌，最後回傳失敗結果（false）。
        /// - 為避免拋出例外影響呼叫端，本方法會捕捉所有例外並在內部記錄，然後回傳 false。呼叫端若需要更詳細的錯誤處理可依回傳值與日誌判斷原因。
        /// - 注意：呼叫方應確保提供的 path 是可存取的，本方法不會額外嘗試調整檔案權限或解鎖檔案占用情況。
        /// </remarks>
        private static bool TrySetFileTimesNative(string path, DateTime localTime)
        {
            try // 嘗試區塊：包住所有原生呼叫以捕捉任何例外，避免拋出到呼叫端
            {
                // 使用 CreateFileW 開啟檔案或目錄的句柄，要求 FILE_WRITE_ATTRIBUTES 權限以便設定時間戳記
                IntPtr h = CreateFileW(
                    path,                                                       // 要開啟的檔案路徑
                    FILE_WRITE_ATTRIBUTES,                                      // 要求寫入檔案屬性（用於設定時間）
                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,     // 允許其他程序讀/寫/刪除共用
                    IntPtr.Zero,                                                // 預設安全性屬性
                    OPEN_EXISTING,                                              // 只開啟已存在的檔案
                    FILE_FLAG_BACKUP_SEMANTICS,                                 // 若為目錄亦可開啟（備援旗標）
                    IntPtr.Zero);                                               // 無範本檔案句柄

                // 檢查 CreateFileW 是否失敗（返回 INVALID_HANDLE_VALUE 或 NULL）
                if (h == IntPtr.Zero || h == new IntPtr(-1))
                {
                    // 取得最後的 Win32 錯誤碼以便紀錄與除錯
                    int err = Marshal.GetLastWin32Error(); // 讀取 CreateFileW 失敗的錯誤碼
                    // 將錯誤資訊寫入應用程式日誌（AppendLog 為靜默記錄，不會拋出）
                    AppendLog($"CreateFileW 失敗: {path} => Error {err}");
                    // 無法取得有效句柄，回傳失敗
                    return false;
                }

                // 將傳入的本地時間轉換為用於 SetFileTime 的 FILETIME（使用 UTC）
                long fileTime = localTime.ToFileTimeUtc(); // 轉成 Windows FILETIME（64-bit）
                // 組成 FILETIME 結構的高低 32-bit 字段
                FILETIME ft = new FILETIME
                {
                    dwLowDateTime = (uint)(fileTime & 0xFFFFFFFF),              // 低 32 位元
                    dwHighDateTime = (uint)((fileTime >> 32) & 0xFFFFFFFF)      // 高 32 位元
                };

                // 使用 SetFileTime 設定最後存取與最後寫入時間（不更動建立時間）
                bool ok = SetFileTime(
                    h,                  // 已開啟的檔案句柄
                    IntPtr.Zero,        // 不變更建立時間（傳入 NULL）
                    ref ft,             // 最後存取時間
                    ref ft);            // 最後寫入時間

                // 若 SetFileTime 回傳失敗，讀取錯誤碼並記錄以便診斷
                if (!ok)
                {
                    int err = Marshal.GetLastWin32Error(); // 讀取 SetFileTime 失敗的錯誤碼
                    AppendLog($"SetFileTime 失敗: {path} => Error {err}"); // 紀錄錯誤
                }

                // 關閉剛剛開啟的句柄以釋放系統資源
                CloseHandle(h); // 無論 SetFileTime 成功或失敗都應關閉句柄

                // 回傳 SetFileTime 的結果（成功為 true，失敗為 false）
                return ok; // 回傳實際設定時間的結果
            }
            catch (Exception ex) // 捕捉任何例外，確保不會向上拋出
            {
                // 將例外資訊寫入日誌以供後續調查
                AppendLog($"TrySetFileTimesNative 發生例外: {path} => {ex.Message}");
                // 發生例外則視為失敗並回傳 false
                return false;
            }
        }
        #endregion

        /// <summary>
        /// 將偵錯或診斷訊息附加寫入臨時目錄中的日誌檔案（UTF-8 編碼）。
        /// </summary>
        /// <param name="message">欲寫入的日誌訊息內容。</param>
        /// <remarks>
        /// 此方法以非阻斷方式嘗試將訊息附加到位於系統暫存目錄的 "E-SOP-FileTime.log" 檔案中。
        /// 若寫入失敗，會靜默忽略例外以避免影響主流程；內部會使用 Try/Catch 保護所有 IO 操作並避免拋出例外給呼叫端。
        /// </remarks>
        private static void AppendLog(string message)
        {
            try // 嘗試執行檔案寫入的保護區塊
            {
                // 取得系統暫存目錄的路徑並與檔名合併成完整的日誌檔路徑
                string logPath = Path.Combine(Path.GetTempPath(), "E-SOP-FileTime.log");
                // 組合要寫入的一行文字，包含當前本地時間、分隔符號與傳入的訊息，並在結尾加入換行符
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine;
                // 使用 UTF-8 編碼將文字附加到日誌檔（若檔案不存在會建立之）
                File.AppendAllText(logPath, line, Encoding.UTF8);
            }
            catch // 捕捉所有例外但不進行處理，避免影響主流程
            {
                // 忽略任何記錄失敗，不影響主流程（保留原本的行為：靜默失敗）
            }
        }
    }
}
