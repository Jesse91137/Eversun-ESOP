using System.Collections;

namespace E_SOP
{
    /// <summary>
    /// 記錄 ESOP 相關目錄與狀態的靜態類別。
    /// 用於儲存各種目錄路徑、部門、站別、注意事項、文件清單等全域設定。
    /// </summary>
    public class Production
    {
        /// <summary>
        /// esop初始目錄
        /// </summary>
        public static string soppath; // esop初始目錄

        /// <summary>
        /// 下拉選單1
        /// </summary>
        public static string production1; // 下拉選單1

        /// <summary>
        /// 下拉選單2
        /// </summary>
        public static string production2; // 下拉選單2

        /// <summary>
        /// 量產SOP或試產SOP目錄
        /// </summary>
        public static string sopfile; // 量產SOP.試產SOP

        /// <summary>
        /// 製程別（部門）
        /// </summary>
        public static string deptItem; // 製程別

        /// <summary>
        /// 站別
        /// </summary>
        public static string station; // 站別

        /// <summary>
        /// 注意事項(不貳過)目錄
        /// </summary>
        public static string twicefail; // 注意事項(不貳過)

        /// <summary>
        /// 包裝-組裝注意事項目錄
        /// </summary>
        public static string pack_assy; // 注意事項(包裝-組裝)

        /// <summary>
        /// 包裝-後製程注意事項目錄
        /// </summary>
        public static string pack_proc; // 注意事項(包裝-後製程)

        /// <summary>
        /// 客戶特殊要求(QA-Zone)目錄
        /// </summary>
        public static string qa_zone; // 客戶特殊要求(QA-Zone)

        /// <summary>
        /// SMD點膠圖目錄
        /// </summary>
        public static string dispens; // SMD點膠圖

        /// <summary>
        /// 生產通知單目錄
        /// </summary>
        public static string producnotice; // 生產通知單

        /// <summary>
        /// 生產通知單按鈕狀態
        /// </summary>
        public static string btn_producN; // 生產通知單按鈕

        /// <summary>
        /// EMO發行文件目錄
        /// </summary>
        public static string emo; // EMO發行文件

        /// <summary>
        /// PCB板料號
        /// </summary>
        public static string pcb; // PCB板料號

        /// <summary>
        /// 無生產通知單狀態
        /// </summary>
        public static string noproducnotice;

        /// <summary>
        /// 檔案完整路徑清單（用於查詢結果）
        /// </summary>
        public static ArrayList filepathlist = new ArrayList();
    }
}
