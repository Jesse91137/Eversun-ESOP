using System;
using System.Windows.Forms;

namespace E_SOP
{
    /// <summary>
    /// SOPForm5 表單，負責顯示 PDF 文件。
    /// </summary>
    public partial class SOPForm5 : Telerik.WinControls.UI.RadForm
    {
        /// <summary>
        /// 靜態實例計數器，用於追蹤 SOPForm5 的實例數量。
        /// </summary>
        public static int m_Static_InstanceCount = 0; // 靜態變數，記錄目前表單實例數量

        /// <summary>
        /// 靜態 SOP 名稱，指定要顯示的 PDF 檔案名稱。
        /// </summary>
        public static string SOPName = ""; // 靜態變數，記錄要顯示的 SOP 檔案名稱

        /// <summary>
        /// SOPForm5 建構函式，初始化元件。
        /// </summary>
        public SOPForm5()
        {
            InitializeComponent(); // 初始化表單元件
        }

        /// <summary>
        /// SOPForm5 載入事件，設定 PDF 控制元件的檔案來源。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void SOPForm5_Load(object sender, EventArgs e)
        {
            // 設定 PDF 控制元件的檔案來源為 Temp 資料夾下的 SOPName 檔案
            axAcroPDF5.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
        }

        /// <summary>
        /// SOPForm5 關閉事件，重設實例計數器。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void SOPForm5_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0; // 關閉表單時，重設實例計數器
        }
    }
}
