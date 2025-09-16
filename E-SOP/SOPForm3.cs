using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;

namespace E_SOP
{
    /// <summary>
    /// SOPForm3 類別，負責顯示 SOP PDF 文件的視窗表單。
    /// </summary>
    public partial class SOPForm3 : Telerik.WinControls.UI.RadForm
    {
        /// <summary>
        /// 靜態變數，記錄 SOPForm3 實例的數量。
        /// </summary>
        public static int m_Static_InstanceCount = 0;

        /// <summary>
        /// 靜態變數，記錄目前 SOP 的檔案名稱。
        /// </summary>
        public static string SOPName = "";

        /// <summary>
        /// SOPForm3 建構函式，初始化元件。
        /// </summary>
        public SOPForm3()
        {
            InitializeComponent();
        }

        /// <summary>
        /// SOPForm3 載入事件處理函式。
        /// 檢查是否在設計模式下，若否則載入指定的 PDF 檔案。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void SOPForm3_Load(object sender, EventArgs e)
        {
            // 檢查是否在設計模式中運行
            if (!DesignMode)
            {
                // 設定 PDF 控制元件的檔案來源路徑
                axAcroPDF3.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
            }
        }

        /// <summary>
        /// SOPForm3 關閉事件處理函式。
        /// 關閉時重設實例計數。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void SOPForm3_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 關閉表單時將實例計數歸零
            m_Static_InstanceCount = 0;
        }
    }
}
