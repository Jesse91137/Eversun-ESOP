using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace E_SOP
{
    /// <summary>
    /// SOPForm2 表單，負責顯示 SOP PDF，並根據 INI 設定檔決定視窗顯示方式。
    /// </summary>
    public partial class SOPForm2 : Telerik.WinControls.UI.RadForm
    {
        /// <summary>
        /// SOPForm2 實例計數器 (靜態)，用於追蹤視窗開啟狀態。
        /// </summary>
        public static int m_Static_InstanceCount = 0;

        /// <summary>
        /// SOP 檔案名稱 (靜態)，用於指定要顯示的 PDF。
        /// </summary>
        public static string SOPName = "";

        /// <summary>
        /// INI 設定檔名稱。
        /// </summary>
        public string filename = "Setup.ini";

        /// <summary>
        /// 雙螢幕設定值，ON 表示啟用雙螢幕。
        /// </summary>
        public string Double;

        /// <summary>
        /// INI 檔案操作物件。
        /// </summary>
        SetupIniIP ini = new SetupIniIP();

        /// <summary>
        /// SOPForm2 建構式，初始化元件。
        /// </summary>
        public SOPForm2()
        {
            InitializeComponent();
        }

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
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath);
                return temp.ToString();
            }
        }

        /// <summary>
        /// SOPForm2 載入事件，根據 INI 設定決定視窗顯示方式並載入 PDF。
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void SOPForm2_Load(object sender, EventArgs e)
        {
            Double = ini.IniReadValue("Double_Monitor", "Double", filename); // 讀取雙螢幕設定
            if (Double == "ON")
            {
                this.Show();   // 顯示自訂視窗
                int x = Screen.PrimaryScreen.WorkingArea.Width; // 取得螢幕寬度
                int y = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2; // 計算垂直置中位置
                this.Location = new Point(x, y); // 設定視窗位置於右側
                axAcroPDF2.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName; // 載入 PDF
            }
            else
            {
                this.WindowState = FormWindowState.Maximized; // 最大化視窗
                axAcroPDF2.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName; // 載入 PDF
            }
        }

        /// <summary>
        /// SOPForm2 關閉事件，重設實例計數器。
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void SOPForm2_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0; // 重設計數器
        }

        /// <summary>
        /// SOPForm2 顯示事件，根據 INI 設定調整視窗狀態。
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void SOPForm2_Shown(object sender, EventArgs e)
        {
            Double = ini.IniReadValue("Double_Monitor", "Double", filename); // 讀取雙螢幕設定
            if (Double == "ON")
            {
                // 雙螢幕模式不做額外處理
            }
            else
            {
                this.WindowState = FormWindowState.Maximized; // 最大化視窗
            }
        }
    }
}
