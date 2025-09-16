using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace E_SOP
{
    /// <summary>
    /// SOPForm6 表單，負責顯示 PDF 文件並根據 INI 設定調整視窗行為。
    /// </summary>
    public partial class SOPForm6 : Telerik.WinControls.UI.RadForm
    {
        /// <summary>
        /// 靜態實例計數器，用於追蹤 SOPForm6 的實例數量。
        /// </summary>
        public static int m_Static_InstanceCount = 0;

        /// <summary>
        /// SOP 文件名稱，供 PDF 控制元件載入使用。
        /// </summary>
        public static string SOPName = "";

        /// <summary>
        /// INI 設定檔名稱。
        /// </summary>
        public string filename = "Setup.ini";

        /// <summary>
        /// 雙螢幕設定值，從 INI 讀取。
        /// </summary>
        public string Double;

        /// <summary>
        /// INI 檔案操作物件。
        /// </summary>
        SetupIniIP ini = new SetupIniIP();

        /// <summary>
        /// SOPForm6 建構式，初始化元件。
        /// </summary>
        public SOPForm6()
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
        /// SOPForm6 載入事件，根據 INI 設定調整視窗位置與載入 PDF。
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void SOPForm6_Load(object sender, EventArgs e)
        {
            Double = ini.IniReadValue("Double_Monitor", "Double", filename); // 讀取雙螢幕設定
            if (Double == "ON")
            {
                this.Show();   // 顯示自訂視窗
                int x = Screen.PrimaryScreen.WorkingArea.Width; // 取得螢幕寬度
                int y = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2; // 垂直置中
                this.Location = new Point(x, y); // 設定視窗位置
                axAcroPDF6.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName; // 載入 PDF
            }
            else
            {
                axAcroPDF6.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName; // 載入 PDF
            }
        }

        /// <summary>
        /// SOPForm6 關閉事件，重設實例計數器。
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void SOPForm6_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0; // 重設計數器
        }

        /// <summary>
        /// SOPForm6 顯示事件，根據 INI 設定調整視窗狀態。
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void SOPForm6_Shown(object sender, EventArgs e)
        {
            Double = ini.IniReadValue("Double_Monitor", "Double", filename); // 讀取雙螢幕設定
            if (Double == "ON")
            {
                // 雙螢幕模式不做最大化
            }
            else
            {
                this.WindowState = FormWindowState.Maximized; // 視窗最大化
            }
        }
    }
}
