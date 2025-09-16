using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using System.Diagnostics;
using System.IO;

namespace E_SOP
{
    /// <summary>
    /// SOPForm1 類別，負責顯示 SOP 文件的主視窗。
    /// </summary>
    public partial class SOPForm1 : Telerik.WinControls.UI.RadForm
    {
        /// <summary>
        /// 靜態計數器，記錄 SOPForm1 實例的數量。
        /// </summary>
        public static int m_Static_InstanceCount = 0;

        /// <summary>
        /// 靜態字串，儲存目前 SOP 文件名稱。
        /// </summary>
        public static string SOPName = "";

        /// <summary>
        /// SOPForm1 建構函式，初始化元件。
        /// </summary>
        public SOPForm1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// SOPForm 載入事件，負責載入 PDF 文件至檢視元件。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void SOPForm_Load(object sender, EventArgs e)
        {
            //radPdfViewer1.LoadDocument("C:\\Users\\beer_yang\\source\\repos\\E-SOP\\E-SOP\\bin\\Debug\\i_Program-Training.pdf");

            //OpenFileDialog openFile = new OpenFileDialog();
            ////open.Filter = "PDF檔案|*.pdf";
            //openFile.ShowDialog();
            axAcroPDF1.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
            //axAcroPDF1.src = "C:\\Users\\beer_yang\\source\\repos\\E-SOP\\E-SOP\\bin\\Debug\\i_Program-Training.pdf";
            //axAcroPDF1.LoadFile(of.FileName);
            //RadForm1 RadForm1 = (RadForm1)this.Owner; //取得父視窗的參考;
        }

        /// <summary>
        /// SOPForm 關閉事件，重設 SOPForm1 實例計數器。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void SOPForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0;
        }
    }
}
