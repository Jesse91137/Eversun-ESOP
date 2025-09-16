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
    public partial class SOPForm1 : Telerik.WinControls.UI.RadForm
    {
        
        public static int m_Static_InstanceCount = 0;
        public static string  SOPName  = "";
        public SOPForm1()
        {
            InitializeComponent();
        }

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

        private void SOPForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0;
        }
    }
}
