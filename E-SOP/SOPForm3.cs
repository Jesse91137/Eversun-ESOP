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
    public partial class SOPForm3 : Telerik.WinControls.UI.RadForm
    {
        public static int m_Static_InstanceCount = 0;
        public static string SOPName = "";
        public SOPForm3()
        {
            InitializeComponent();
        }

        private void SOPForm3_Load(object sender, EventArgs e)
        {
            axAcroPDF3.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
        }
        private void SOPForm3_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0;
        }
    }
}
