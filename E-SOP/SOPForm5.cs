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
    public partial class SOPForm5 : Telerik.WinControls.UI.RadForm
    {
        public static int m_Static_InstanceCount = 0;
        public static string SOPName = "";
        public SOPForm5()
        {
            InitializeComponent();
        }

        private void SOPForm5_Load(object sender, EventArgs e)
        {
            axAcroPDF5.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
        }
        private void SOPForm5_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0;
        }
    }
}
