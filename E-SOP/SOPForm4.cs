using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using System.IO;
using System.Runtime.InteropServices;

namespace E_SOP
{
    public partial class SOPForm4 : Telerik.WinControls.UI.RadForm
    {
        public static int m_Static_InstanceCount = 0;
        public static string SOPName = "";
        public string filename = "Setup.ini";
        public string Double;
        SetupIniIP ini = new SetupIniIP();
        public SOPForm4()
        {
            InitializeComponent();
        }
        public class SetupIniIP
        { //api ini
            public string path;
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern int GetPrivateProfileString(string section,
            string key, string def, StringBuilder retVal,
            int size, string filePath);
            public void IniWriteValue(string Section, string Key, string Value, string inipath)
            {
                WritePrivateProfileString(Section, Key, Value, Application.StartupPath + "\\" + inipath);
            }
            public string IniReadValue(string Section, string Key, string inipath)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath);
                return temp.ToString();
            }
        }
        private void SOPForm4_Load(object sender, EventArgs e)
        {
            Double = ini.IniReadValue("Double_Monitor", "Double", filename);
            if (Double == "ON")
            {
                this.Show();   //<~若是自定的視窗~可在此處加這行顯示視窗
                int x = Screen.PrimaryScreen.WorkingArea.Width;
                int y = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2;
                this.Location = new Point(x, y);
                axAcroPDF4.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
            }
            else
            {
                axAcroPDF4.src = System.Windows.Forms.Application.StartupPath + "\\" + "Temp" + "\\" + SOPName;
            }
            
        }

        private void SOPForm4_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_Static_InstanceCount = 0;
        }

        private void SOPForm4_Shown(object sender, EventArgs e)
        {
            Double = ini.IniReadValue("Double_Monitor", "Double", filename);
            if (Double == "ON")
            {

            }
            else
            {

                this.WindowState = FormWindowState.Maximized;

            }
        }
    }
}
