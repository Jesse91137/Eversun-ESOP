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
    public partial class PE_Note : Telerik.WinControls.UI.RadForm
    {
        public static string Note = "";
        public PE_Note()
        {
            InitializeComponent();
        }

        private void PE_Note_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Note))
            {
                string[] SOPS = Note.Trim().Split('\n');
                for (int i = 0; i < SOPS.Length; i++)
                {
                    li_note.Items.Add(SOPS[i].ToString());
                }
            }
            else
            {
                Close();
            }
           // txt_note.Text = Note;
        }

        private void RadButton1_Click(object sender, EventArgs e)
        {
            try
            {
                
                DialogResult = DialogResult.OK;
                Close();
            }
            catch
            {
                DialogResult = DialogResult.Cancel;
            }
        }
    }
}
