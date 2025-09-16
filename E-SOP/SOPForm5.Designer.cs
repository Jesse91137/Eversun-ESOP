namespace E_SOP
{
    partial class SOPForm5
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SOPForm5));
            this.axAcroPDF5 = new AxAcroPDFLib.AxAcroPDF();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // axAcroPDF5
            // 
            this.axAcroPDF5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axAcroPDF5.Enabled = true;
            this.axAcroPDF5.Location = new System.Drawing.Point(0, 0);
            this.axAcroPDF5.Name = "axAcroPDF5";
            this.axAcroPDF5.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF5.OcxState")));
            this.axAcroPDF5.Size = new System.Drawing.Size(1018, 738);
            this.axAcroPDF5.TabIndex = 0;
            // 
            // SOPForm5
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1018, 738);
            this.Controls.Add(this.axAcroPDF5);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SOPForm5";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "SOP5";
            this.ThemeName = "MaterialPink";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SOPForm5_FormClosing);
            this.Load += new System.EventHandler(this.SOPForm5_Load);
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxAcroPDFLib.AxAcroPDF axAcroPDF5;
    }
}
