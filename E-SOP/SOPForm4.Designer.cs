namespace E_SOP
{
    partial class SOPForm4
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SOPForm4));
            this.axAcroPDF4 = new AxAcroPDFLib.AxAcroPDF();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // axAcroPDF4
            // 
            this.axAcroPDF4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axAcroPDF4.Enabled = true;
            this.axAcroPDF4.Location = new System.Drawing.Point(0, 0);
            this.axAcroPDF4.Name = "axAcroPDF4";
            this.axAcroPDF4.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF4.OcxState")));
            this.axAcroPDF4.Size = new System.Drawing.Size(1020, 740);
            this.axAcroPDF4.TabIndex = 0;
            // 
            // SOPForm4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1020, 740);
            this.Controls.Add(this.axAcroPDF4);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SOPForm4";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "SOP4";
            this.ThemeName = "MaterialPink";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SOPForm4_FormClosing);
            this.Load += new System.EventHandler(this.SOPForm4_Load);
            this.Shown += new System.EventHandler(this.SOPForm4_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxAcroPDFLib.AxAcroPDF axAcroPDF4;
    }
}
