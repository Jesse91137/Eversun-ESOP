namespace E_SOP
{
    partial class SOPForm2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SOPForm2));
            this.axAcroPDF2 = new AxAcroPDFLib.AxAcroPDF();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // axAcroPDF2
            // 
            this.axAcroPDF2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axAcroPDF2.Enabled = true;
            this.axAcroPDF2.Location = new System.Drawing.Point(0, 0);
            this.axAcroPDF2.Name = "axAcroPDF2";
            this.axAcroPDF2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF2.OcxState")));
            this.axAcroPDF2.Size = new System.Drawing.Size(1020, 740);
            this.axAcroPDF2.TabIndex = 0;
            // 
            // SOPForm2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1020, 740);
            this.Controls.Add(this.axAcroPDF2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SOPForm2";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "SOP2";
            this.ThemeName = "MaterialPink";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SOPForm2_FormClosing);
            this.Load += new System.EventHandler(this.SOPForm2_Load);
            this.Shown += new System.EventHandler(this.SOPForm2_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxAcroPDFLib.AxAcroPDF axAcroPDF2;
    }
}
