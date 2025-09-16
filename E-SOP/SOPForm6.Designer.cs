namespace E_SOP
{
    partial class SOPForm6
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SOPForm6));
            this.axAcroPDF6 = new AxAcroPDFLib.AxAcroPDF();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // axAcroPDF6
            // 
            this.axAcroPDF6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axAcroPDF6.Enabled = true;
            this.axAcroPDF6.Location = new System.Drawing.Point(0, 0);
            this.axAcroPDF6.Name = "axAcroPDF6";
            this.axAcroPDF6.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF6.OcxState")));
            this.axAcroPDF6.Size = new System.Drawing.Size(1020, 740);
            this.axAcroPDF6.TabIndex = 0;
            // 
            // SOPForm6
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1020, 740);
            this.Controls.Add(this.axAcroPDF6);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SOPForm6";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "SOP6";
            this.ThemeName = "MaterialPink";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SOPForm6_FormClosing);
            this.Load += new System.EventHandler(this.SOPForm6_Load);
            this.Shown += new System.EventHandler(this.SOPForm6_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxAcroPDFLib.AxAcroPDF axAcroPDF6;
    }
}
