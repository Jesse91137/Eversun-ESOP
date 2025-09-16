namespace E_SOP
{
    partial class PE_Note
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
            this.radThemeManager1 = new Telerik.WinControls.RadThemeManager();
            this.materialBlueGreyTheme1 = new Telerik.WinControls.Themes.MaterialBlueGreyTheme();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            this.li_note = new Telerik.WinControls.UI.RadListControl();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.li_note)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radButton1
            // 
            this.radButton1.Location = new System.Drawing.Point(712, 253);
            this.radButton1.Name = "radButton1";
            this.radButton1.Size = new System.Drawing.Size(137, 30);
            this.radButton1.TabIndex = 1;
            this.radButton1.Text = "確認";
            this.radButton1.ThemeName = "CrystalDark";
            this.radButton1.Click += new System.EventHandler(this.RadButton1_Click);
            // 
            // li_note
            // 
            this.li_note.ItemHeight = 28;
            this.li_note.Location = new System.Drawing.Point(28, 30);
            this.li_note.Name = "li_note";
            this.li_note.Size = new System.Drawing.Size(821, 193);
            this.li_note.TabIndex = 2;
            this.li_note.ThemeName = "Material";
            // 
            // PE_Note
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 295);
            this.ControlBox = false;
            this.Controls.Add(this.li_note);
            this.Controls.Add(this.radButton1);
            this.Name = "PE_Note";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "PE備註";
            this.ThemeName = "MaterialBlueGrey";
            this.Load += new System.EventHandler(this.PE_Note_Load);
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.li_note)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Telerik.WinControls.RadThemeManager radThemeManager1;
        private Telerik.WinControls.Themes.MaterialBlueGreyTheme materialBlueGreyTheme1;
        private Telerik.WinControls.UI.RadButton radButton1;
        private Telerik.WinControls.UI.RadListControl li_note;
    }
}
