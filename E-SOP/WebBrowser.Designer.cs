namespace E_SOP
{
    partial class WebBrowser
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
            this.components = new System.ComponentModel.Container();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.txt_Check_Name = new Telerik.WinControls.UI.RadTextBox();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            this.radLabel1 = new Telerik.WinControls.UI.RadLabel();
            this.radGroupBox1 = new Telerik.WinControls.UI.RadGroupBox();
            this.lbl_msg = new System.Windows.Forms.Label();
            this.radThemeManager1 = new Telerik.WinControls.RadThemeManager();
            this.materialPinkTheme1 = new Telerik.WinControls.Themes.MaterialPinkTheme();
            this.radStatusStrip1 = new Telerik.WinControls.UI.RadStatusStrip();
            this.System_date_ID = new Telerik.WinControls.UI.RadLabelElement();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.txt_Check_Name)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radLabel1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGroupBox1)).BeginInit();
            this.radGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(2, 61);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.ScriptErrorsSuppressed = true;
            this.webBrowser1.Size = new System.Drawing.Size(1051, 523);
            this.webBrowser1.TabIndex = 0;
            this.webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.WebBrowser1_DocumentCompleted);
            // 
            // txt_Check_Name
            // 
            this.txt_Check_Name.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Check_Name.Location = new System.Drawing.Point(203, 20);
            this.txt_Check_Name.Name = "txt_Check_Name";
            this.txt_Check_Name.Size = new System.Drawing.Size(125, 32);
            this.txt_Check_Name.TabIndex = 1;
            this.txt_Check_Name.ThemeName = "ControlDefault";
            // 
            // radButton1
            // 
            this.radButton1.Location = new System.Drawing.Point(344, 16);
            this.radButton1.Name = "radButton1";
            this.radButton1.Size = new System.Drawing.Size(65, 30);
            this.radButton1.TabIndex = 2;
            this.radButton1.Text = "送出";
            this.radButton1.ThemeName = "CrystalDark";
            this.radButton1.Click += new System.EventHandler(this.RadButton1_Click);
            // 
            // radLabel1
            // 
            this.radLabel1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radLabel1.Location = new System.Drawing.Point(134, 21);
            this.radLabel1.Name = "radLabel1";
            this.radLabel1.Size = new System.Drawing.Size(63, 25);
            this.radLabel1.TabIndex = 3;
            this.radLabel1.Text = "確認者:";
            // 
            // radGroupBox1
            // 
            this.radGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this.radGroupBox1.Controls.Add(this.lbl_msg);
            this.radGroupBox1.Controls.Add(this.txt_Check_Name);
            this.radGroupBox1.Controls.Add(this.radLabel1);
            this.radGroupBox1.Controls.Add(this.radButton1);
            this.radGroupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.radGroupBox1.GroupBoxStyle = Telerik.WinControls.UI.RadGroupBoxStyle.Office;
            this.radGroupBox1.HeaderText = "生產注意事項確認";
            this.radGroupBox1.Location = new System.Drawing.Point(0, 0);
            this.radGroupBox1.Name = "radGroupBox1";
            this.radGroupBox1.Size = new System.Drawing.Size(1056, 56);
            this.radGroupBox1.TabIndex = 4;
            this.radGroupBox1.Text = "生產注意事項確認";
            this.radGroupBox1.ThemeName = "MaterialPink";
            // 
            // lbl_msg
            // 
            this.lbl_msg.AutoSize = true;
            this.lbl_msg.ForeColor = System.Drawing.Color.Red;
            this.lbl_msg.Location = new System.Drawing.Point(460, 21);
            this.lbl_msg.Name = "lbl_msg";
            this.lbl_msg.Size = new System.Drawing.Size(226, 22);
            this.lbl_msg.TabIndex = 4;
            this.lbl_msg.Text = "此機種無生產注意事項資料";
            // 
            // radStatusStrip1
            // 
            this.radStatusStrip1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.System_date_ID});
            this.radStatusStrip1.Location = new System.Drawing.Point(0, 590);
            this.radStatusStrip1.Name = "radStatusStrip1";
            this.radStatusStrip1.Size = new System.Drawing.Size(1056, 25);
            this.radStatusStrip1.TabIndex = 9;
            this.radStatusStrip1.ThemeName = "MaterialPink";
            // 
            // System_date_ID
            // 
            this.System_date_ID.Name = "System_date_ID";
            this.radStatusStrip1.SetSpring(this.System_date_ID, false);
            this.System_date_ID.Text = "System_date_ID";
            this.System_date_ID.TextWrap = true;
            this.System_date_ID.UseCompatibleTextRendering = false;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.Timer1_Tick);
            // 
            // WebBrowser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1056, 615);
            this.ControlBox = false;
            this.Controls.Add(this.radStatusStrip1);
            this.Controls.Add(this.radGroupBox1);
            this.Controls.Add(this.webBrowser1);
            this.MinimizeBox = false;
            this.Name = "WebBrowser";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "不二過警示";
            this.ThemeName = "MaterialPink";
            this.Load += new System.EventHandler(this.WebBrowser_Load);
            ((System.ComponentModel.ISupportInitialize)(this.txt_Check_Name)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radLabel1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGroupBox1)).EndInit();
            this.radGroupBox1.ResumeLayout(false);
            this.radGroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser1;
        private Telerik.WinControls.UI.RadTextBox txt_Check_Name;
        private Telerik.WinControls.UI.RadButton radButton1;
        private Telerik.WinControls.UI.RadLabel radLabel1;
        private Telerik.WinControls.UI.RadGroupBox radGroupBox1;
        private Telerik.WinControls.RadThemeManager radThemeManager1;
        private Telerik.WinControls.Themes.MaterialPinkTheme materialPinkTheme1;
        private Telerik.WinControls.UI.RadStatusStrip radStatusStrip1;
        private Telerik.WinControls.UI.RadLabelElement System_date_ID;
        private System.Windows.Forms.Label lbl_msg;
        private System.Windows.Forms.Timer timer1;
    }
}
