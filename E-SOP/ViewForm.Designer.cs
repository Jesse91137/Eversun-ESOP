
namespace E_SOP
{
    partial class ViewForm
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lab1 = new System.Windows.Forms.Label();
            this.lab_noteasy = new System.Windows.Forms.Label();
            this.lab4 = new System.Windows.Forms.Label();
            this.lab2 = new System.Windows.Forms.Label();
            this.lab3 = new System.Windows.Forms.Label();
            this.btn_up = new System.Windows.Forms.Button();
            this.btn_down = new System.Windows.Forms.Button();
            this.lab_count = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lab_notice = new System.Windows.Forms.Label();
            this.lab_item = new System.Windows.Forms.Label();
            this.lab_engsr = new System.Windows.Forms.Label();
            this.lab_spec = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(33, 245);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(431, 193);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // lab1
            // 
            this.lab1.AutoSize = true;
            this.lab1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab1.Location = new System.Drawing.Point(45, 26);
            this.lab1.Name = "lab1";
            this.lab1.Size = new System.Drawing.Size(85, 20);
            this.lab1.TabIndex = 1;
            this.lab1.Text = "機種名稱 : ";
            // 
            // lab_noteasy
            // 
            this.lab_noteasy.AutoSize = true;
            this.lab_noteasy.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_noteasy.Location = new System.Drawing.Point(45, 62);
            this.lab_noteasy.Name = "lab_noteasy";
            this.lab_noteasy.Size = new System.Drawing.Size(25, 20);
            this.lab_noteasy.TabIndex = 3;
            this.lab_noteasy.Text = "有";
            // 
            // lab4
            // 
            this.lab4.AutoSize = true;
            this.lab4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab4.Location = new System.Drawing.Point(45, 97);
            this.lab4.Name = "lab4";
            this.lab4.Size = new System.Drawing.Size(33, 20);
            this.lab4.TabIndex = 4;
            this.lab4.Text = "需  ";
            // 
            // lab2
            // 
            this.lab2.AutoSize = true;
            this.lab2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab2.Location = new System.Drawing.Point(45, 135);
            this.lab2.Name = "lab2";
            this.lab2.Size = new System.Drawing.Size(53, 20);
            this.lab2.TabIndex = 5;
            this.lab2.Text = "料號 : ";
            // 
            // lab3
            // 
            this.lab3.AutoSize = true;
            this.lab3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab3.Location = new System.Drawing.Point(45, 178);
            this.lab3.Name = "lab3";
            this.lab3.Size = new System.Drawing.Size(53, 20);
            this.lab3.TabIndex = 6;
            this.lab3.Text = "規格 : ";
            // 
            // btn_up
            // 
            this.btn_up.Location = new System.Drawing.Point(33, 216);
            this.btn_up.Name = "btn_up";
            this.btn_up.Size = new System.Drawing.Size(75, 23);
            this.btn_up.TabIndex = 7;
            this.btn_up.Text = "上一筆";
            this.btn_up.UseVisualStyleBackColor = true;
            this.btn_up.Click += new System.EventHandler(this.btn_down_Click);
            // 
            // btn_down
            // 
            this.btn_down.Location = new System.Drawing.Point(389, 216);
            this.btn_down.Name = "btn_down";
            this.btn_down.Size = new System.Drawing.Size(75, 23);
            this.btn_down.TabIndex = 8;
            this.btn_down.Text = "下一筆";
            this.btn_down.UseVisualStyleBackColor = true;
            this.btn_down.Click += new System.EventHandler(this.btn_down_Click);
            // 
            // lab_count
            // 
            this.lab_count.AutoSize = true;
            this.lab_count.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_count.ForeColor = System.Drawing.Color.Red;
            this.lab_count.Location = new System.Drawing.Point(79, 62);
            this.lab_count.Name = "lab_count";
            this.lab_count.Size = new System.Drawing.Size(25, 20);
            this.lab_count.TabIndex = 9;
            this.lab_count.Text = "幾";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(111, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 20);
            this.label1.TabIndex = 10;
            this.label1.Text = "顆零件不易維修";
            // 
            // lab_notice
            // 
            this.lab_notice.AutoSize = true;
            this.lab_notice.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_notice.ForeColor = System.Drawing.Color.Red;
            this.lab_notice.Location = new System.Drawing.Point(71, 97);
            this.lab_notice.Name = "lab_notice";
            this.lab_notice.Size = new System.Drawing.Size(41, 20);
            this.lab_notice.TabIndex = 11;
            this.lab_notice.Text = "注意";
            // 
            // lab_item
            // 
            this.lab_item.AutoSize = true;
            this.lab_item.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_item.ForeColor = System.Drawing.Color.Red;
            this.lab_item.Location = new System.Drawing.Point(107, 135);
            this.lab_item.Name = "lab_item";
            this.lab_item.Size = new System.Drawing.Size(73, 20);
            this.lab_item.TabIndex = 12;
            this.lab_item.Text = "料號在這";
            // 
            // lab_engsr
            // 
            this.lab_engsr.AutoSize = true;
            this.lab_engsr.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_engsr.ForeColor = System.Drawing.Color.Red;
            this.lab_engsr.Location = new System.Drawing.Point(136, 26);
            this.lab_engsr.Name = "lab_engsr";
            this.lab_engsr.Size = new System.Drawing.Size(105, 20);
            this.lab_engsr.TabIndex = 13;
            this.lab_engsr.Text = "機種名稱在這";
            // 
            // lab_spec
            // 
            this.lab_spec.AutoSize = true;
            this.lab_spec.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_spec.ForeColor = System.Drawing.Color.Red;
            this.lab_spec.Location = new System.Drawing.Point(107, 178);
            this.lab_spec.Name = "lab_spec";
            this.lab_spec.Size = new System.Drawing.Size(73, 20);
            this.lab_spec.TabIndex = 14;
            this.lab_spec.Text = "規格在這";
            // 
            // ViewForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(506, 450);
            this.Controls.Add(this.lab_spec);
            this.Controls.Add(this.lab_engsr);
            this.Controls.Add(this.lab_item);
            this.Controls.Add(this.lab_notice);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lab_count);
            this.Controls.Add(this.btn_down);
            this.Controls.Add(this.btn_up);
            this.Controls.Add(this.lab3);
            this.Controls.Add(this.lab2);
            this.Controls.Add(this.lab4);
            this.Controls.Add(this.lab_noteasy);
            this.Controls.Add(this.lab1);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "ViewForm";
            this.Text = "ViewForm";
            this.Load += new System.EventHandler(this.ViewForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lab1;
        private System.Windows.Forms.Label lab_noteasy;
        private System.Windows.Forms.Label lab4;
        private System.Windows.Forms.Label lab2;
        private System.Windows.Forms.Label lab3;
        private System.Windows.Forms.Button btn_up;
        private System.Windows.Forms.Button btn_down;
        private System.Windows.Forms.Label lab_count;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lab_notice;
        private System.Windows.Forms.Label lab_item;
        private System.Windows.Forms.Label lab_engsr;
        private System.Windows.Forms.Label lab_spec;
    }
}