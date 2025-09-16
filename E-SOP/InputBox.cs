using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Telerik.WinControls;


namespace E_SOP
{
    public partial class InputBox : Telerik.WinControls.UI.RadForm
    {
        public InputBox()
        {
            this.ControlBox = false;
            InitializeComponent();
        }        
        public string MsgModel
        {
            set
            {
                label1.Text = value;
            }
            get
            {
                return label1.Text;
            }
        }
        public string MsgWip_No
        {
            set
            {
                label2.Text = value;
            }
            get
            {
                return label2.Text;
            }
        }
        
        public string MsgRoute
        {
            set
            {
                label3.Text = value;
            }
            get
            {
                return label3.Text;
            }
        }
        /// <summary>
        /// 登入者確認
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radButton1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Regex.IsMatch(txt_Check_Name.Text, @"^[0-9]+$"))
                {
                    MessageBox.Show("請輸入數字!!");
                    return;
                }
                if (txt_Check_Name.Text != "")
                {
                    while (txt_Check_Name.Text.Length < 5)
                    {
                        txt_Check_Name.Text = "0" + txt_Check_Name.Text;
                    }

                    string sqlstr = @"select * from i_Factory_EversunUser_Tabel where USER_ID= '" + txt_Check_Name.Text + "'";
                    DataSet ds = db.reDs(sqlstr);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        insertSQLScan();
                        this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    }
                    else
                    {
                        MessageBox.Show("員工編號輸入錯誤!");
                    }                                                            
                    //Close();
                }
                else
                {
                    MessageBox.Show("確認者欄位不可以空白!!");
                }

            }
            catch
            {
                DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
        }
        public void insertSQLScan() //建立資料庫
        {
            try
            {
                string insSql;
                insSql = "";

                insSql = @"INSERT INTO E_SOP_Product_Information_Check_Table (Record_Time,Model,Name,Wip_No,Process) VALUES("
                                       + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                       + "'" + MsgModel + "',"
                                       + "N'" + txt_Check_Name.Text.Trim() + "',"
                                       + "N'" + MsgWip_No + "',"
                                       + "N'" + MsgRoute + "')";

                if (db.Exsql(insSql) == true)
                {
                    return;
                }
                else
                {
                    MessageBox.Show("Windows資料庫上傳失敗");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Windows資料庫上傳失敗時發生錯誤：{ex.Message}");
                MessageBox.Show("Windows資料庫上傳失敗");
            }
        }

        private void txt_Check_Name_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                radButton1_Click(sender, e);
            }
        }
    }
}
