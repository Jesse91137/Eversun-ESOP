using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E_SOP
{
    public partial class ViewForm : Form
    {
        DataTable dt = new DataTable();
        string _engsr;
        #region 批次出庫來的參數
        private string strEngsr;                //工程編號
        private string strSpec;                 //規格
        private string strNotice;              //注意事項
        private string strItem;                 //料號
        private string strQuantity;         //數量
        private DataTable tables;
        public static int i = 0;
        public string Engsr
        {
            set { strEngsr = value; }
        }
        public string Spec
        {
            set { strSpec = value; }
        }
        public string Notice
        {
            set { strNotice = value; }
        }
        public string Item
        {
            set { strItem = value; }
        }
        public string Quantity
        {
            set { strQuantity = value; }
        }
        public DataTable Tables
        {
            set { tables = value; }
        }
        public void setValue()
        {
            _engsr = strEngsr;
            dt = tables;
        }
        #endregion
        public ViewForm()
        {
            InitializeComponent();
        }

        private void ViewForm_Load(object sender, EventArgs e)
        {
            ShowRowData(0);
        }
        public void ShowRowData(int i)
        {            
            //機種名稱            
            lab_engsr.Text = strEngsr;
            //找到幾筆
            lab_count.Text = dt.Rows.Count.ToString();
            //注意事項
            lab_notice.Text = dt.Rows[i]["注意事項"].ToString();
            //料號
            lab_item.Text = dt.Rows[i]["料號"].ToString();
            //規格
            lab_spec.Text = dt.Rows[i]["規格"].ToString();
            //開啟圖片
            //\\192.168.4.11\全廠共用\31-SMD拋料率及不易維修紀錄\01-不易維修照片\
            string pic_path = @"\\192.168.4.11\全廠共用\31-SMD拋料率及不易維修紀錄\01-不易維修照片\";
            FileStream fs = File.OpenRead(pic_path+ dt.Rows[i]["料號"].ToString()+".jpg");
            this.pictureBox1.Image = Image.FromStream(fs);
            fs.Close();
        }

        private void btn_down_Click(object sender, EventArgs e)
        {
            string btn = ((Button)sender).Name;
            switch (btn)
            {
                case "btn_down":
                    i++;
                    if (i > dt.Rows.Count - 1) i--;
                    ShowRowData(i);
                    break;
                case "btn_up":
                    i--;
                    if (i <= 0) i = 0;
                    ShowRowData(i);
                    break;
                default:
                    break;
            }
        }
    }
}
