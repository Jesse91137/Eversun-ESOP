using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace E_SOP
{
    /// <summary>
    /// 顯示資料的表單，支援批次出庫資料瀏覽與圖片顯示
    /// </summary>
    public partial class ViewForm : Form
    {
        /// <summary>
        /// 用於儲存資料表
        /// </summary>
        DataTable dt = new DataTable();

        /// <summary>
        /// 工程編號(內部使用)
        /// </summary>
        string _engsr;

        #region 批次出庫來的參數

        /// <summary>
        /// 工程編號
        /// </summary>
        private string strEngsr;

        /// <summary>
        /// 規格
        /// </summary>
        private string strSpec;

        /// <summary>
        /// 注意事項
        /// </summary>
        private string strNotice;

        /// <summary>
        /// 料號
        /// </summary>
        private string strItem;

        /// <summary>
        /// 數量
        /// </summary>
        private string strQuantity;

        /// <summary>
        /// 批次資料表
        /// </summary>
        private DataTable tables;

        /// <summary>
        /// 當前資料索引
        /// </summary>
        public static int i = 0;

        /// <summary>
        /// 設定工程編號
        /// </summary>
        public string Engsr
        {
            set { strEngsr = value; }
        }

        /// <summary>
        /// 設定規格
        /// </summary>
        public string Spec
        {
            set { strSpec = value; }
        }

        /// <summary>
        /// 設定注意事項
        /// </summary>
        public string Notice
        {
            set { strNotice = value; }
        }

        /// <summary>
        /// 設定料號
        /// </summary>
        public string Item
        {
            set { strItem = value; }
        }

        /// <summary>
        /// 設定數量
        /// </summary>
        public string Quantity
        {
            set { strQuantity = value; }
        }

        /// <summary>
        /// 設定資料表
        /// </summary>
        public DataTable Tables
        {
            set { tables = value; }
        }

        /// <summary>
        /// 將批次出庫參數賦值到內部欄位
        /// </summary>
        public void setValue()
        {
            _engsr = strEngsr;
            dt = tables;
        }
        #endregion

        /// <summary>
        /// 建構子，初始化元件
        /// </summary>
        public ViewForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 表單載入事件，顯示第一筆資料
        /// </summary>
        private void ViewForm_Load(object sender, EventArgs e)
        {
            ShowRowData(0);
        }

        /// <summary>
        /// 顯示指定索引的資料行內容
        /// </summary>
        /// <param name="i">資料行索引</param>
        public void ShowRowData(int i)
        {
            // 機種名稱
            lab_engsr.Text = strEngsr;
            // 找到幾筆
            lab_count.Text = dt.Rows.Count.ToString();
            // 注意事項
            lab_notice.Text = dt.Rows[i]["注意事項"].ToString();
            // 料號
            lab_item.Text = dt.Rows[i]["料號"].ToString();
            // 規格
            lab_spec.Text = dt.Rows[i]["規格"].ToString();
            // 開啟圖片
            // \\192.168.4.11\全廠共用\31-SMD拋料率及不易維修紀錄\01-不易維修照片\
            string pic_path = @"\\192.168.4.11\全廠共用\31-SMD拋料率及不易維修紀錄\01-不易維修照片\";
            FileStream fs = File.OpenRead(pic_path + dt.Rows[i]["料號"].ToString() + ".jpg");
            this.pictureBox1.Image = Image.FromStream(fs);
            fs.Close();
        }

        /// <summary>
        /// 按鈕點擊事件，支援上下切換資料
        /// </summary>
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
