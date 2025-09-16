using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace E_SOP
{
    /// <summary>
    /// 提供資料庫連線與常用操作的靜態方法。
    /// </summary>
    static class db // 定義靜態類別 db
    {
        /// <summary>
        /// 取得並開啟一個 SQL Server 資料庫連線。
        /// </summary>
        /// <returns>已開啟的 <see cref="SqlConnection"/> 物件。</returns>
        /// <remarks>
        /// 連線字串從組態檔的 "cnstr" 取得。
        /// 若連線已開啟則先關閉再重新開啟。
        /// </remarks>
        public static SqlConnection GetCon() // 取得資料庫連線的方法
        {
            // 從組態檔取得連線字串
            string cnstr = ConfigurationManager.ConnectionStrings["cnstr"]?.ConnectionString;

            // 建立 SQL 連線物件
            SqlConnection icn = new SqlConnection();
            // 設定連線字串
            icn.ConnectionString = cnstr;
            // 如果連線已開啟則先關閉
            if (icn.State == ConnectionState.Open) icn.Close();
            // 開啟連線
            icn.Open();

            // 回傳已開啟的連線
            return icn;
        }

        /// <summary>
        /// 執行指定的 SQL 指令（非查詢），例如 INSERT、UPDATE 或 DELETE。
        /// </summary>
        /// <param name="cmdtxt">要執行的 SQL 指令字串。</param>
        /// <returns>若執行成功則回傳 true，否則回傳 false。</returns>
        /// <remarks>
        /// 此方法會自動建立並開啟資料庫連線，執行完畢後自動關閉連線。
        /// 若執行過程發生例外，會顯示錯誤訊息視窗。
        /// </remarks>
        public static bool Exsql(string cmdtxt) // 執行非查詢 SQL 指令的方法
        {
            // 取得資料庫連線
            SqlConnection con = db.GetCon();
            // 建立 SQL 指令物件
            SqlCommand cmd = new SqlCommand(cmdtxt, con);
            try
            {
                // 執行 SQL 指令
                cmd.ExecuteNonQuery();
                // 執行成功回傳 true
                return true;
            }
            catch (Exception e)
            {
                // 發生例外顯示錯誤訊息
                MessageBox.Show(e.ToString());
                // 回傳 false
                return false;
            }
            finally
            {
                // 關閉並釋放連線
                con.Dispose();
                con.Close();
            }
        }

        /// <summary>
        /// 執行 SQL 查詢並回傳查詢結果的 DataSet。
        /// </summary>
        /// <param name="cmdtxt">要執行的 SQL 查詢語句。</param>
        /// <returns>包含查詢結果的 <see cref="DataSet"/> 物件。</returns>
        public static DataSet reDs(string cmdtxt) // 執行查詢並回傳 DataSet 的方法
        {
            // 取得資料庫連線
            SqlConnection con = db.GetCon();
            // 建立資料配接器
            SqlDataAdapter da = new SqlDataAdapter(cmdtxt, con);
            // 建立 DataSet 物件
            DataSet ds = new DataSet();
            // 填充查詢結果
            da.Fill(ds);
            // 回傳查詢結果
            return ds;
        }

        /// <summary>
        /// 執行帶參數的 SQL 查詢，並回傳查詢結果的 DataSet。
        /// </summary>
        /// <param name="sqlCmd">要執行的 SQL 查詢語句。</param>
        /// <param name="parameters">查詢所需的參數集合，Key 為參數名稱，Value 為參數值。</param>
        /// <returns>查詢結果的 <see cref="DataSet"/> 物件。</returns>
        /// <exception cref="Exception">查詢過程中發生的任何例外。</exception>
        public static DataSet reDsWithParams(string sqlCmd, Dictionary<string, object> parameters) // 執行帶參數查詢的方法
        {
            // 取得資料庫連線
            SqlConnection con = db.GetCon();
            try
            {
                // 使用 using 管理 SQL 指令物件
                using (SqlCommand cmd = new SqlCommand(sqlCmd, con))
                {
                    // 加入所有參數
                    foreach (var param in parameters)
                    {
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                    }
                    // 建立資料配接器
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    // 建立 DataSet 物件
                    DataSet ds = new DataSet();
                    // 填充查詢結果
                    da.Fill(ds);
                    // 回傳查詢結果
                    return ds;
                }
            }
            catch (Exception ex)
            {
                // 查詢發生錯誤時輸出訊息
                Console.WriteLine($"SQL查詢時發生錯誤：{ex.Message}");
                // 丟出例外
                throw;
            }
        }

        /// <summary>
        /// 執行 SQL 語法並返回單一值。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語法。</param>
        /// <param name="parameters">參數字典，包含參數名稱和值。</param>
        /// <returns>查詢結果的第一行第一列的值，如果沒有結果則返回 null。</returns>
        public static object ExecuteScalar(string sql, Dictionary<string, object> parameters) // 執行查詢並回傳單一值的方法
        {
            // 宣告結果變數
            object result = null;
            // 使用 using 確保連線正確釋放
            using (SqlConnection connection = db.GetCon())
            {
                // 建立 SQL 指令物件
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    // 設定指令型態為文字
                    command.CommandType = CommandType.Text;
                    // 如果有參數則加入參數
                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            // 加入參數，若值為 null 則用 DBNull.Value
                            command.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
                        }
                    }
                    try
                    {
                        // 執行查詢並取得結果
                        result = command.ExecuteScalar();
                    }
                    catch (SqlException ex)
                    {
                        // 查詢失敗時輸出錯誤訊息
                        Console.WriteLine("資料庫查詢失敗：" + ex.Message);
                        // 丟出例外
                        throw;
                    }
                }
            }
            // 回傳查詢結果
            return result;
        }

        /// <summary>
        /// 執行 SQL 查詢並回傳單一結果（字串）。
        /// </summary>
        /// <param name="str_select">要執行的 SQL 查詢語句。</param>
        /// <returns>
        /// 查詢結果的第一行第一列資料（字串型別）。
        /// 若查詢過程發生例外，則回傳例外訊息字串。
        /// </returns>
        public static string scalDs(string str_select) // 執行查詢並回傳單一字串的方法
        {
            // 建立 SQL 連線
            SqlConnection con = db.GetCon();
            // 建立 SQL 指令
            SqlCommand com_select = new SqlCommand(str_select, con);
            try
            {
                // 開啟資料庫連線
                con.Open();
                // 執行查詢並取得第一行第一列的資料，轉為字串
                str_select = Convert.ToString(com_select.ExecuteScalar());
            }
            catch (Exception ex)
            {
                // 發生例外時關閉連線
                con.Close();
                // 回傳例外訊息字串
                return Convert.ToString(ex);
            }
            finally
            {
                // 最後關閉連線
                con.Close();
            }
            // 回傳查詢結果
            return str_select;
        }
    }
}
