namespace E_SOP
{
    /// <summary>
    /// MyItem 類別用於表示下拉選單的項目，包含顯示文字與對應值。
    /// </summary>
    internal class MyItem
    {
        /// <summary>
        /// 顯示於下拉選單的文字。
        /// </summary>
        public string text;

        /// <summary>
        /// 對應的值（如資料庫 ID）。
        /// </summary>
        public string value;

        /// <summary>
        /// 建構函式，初始化 MyItem 物件。
        /// </summary>
        /// <param name="text">顯示文字。</param>
        /// <param name="value">對應值。</param>
        public MyItem(string text, string value)
        {
            this.text = text;   // 設定顯示文字
            this.value = value; // 設定對應值
        }

        /// <summary>
        /// 覆寫 ToString 方法，回傳顯示文字。
        /// </summary>
        /// <returns>顯示文字。</returns>
        public override string ToString()
        {
            return text; // 回傳顯示文字
        }
    }
}
