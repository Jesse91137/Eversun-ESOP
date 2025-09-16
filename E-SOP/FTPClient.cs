using System;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Net;

namespace E_SOP
{
    class FTPClient
    {
        /// <summary>
        /// 記錄從控制連線接收的原始訊息串流 (包含可能的多行回應)。
        /// 使用 ReadLine 與 ReadReply 來累積並解析此緩衝訊息。
        /// </summary>
        private string strMsg;

        /// <summary>
        /// FTP 伺服器最後一筆回應的單行文字內容。
        /// 典型格式為 "XXX message"，其中 XXX 為回應碼 (3 位數)。
        /// ReadReply 會更新此欄位以供後續檢查與例外處理使用。
        /// </summary>
        private string strReply;

        /// <summary>
        /// 解析自 strReply 的數值回應碼 (前三個字元)。
        /// 使用於檢查 FTP 指令是否成功或決定後續處理邏輯。
        /// </summary>
        private int iReplyCode;

        /// <summary>
        /// 控制連線的 Socket，負責傳送指令與接收控制通道回應。
        /// 此 Socket 以 TCP stream 模式與 FTP 伺服器建立連線。
        /// </summary>
        private Socket socketControl;

        /// <summary>
        /// 記錄目前的傳輸類型 (Binary 或 ASCII)。
        /// 由 SetTransferType 設定，影響檔案上傳/下載時的處理行為。
        /// </summary>
        private TransferType trType;

        /// <summary>
        /// 內部使用的資料區塊大小 (位元組)。
        /// 用於檔案傳輸時的讀/寫緩衝長度；預設為 512 bytes。
        /// 可視需求調整以平衡記憶體與效能。
        /// </summary>
        private static int BLOCK_SIZE = 512;

        /// <summary>
        /// 傳輸緩衝區陣列，大小等於 BLOCK_SIZE。
        /// 於上傳 (Put) 與下載 (Get) 等資料通道操作時使用。
        /// </summary>
        Byte[] buffer = new Byte[BLOCK_SIZE];

        /// <summary>
        /// 用來將位元組與字串相互轉換的 Encoding 實例。
        /// 名稱為 ASCII，但實際設定為 UTF-8，以配合伺服器訊息處理與泛用性。
        /// 若需精確 ASCII 行為，可考慮改為 Encoding.ASCII。
        /// </summary>
        Encoding ASCII = Encoding.UTF8;
        //--------------------------------------------------------------
        public FTPClient()
        {
            strRemoteHost = "";
            strRemotePath = "";
            strRemoteUser = "";
            strRemotePass = "";
            strRemotePort = 21;
            bConnected = false;
        }
        public FTPClient(string remoteHost, string remotePath, string remoteUser, string remotePass, int remotePort)
        {
            strRemoteHost = remoteHost;
            strRemotePath = remotePath;
            strRemoteUser = remoteUser;
            strRemotePass = remotePass;
            strRemotePort = remotePort;
            Connect();
        }
        /// FTP Server IP Address
        private string strRemoteHost;
        public string RemoteHost
        {
            get
            {
                return strRemoteHost;
            }
            set
            {
                strRemoteHost = value;
            }
        }
        // FTP Server Port
        private int strRemotePort;
        public int RemotePort
        {
            get
            {
                return strRemotePort;
            }
            set
            {
                strRemotePort = value;
            }
        }
        // Server Folder
        private string strRemotePath;
        public string RemotePath
        {
            get
            {
                return strRemotePath;
            }
            set
            {
                strRemotePath = value;
            }
        }
        // User帳號
        private string strRemoteUser;
        public string RemoteUser
        {
            set
            {
                strRemoteUser = value;
            }
        }
        //Password
        private string strRemotePass;
        public string RemotePass
        {
            set
            {
                strRemotePass = value;
            }
        }
        // 是否已登入
        private Boolean bConnected;
        public bool Connected
        {
            get
            {
                return bConnected;
            }
        }
        // 建立連線
        public void Connect()
        {
            socketControl = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IPEndPoint ep = new IPEndPoint(IPAddress.Parse(RemoteHost), strRemotePort);
            // 連線
            try
            {
                socketControl.Connect(ep);
            }
            catch (Exception)
            {
                throw new IOException("Couldn't connect to remote server");
            }
            // 取得回應
            ReadReply();
            if (iReplyCode != 220)
            {
                DisConnect();
                throw new IOException(strReply.Substring(4));
            }
            // 登入
            SendCommand1("USER " + strRemoteUser);
            if (!(iReplyCode == 331 || iReplyCode == 230))
            {
                CloseSocketConnect();//如果有錯誤就關閉連線
                throw new IOException(strReply.Substring(4));
            }
            if (iReplyCode != 230)
            {
                SendCommand1("PASS " + strRemotePass);
                if (!(iReplyCode == 230 || iReplyCode == 202))
                {
                    CloseSocketConnect();
                    throw new IOException(strReply.Substring(4));
                }
            }
            bConnected = true;
            // 切換到所選的目錄
            ChDir(strRemotePath);
        }
        //關閉連線
        public void DisConnect()
        {
            if (socketControl != null)
            {
                SendCommand1("QUIT");
            }
            CloseSocketConnect();
        }
        //傳輸模式
        public enum TransferType { Binary, ASCII };
        //設定傳輸模式
        public void SetTransferType(TransferType ttType)
        {
            if (ttType == TransferType.Binary)
            {
                SendCommand1("TYPE I");//binary
            }
            else
            {
                SendCommand1("TYPE A");//ASCII
            }
            if (iReplyCode != 200)
            {
                throw new IOException(strReply.Substring(4));
            }
            else
            {
                trType = ttType;
            }
        }
        /// 取得傳輸模式
        public TransferType GetTransferType()
        {
            return trType;
        }
        // 取得文件大小
        private long GetFileSize(string strFileName)
        {
            if (!bConnected)
            {
                Connect();
            }
            SendCommand1("SIZE " + Path.GetFileName(strFileName));
            long lSize = 0;
            if (iReplyCode == 213)
            {
                lSize = Int64.Parse(strReply.Substring(4));
            }
            else
            {
                throw new IOException(strReply.Substring(4));
            }
            return lSize;
        }
        // 要刪除的文件名稱
        public void Delete(string strFileName)
        {
            if (!bConnected)
            {
                Connect();
            }
            SendCommand1("DELE " + strFileName);
            if (iReplyCode != 250)
            {
                throw new IOException(strReply.Substring(4));
            }
        }
        //Rename(如果要改的名稱已存在將會覆蓋)
        //舊文件名
        //新文件名
        public void Rename(string strOldFileName, string strNewFileName)
        {
            if (!bConnected)
            {
                Connect();
            }
            SendCommand1("RNFR " + strOldFileName);
            if (iReplyCode != 350)
            {
                throw new IOException(strReply.Substring(4));
            }

            SendCommand1("RNTO " + strNewFileName);
            if (iReplyCode != 250)
            {
                throw new IOException(strReply.Substring(4));
            }
        }
        //下載一個文件
        //要下載的文件名稱
        //本機的目錄(不得以\结束)
        //本機的檔名
        public void Get(string strRemoteFileName, string strFolder, string strLocalFileName)
        {
            if (!bConnected)
            {
                Connect();
            }
            SetTransferType(TransferType.Binary);
            if (strLocalFileName.Equals(""))
            {
                strLocalFileName = strRemoteFileName;
            }
            if (!File.Exists(strLocalFileName))
            {
                Stream st = File.Create(strLocalFileName);
                st.Close();
            }
            FileStream output = new
            FileStream(strFolder + "\\" + strLocalFileName, FileMode.Create);
            Socket socketData = CreateDataSocket();
            SendCommand1("RETR " + strRemoteFileName);
            if (!(iReplyCode == 150 || iReplyCode == 125
            || iReplyCode == 226 || iReplyCode == 250))
            {
                throw new IOException(strReply.Substring(4));
            }
            while (true)
            {
                int iBytes = socketData.Receive(buffer, buffer.Length, 0);
                output.Write(buffer, 0, iBytes);
                if (iBytes <= 0)
                {
                    break;
                }
            }
            output.Close();
            if (socketData.Connected)
            {
                socketData.Close();
            }
            if (!(iReplyCode == 226 || iReplyCode == 250))
            {
                ReadReply();
                if (!(iReplyCode == 226 || iReplyCode == 250))
                {
                    throw new IOException(strReply.Substring(4));
                }
            }
        }
        //上傳一個文件
        //本機端的檔案名稱
        public void Put(string strFileName)
        {
            if (!bConnected)
            {
                Connect();
            }
            Socket socketData = CreateDataSocket();
            SendCommand1("STOR " + Path.GetFileName(strFileName));
            if (!(iReplyCode == 125 || iReplyCode == 150))
            {
                throw new IOException(strReply.Substring(4));
            }
            FileStream input = new
            FileStream(strFileName, FileMode.Open);
            int iBytes = 0;
            while ((iBytes = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                socketData.Send(buffer, iBytes, 0);
            }
            input.Close();
            if (socketData.Connected)
            {
                socketData.Close();
            }
            if (!(iReplyCode == 226 || iReplyCode == 250))
            {
                ReadReply();
                if (!(iReplyCode == 226 || iReplyCode == 250))
                {
                    throw new IOException(strReply.Substring(4));
                }
            }
        }
        //改變目錄
        //新的目錄名稱
        public void ChDir(string strDirName)
        {
            if (strDirName.Equals(".") || strDirName.Equals(""))
            {
                return;
            }
            if (!bConnected)
            {
                Connect();
            }
            SendCommand1("CWD " + strDirName);
            if (iReplyCode != 250)
            {
                throw new IOException(strReply.Substring(4));
            }
            this.strRemotePath = strDirName;
        }
        private void ReadReply()
        {
            strMsg = "";
            strReply = ReadLine();
            iReplyCode = Int32.Parse(strReply.Substring(0, 3));
        }
        /// Socket
        private Socket CreateDataSocket()
        {
            SendCommand1("PASV");
            if (iReplyCode != 227)
            {
                throw new IOException(strReply.Substring(4));
            }
            int index1 = strReply.IndexOf('(');
            int index2 = strReply.IndexOf(')');
            string ipData =
             strReply.Substring(index1 + 1, index2 - index1 - 1);
            int[] parts = new int[6];
            int len = ipData.Length;
            int partCount = 0;
            string buf = "";
            for (int i = 0; i < len && partCount <= 6; i++)
            {
                char ch = Char.Parse(ipData.Substring(i, 1));
                if (Char.IsDigit(ch))
                    buf += ch;
                else if (ch != ',')
                {
                    throw new IOException("Malformed PASV strReply: " +
                     strReply);
                }
                if (ch == ',' || i + 1 == len)
                {
                    try
                    {
                        parts[partCount++] = Int32.Parse(buf);
                        buf = "";
                    }
                    catch (Exception)
                    {
                        throw new IOException("Malformed PASV strReply: " +
                         strReply);
                    }
                }
            }
            string ipAddress = parts[0] + "." + parts[1] + "." +
             parts[2] + "." + parts[3];
            int port = (parts[4] << 8) + parts[5];
            Socket s = new
            Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IPEndPoint ep = new IPEndPoint(IPAddress.Parse(ipAddress), port);
            try
            {
                s.Connect(ep);
            }
            catch (Exception)
            {
                throw new IOException("Can't connect to remote server");
            }
            return s;
        }

                /// <summary>
                /// 關閉控制連線的 Socket 並重設連線狀態。
                /// </summary>
                /// <remarks>
                /// 如果 <c>socketControl</c> 不為 <c>null</c>，此方法會呼叫 <see cref="System.Net.Sockets.Socket.Close"/> 關閉底層連線，
                /// 並將 <c>socketControl</c> 設為 <c>null</c> 以釋放參考。無論是否有可關閉的 socket，
                /// 最後都會將 <c>bConnected</c> 設為 <c>false</c> 表示目前不再處於已連線狀態。
                /// 此方法不會重新拋出底層 Socket 的例外；呼叫端若需處理底層例外可自行包裝呼叫。
                /// </remarks>
                /// <exception cref="System.ObjectDisposedException">
                /// 當底層 Socket 已被處置且 Close 操作導致例外時可能拋出。
                /// </exception>
                /// <exception cref="System.Net.Sockets.SocketException">
                /// 當底層 Socket 在關閉時發生 I/O 錯誤時可能拋出。
                /// </exception>
                /// <seealso cref="DisConnect"/>
                private void CloseSocketConnect()
        {
            if (socketControl != null)
            {
                socketControl.Close();
                socketControl = null;
            }
            bConnected = false;
        }

        /// <summary>
        /// 以阻塞方式從 FTP 控制通道讀取並解析一個完整的回應行 (single-line reply)。
        /// </summary>
        /// <returns>
        /// 解析後的單行回應字串，格式通常為 "XXX message"（其中 XXX 為 3 位數回應碼）。
        /// 回傳字串會儲存在內部欄位 <c>strMsg</c> 並供呼叫端繼續處理（例如由 ReadReply 使用）。
        /// </returns>
        /// <remarks>
        /// - 方法會從 <c>socketControl</c> 讀取資料並將位元組依 <c>ASCII</c>（實作為 UTF-8）編碼轉成字串，累積到 <c>strMsg</c> 中；
        ///   當本次 Receive 傳回的位元數小於緩衝大小時視為本次接收已結束，接著以換行字元拆分累積字串並取得最後的完整行作為回傳值。
        /// - FTP 協定的多行回應在每行的前三位為回應碼，第四位若為連字符 '-' 表示尚未結束；本方法以第四位為空白字元 (' ') 判斷回應行是否結束，
        ///   若判斷為未結束則遞迴呼叫自身以繼續接收與解析後續資料。
        /// - 此方法會阻塞直到有足夠資料可解析出單行回應或發生底層 Socket/系統例外。呼叫端應處理可能的 <see cref="System.Net.Sockets.SocketException"/>、
        ///   <see cref="System.ObjectDisposedException"/> 與 <see cref="System.ArgumentOutOfRangeException"/> 等例外情況。
        /// - 注意遞迴機制：在極端或錯誤的回應格式下可能導致深度遞迴或例外，呼叫此方法的上下文應確保連線與伺服器回應為正確格式。
        /// </remarks>
        /// <exception cref="System.Net.Sockets.SocketException">當底層 Socket 接收資料時發生錯誤。</exception>
        /// <exception cref="System.ObjectDisposedException">當 <c>socketControl</c> 已關閉或被釋放而嘗試接收資料時。</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">當解析回應字串長度不足或格式不正確而存取字串索引超出範圍時。</exception>
        private string ReadLine()
        {
            while (true)
            {
                int iBytes = socketControl.Receive(buffer, buffer.Length, 0);
                strMsg += ASCII.GetString(buffer, 0, iBytes);
                if (iBytes < buffer.Length)
                {
                    break;
                }
            }
            char[] seperator = { '\n' };
            string[] mess = strMsg.Split(seperator);
            if (strMsg.Length > 2)
            {
                strMsg = mess[mess.Length - 2];
            }
            else
            {
                strMsg = mess[0];
            }
            if (!strMsg.Substring(3, 1).Equals(" "))
            {
                return ReadLine();
            }
            return strMsg;
        }
        /// <summary>
        /// 將 FTP 指令送至控制通道並立即讀取伺服器回應以更新內部回應狀態。
        /// </summary>
        /// <param name="strCommand">
        /// 要送出的 FTP 指令，不需包含結尾 CRLF。方法會自動在指令後附加 "\r\n" 並以 UTF-8 編碼傳送。
        /// </param>
        /// <exception cref="IOException">
        /// 當與遠端伺服器連線失敗、傳送或接收資料時發生 I/O 錯誤，本方法可能會導致 Socket 或 ReadReply 所拋出的例外被外層以 IOException 形式處理。
        /// </exception>
        /// <remarks>
        /// 傳送完成後會呼叫 ReadReply() 以同步更新 strReply 與 iReplyCode。此方法假設 socketControl 已建立並連線；若 socketControl 為 null 或未連線，呼叫端應先建立/檢查連線。
        /// 若需改變編碼行為，可考慮改為 Encoding.ASCII 以符合傳統 FTP 控制通道的 ASCII 規範。
        /// </remarks>
        private void SendCommand1(String strCommand)
        {
            // Encoding e = Encoding.GetEncoding("utf-8");
            Byte[] cmdBytes = Encoding.UTF8.GetBytes((strCommand + "\r\n").ToCharArray());
            socketControl.Send(cmdBytes, cmdBytes.Length, 0);
            ReadReply();
        }
    }
}
