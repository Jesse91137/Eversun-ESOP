using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Net;

namespace E_SOP
{
    class FTPClient
    {
        private string strMsg;
        private string strReply;
        private int iReplyCode;
        private Socket socketControl;
        private TransferType trType;
        private static int BLOCK_SIZE = 512;
        Byte[] buffer = new Byte[BLOCK_SIZE];
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
        private void CloseSocketConnect()
        {
            if (socketControl != null)
            {
                socketControl.Close();
                socketControl = null;
            }
            bConnected = false;
        }
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
        private void SendCommand1(String strCommand)
        {
           // Encoding e = Encoding.GetEncoding("utf-8");
            Byte[] cmdBytes = Encoding.UTF8.GetBytes((strCommand + "\r\n").ToCharArray());
            socketControl.Send(cmdBytes, cmdBytes.Length, 0);
            ReadReply();
        }
    }
}
