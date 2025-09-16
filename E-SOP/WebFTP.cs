using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using System.Data;


namespace E_SOP
{
    public class WebFTP
    {
        public static string FTP_Server;
        
        public static string FTP_User;
        public static string FTP_Password ;
        //登入
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
        string lpszUsername,
        string lpszDomain,
        string lpszPassword,
        int dwLogonType,
        int dwLogonProvider,
        ref IntPtr phToken);
        int totalSize;
        int position;
        const int BUFFER_SIZE = 4096;
        byte[] buffer;
        Stream stream;
        //登出
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);

        /// <summary>
        /// 取的檔案大小 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static long GetFileSize(string filename , string Path , string NASserver_master)
        {
            FtpWebRequest reqFTP;
            long fileSize = 0;
            try
            {
                reqFTP = (FtpWebRequest)FtpWebRequest.Create("ftp://" + "172.22.2.82" + "/" + "ESOP/PDF/" + filename);
                reqFTP.Method = WebRequestMethods.Ftp.GetFileSize;
                reqFTP.UseBinary = true;
                reqFTP.KeepAlive = true;
                reqFTP.Credentials = new NetworkCredential(FTP_User, FTP_Password);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();
                fileSize = response.ContentLength;
                ftpStream.Close();
                response.Close();
            }
            catch(Exception ee)
            {
                
            }
            return fileSize;
        }
        /// <summary>
        /// 檔案上傳
        /// </summary>
        /// <param name="updatefilename"></param>
        /// <param name="filename"></param>
        /// <param name="Path"></param>
        /// <param name="NASserver_master"></param>
        /// <param name="Bar1"></param>
        public static bool UploadFile(string updatefilename,string filename, string file, string NASserver_master, ProgressBar Bar1)
        {
            bool result = true;
            

            try
            {
                Getftp(NASserver_master);

               
                FileInfo finfo = new FileInfo(updatefilename);

                FtpWebResponse response = null;
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + FTP_Server + "/" + file + "/" + filename );
                request.KeepAlive = true;
                request.UseBinary = true;
                request.Credentials = new NetworkCredential(FTP_User, FTP_Password);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.ContentLength = finfo.Length;//指定上傳文件的大小
                response = request.GetResponse() as FtpWebResponse;
                int buffLength = 2048;
                byte[] buffer = new byte[buffLength];
                int contentLen;
                FileStream fs = File.OpenRead(updatefilename);
                Stream ftpstream = request.GetRequestStream();
                contentLen = fs.Read(buffer, 0, buffer.Length);
                int allbye = (int)finfo.Length;
                Form.CheckForIllegalCrossThreadCalls = false;
                Bar1.Maximum = allbye;//設定進度條長度
                int startbye = 0;
                while (contentLen != 0)
                {
                    startbye = contentLen + startbye;
                    ftpstream.Write(buffer, 0, contentLen);
                    //更新進度
                    if (Bar1 != null)
                    {
                        Bar1.Value += contentLen;//更新進度條
                    }
                    contentLen = fs.Read(buffer, 0, buffLength);
                }
                fs.Close();
                ftpstream.Close();
                response.Close();
           

            }
            catch (Exception ftp)
            {
                result = false;
                MessageBox.Show(ftp.Message);
            }

            return result;
        }
        /// <summary>
        /// 下載
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        public static bool Downloadfile(string DLfilename, string LocalPath, string NASserver_master, string FilePath)
        {
            bool result = true;
            
            try
            {

                Getftp(NASserver_master);
                string tempStoragePath = LocalPath ; ;

                //FtpWebRequest
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create("ftp://" + FTP_Server + "/" + FilePath +"/"+ DLfilename);
                NetworkCredential ftpCredential = new NetworkCredential(FTP_User, FTP_Password);
                ftpRequest.Credentials = ftpCredential;
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                //FtpWebResponse
                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                //Get Stream From FtpWebResponse
                Stream ftpStream = ftpResponse.GetResponseStream();
                using (FileStream fileStream = new FileStream(tempStoragePath, FileMode.Create))
                {
                    int bufferSize = 2048;
                    int readCount;
                    byte[] buffer = new byte[bufferSize];

                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                    int allbye = (int)fileStream.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;
                 
                    while (readCount > 0)
                    {
                        
                        fileStream.Write(buffer, 0, readCount);
                        readCount = ftpStream.Read(buffer, 0, bufferSize);
                    }
                }
                ftpStream.Close();
                ftpResponse.Close();

                return result;

            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                result = false;
                return result;
            }
            
        }
        public static bool SAMbar (string updatefilename, string filename, string file, string NASserver_master, ProgressBar Bar1)
        {
            bool result = true;
            try
            {
                string MachineName = NASserver_master;
                string UserName = FTP_User;
                string Pw = FTP_Password;
                string IPath = String.Format(@"\\{0}\abc", MachineName);
                const int LOGON32_PROVIDER_DEFAULT = 0;
                const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
                IntPtr tokenHandle = new IntPtr(0);
                tokenHandle = IntPtr.Zero;

                bool returnValue = LogonUser(UserName, MachineName, Pw,
                LOGON32_LOGON_NEW_CREDENTIALS,
                LOGON32_PROVIDER_DEFAULT,
                ref tokenHandle);

                WindowsIdentity w = new WindowsIdentity(tokenHandle);
                w.Impersonate();
                if (false == returnValue)
                {
                    result = false;
                }
                FileInfo finfo = new FileInfo(filename);
                updatefilename = Path.GetFileNameWithoutExtension(filename) + Path.GetExtension(filename);
                DirectoryInfo dir = new DirectoryInfo(file);

                File.Copy(filename, @"\\" + MachineName + "\\" + file + updatefilename, true);
                if(File.Exists(@"\\" + MachineName + "\\" + file + updatefilename) == true)
                {
                    result = true;
                }
            }
            catch (Exception ee)
            {
                result = false;
            }
            return result;
        }

        public static void Getftp(string Ftp_Server_name)//ftp資訊
        {


            string sqlCmd = "SELECT [Ftp_Server_OA_Ip],[Ftp_Username],[Ftp_Password],[Ftp_Server_name] FROM [iFactory].[i_Program].[i_Program_FtpServer_Table] where [Ftp_Server_name] ='" + Ftp_Server_name + "' ";
            DataSet ds = E_SOP.db.reDs(sqlCmd);
            if (ds.Tables[0].Rows.Count != 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    FTP_Server = ds.Tables[0].Rows[0]["Ftp_Server_OA_Ip"].ToString().Trim();
                    FTP_User = ds.Tables[0].Rows[0]["Ftp_Username"].ToString().Trim();
                    FTP_Password = ds.Tables[0].Rows[0]["Ftp_Password"].ToString().Trim();


                }
            }




        }
    }
}
