using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WeighingSystem
{
    public class SingleFtpClient
    {
        #region Properties

        public string Host { get; set; }

        public string ID { get; set; }

        public string Password { get; set; }

        private object _syncObject = new object();

        private NetworkCredential _credential = null;

        private static readonly object _syncRoot = new object();

        public static SingleFtpClient Instance
        {
            get
            {
                lock (_syncRoot)
                {
                    if (_instance == null)
                        _instance = new SingleFtpClient();

                    return _instance;
                }
            }
        }
        private static SingleFtpClient _instance = null;

        #endregion

        #region Constructor

        public SingleFtpClient()
        {

        }

        #endregion

        #region Events

        #endregion

        #region Methods

        public void Initialize(string host, string id, string pwd)
        {
            Host = host;
            ID = id;
            Password = pwd;

            _credential = new NetworkCredential(id, pwd);
        }

        public bool FtpDirectioryCheck(string path)
        {
            Exception ex = null;
            bool existRet = IsExtistFtpDirectory(path, out ex);

            // ftp directory 검사시 오류발생
            if (ex != null)
                return false;

            if (existRet)
            {
                // ftp directory 존재함
                return true;
            }

            return MakeFtpDirectory(path);
        }

        public bool FtpUpload(string ipAddress, string subPass, string fileName)
        {
            if (!File.Exists(fileName))
            {
                Log.AddLog(string.Concat("FtpUpload - ", fileName, " 파일이 존재하지 않습니다."));
                return false;
            }

            WebClient client = new WebClient();
            client.Credentials = _credential;

            try
            {
                string uploadPosition = string.Empty;
                if (subPass == "")
                    uploadPosition = string.Format("ftp://{0}/", this.Host);
                else
                    uploadPosition = string.Format("ftp://{0}/{1}/", this.Host, subPass);
                    
                // URI (Uniform Resource Identifier)
                string uriString = string.Concat(uploadPosition, Path.GetFileName(fileName));
                Uri uri = new Uri(uriString);

                // Log.AddLog(uriString);

                client.UploadFile(uri, fileName);
                // _client.UploadFileAsync(uri, curr_file);   

                return true;
            }
            catch (Exception ex)
            {
                Log.AddLog(string.Concat("FtpUpload exception - ", ex.Message));
                return false;
            }
            finally
            {
                client.Dispose();
            }
        }

        public bool CheckFileExist(string uriString, out Exception e)
        {
            e = null;

            FtpWebRequest ftpRequest = null;
            FtpWebResponse ftpResponse = null;

            bool isExists = true;

            try
            {
                ftpRequest = (FtpWebRequest)WebRequest.Create(uriString);
                ftpRequest.Credentials = _credential;
                ftpRequest.Method = WebRequestMethods.Ftp.GetFileSize;

                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                ftpResponse.Close();

                ftpRequest = null;

            }
            catch (Exception ex)
            {
                e = ex;
                isExists = false;
            }
            finally
            {
                if (ftpResponse != null)
                    ftpResponse.Close();
            }

            return isExists;
        }

        #endregion

        #region Implementations

        private bool IsExtistFtpDirectory(string currentDirectory, out Exception e)
        {
            e = null;

            string ftpFullPath = string.Format("ftp://{0}{1}", this.Host, GetParentDirectory(currentDirectory));

            string currDir = string.Empty;
            int location = currentDirectory.LastIndexOf('/');
            if (location == -1)
                currDir = currentDirectory;
            else
                currDir = currentDirectory.Substring(location + 1);

            FtpWebRequest ftpWebRequest = WebRequest.Create(new Uri(ftpFullPath)) as FtpWebRequest;
            ftpWebRequest.Credentials = _credential;
            ftpWebRequest.UseBinary = true;
            ftpWebRequest.UsePassive = true;
            ftpWebRequest.Timeout = 10000;
            ftpWebRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;

            FtpWebResponse response = null;
            string data = string.Empty;

            try
            {
                response = ftpWebRequest.GetResponse() as FtpWebResponse;
                if (response != null)
                {
                    StreamReader streamReader = new StreamReader(response.GetResponseStream(), Encoding.Default);
                    data = streamReader.ReadToEnd();
                }

                string[] directories = data.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                //char[] splitChar = new char[5] { '<', 'D', 'I', 'R', '>' };
                char[] splitChar = new char[2] { '<', '>' };
                foreach (string dir in directories)
                {
                    string[] dirInfo = dir.Split(splitChar, StringSplitOptions.RemoveEmptyEntries);
                    if (string.Equals(dirInfo[2].Trim().ToLower(), currDir.ToLower()))
                        return true;
                }

                return false;

            }
            catch (Exception ex)
            {
                e = ex;
                Log.AddLog(string.Concat("IsExtistFtpDirectory exception - ", ex.Message));

                return false;
            }
            finally
            {
                if (response != null)
                    response.Close();
            }
        }

        /// <summary>
        /// FTP에 해당 디렉터리를 만든다.
        /// </summary>
        /// <param name="dirPath"></param>
        /// <returns></returns>
        private bool MakeFtpDirectory(string dirPath)
        {
            string uri = string.Format("ftp://{0}/{1}", this.Host, dirPath);

            FtpWebRequest ftpWebRequest = WebRequest.Create(new Uri(uri)) as FtpWebRequest;
            ftpWebRequest.Credentials = _credential;
            ftpWebRequest.KeepAlive = false;
            ftpWebRequest.Method = System.Net.WebRequestMethods.Ftp.MakeDirectory;

            FtpWebResponse response = null;
            string data = string.Empty;

            try
            {
                response = ftpWebRequest.GetResponse() as FtpWebResponse;
                if (response != null)
                {
                    StreamReader streamReader = new StreamReader(response.GetResponseStream(), Encoding.Default);
                    data = streamReader.ReadToEnd();
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.AddLog(string.Concat("MakeFtpDirectory exception - ", ex.Message));
                return false;
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private string GetParentDirectory(string currentDirectory)
        {
            string[] directorys = currentDirectory.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);

            string parentDirectory = string.Empty;
            for (int i = 0; i < directorys.Length - 1; i++)
                parentDirectory += "/" + directorys[i];

            return parentDirectory;
        }

        #endregion

    }
}
