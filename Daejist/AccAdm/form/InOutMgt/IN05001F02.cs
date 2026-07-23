using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using System.Net;
using ComLib;

namespace AccAdm
{
    public partial class IN05001F02 : DevExpress.XtraEditors.XtraForm
    {
        public IN05001F02()
        {
            InitializeComponent();
        }
        private readonly string[] AllowableFileTypes = { ".jpg", ".jpeg", ".png", ".bmp" };
        
        public string JUNPYO_ID = string.Empty;
        private void IN05001F02_Load(object sender, EventArgs e)
        {
            GetImages(JUNPYO_ID);
        }

        private void GetImages(string sJunpyoId)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.J_DATE ");
                strSql.AppendLine("   FROM MESURING A ");
                strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sJ_Date = string.Empty;
                if (dt.Rows.Count > 0)
                {
                    sJ_Date = dt.Rows[0]["J_DATE"]?.ToString();
                }

                string[] sJDateArr = sJ_Date.Split(' ');
                string sJDate = sJDateArr[0];
                string[] strArr = sJDate.Split('-');
                //string ftpPath = "ftp://192.168.0.202/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
                //string user = "Administrator";
                //string pw = "Admin12345";
                string ftpPath = "ftp://"+ComnEtcFunc.FTP_IP+"/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                string fileName = null;

                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                //FTP 이미지를 Byte[]로 파싱하여 저장할 Dictionary 객체 초기세팅
                Dictionary<string, Image> dicPicture = new Dictionary<string, Image>();
                dicPicture.Add("1_1", null);
                dicPicture.Add("1_2", null);
                dicPicture.Add("1_3", null);
                dicPicture.Add("2_1", null);
                dicPicture.Add("2_2", null);
                dicPicture.Add("2_3", null);

                Dictionary<string, Image> dicCopy = new Dictionary<string, Image>();
                foreach (KeyValuePair<string, Image> item in dicPicture)
                    dicCopy.Add(item.Key, null);

                string[] filesInDirectory = null;
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    string[] filesCopy = new string[filesInDirectory.Length];
                    //파일이름은 전표_(1, 2)_(1, 2, 3) 순으로 다시 파싱
                    for (int i = 0; i < filesInDirectory.Length; i++)
                    {
                        string[] sVal = filesInDirectory[i].Split('_');
                        if (sVal.Length < 5)
                        {
                            filesCopy[i] = string.Empty;
                        }
                        else
                        {
                            string[] sVal2 = sVal[4].Split('.');
                            filesCopy[i] = string.Format("{0}_{1}_{2}", sVal[2], sVal[3], sVal2[0]);
                        }
                    }

                    string[] names = new string[dicPicture.Count];
                    int[] iArrIdx = new int[names.Length];
                    foreach (KeyValuePair<string, Image> item in dicPicture)
                    {
                        //해당 파일 Index
                        int findIndex = Array.FindIndex(filesCopy, i => i == string.Format("{0}_{1}", sJunpyoId, item.Key));
                        if (findIndex >= 0)
                        {
                            fileName = filesInDirectory[findIndex];
                            dicCopy[item.Key] = DownloadFTPFile(string.Format(@"{0}\{1}", ftpPath, fileName), user, pw);
                        }

                        if (CheckImageNull(dicCopy))
                            continue;
                        else
                            break;
                    }
                }

                Pic1_1.Image = dicCopy["1_1"];
                Pic1_2.Image = dicCopy["1_2"];
                Pic1_3.Image = dicCopy["1_3"];
                Pic2_1.Image = dicCopy["2_1"];
                Pic2_2.Image = dicCopy["2_2"];
                Pic2_3.Image = dicCopy["2_3"];

                dicPicture = null;

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                return;
            }

            Cursor = Cursors.Default;
        }


        #region FTP 파일 다운로드하기 - DownloadFTPFile(sourceFileURI, targetFilePath, userID, password)

        /// <summary>
        /// FTP 파일 다운로드하기
        /// </summary>
        /// <param name="sourceFileURI">소스 파일 URI</param>
        /// <param name="targetFilePath">타겟 파일 경로</param>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <returns>처리 결과</returns>

        public Image DownloadFTPFile(string sourceFileURI, string user, string pw)
        {
            Image img = null;
            try
            {
                Uri sourceFileUri = new Uri(sourceFileURI);
                FtpWebRequest ftpWebRequest = WebRequest.Create(sourceFileUri) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(user, pw);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;

                Stream sourceStream = ftpWebResponse.GetResponseStream();
                img = Image.FromStream(sourceStream);
                sourceStream.Close();
            }
            catch (Exception ex)
            {
                return null;
            }

            return img;
        }

        #endregion

        private bool CheckImageNull(Dictionary<string, Image> dicParams)
        {
            bool bYN = false;
            foreach (KeyValuePair<string, Image> item in dicParams)
            {
                if (item.Value == null)
                {
                    bYN = true;
                    break;
                }
            }
            return bYN;
        }

        private void BtnUpload_Click(object sender, EventArgs e)
        {
            if (XtraMessageBox.Show("현재 이미지가 저장됩니다", "이미지 저장여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
            string sJunpyoId = JUNPYO_ID;

            if (Pic1_1.Image != null)
                dicImages.Add(string.Format("{0}_{1}", JUNPYO_ID, "1_1"), Pic1_1.Image);

            if (Pic1_2.Image != null)
                dicImages.Add(string.Format("{0}_{1}", JUNPYO_ID, "1_2"), Pic1_2.Image);

            if (Pic1_3.Image != null)
                dicImages.Add(string.Format("{0}_{1}", JUNPYO_ID, "1_3"), Pic1_3.Image);

            if (Pic2_1.Image != null)
                dicImages.Add(string.Format("{0}_{1}", JUNPYO_ID, "2_1"), Pic2_1.Image);

            if (Pic2_2.Image != null)
                dicImages.Add(string.Format("{0}_{1}", JUNPYO_ID, "2_2"), Pic2_2.Image);

            if (Pic2_3.Image != null)
                dicImages.Add(string.Format("{0}_{1}", JUNPYO_ID, "2_3"), Pic2_3.Image);

            try
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.J_DATE ");
                strSql.AppendLine("   FROM MESURING A ");
                strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sJ_Date = string.Empty;
                if (dt.Rows.Count > 0)
                {
                    sJ_Date = dt.Rows[0]["J_DATE"]?.ToString();
                }

                string[] sJDateArr = sJ_Date.Split(' ');
                string sJDate = sJDateArr[0];
                string[] strArr = sJDate.Split('-');
                //string ftpPath = "ftp://192.168.0.202/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
                string ftpPath = "ftp://"+ComnEtcFunc.FTP_IP+"/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                string fileName = null;

                FTPDirectioryCheck(ftpPath, user, pw);
                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                string[] filesInDirectory = null;
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    req1Res.Close();
                }
                
                foreach (KeyValuePair<string, Image> item in dicImages)
                {
                    //해당 파일 Index
                    int findIndex = Array.FindIndex(filesInDirectory, i => i == string.Format("{0}", item.Key));
                    //기존 파일이 있을경우 백업폴더 이동 후 기존폴더 파일 삭제
                    if (findIndex >= 0)
                    {
                        fileName = filesInDirectory[findIndex];
                        //string path = string.Format(@"ftp://192.168.0.202/Images_Backup/{0}/{1}/{2}", strArr[0], strArr[1], sJDate);
                        string path = string.Format(@"ftp://"+ComnEtcFunc.FTP_IP+"/Images_Backup/{0}/{1}/{2}", strArr[0], strArr[1], sJDate);

                        //기존파일 이미지 Byte[] 파싱
                        byte[] moveData;
                        using (var ms = new MemoryStream())
                        {
                            item.Value.Save(ms, item.Value.RawFormat);
                            moveData = ms.ToArray();
                        }

                        FTPDirectioryCheck(path, user, pw);
                        FTPUpload(string.Format(path + "/{0}", fileName), user, pw, moveData);
                        FTPDelete(string.Format(ftpPath + "/{0}", fileName), user, pw);
                    }

                    //이미지 Byte[] 파싱
                    byte[] data;
                    using (var ms = new MemoryStream())
                    {
                        item.Value.Save(ms, item.Value.RawFormat);
                        data = ms.ToArray();
                    }

                    Console.Write(DateTime.Now.ToString("yyyyMMdd"));
                    Console.Write(DateTime.Now.ToString("HHmmss"));

                    //string sUploadPath = string.Format(ftpPath + "/{0}_{1}_{2}.jpg", DateTime.Now.ToString("yyyyMMdd")
                    //                                                               , DateTime.Now.ToString("HHmmss")
                    //                                                               , item.Key);
                    string sUploadPath = string.Format(ftpPath + "/{0}.jpg", item.Key);
                    //업로드 위한 설정
                    FtpWebRequest req = (FtpWebRequest)WebRequest.Create(sUploadPath);
                    req.Method = WebRequestMethods.Ftp.UploadFile;
                    req.Credentials = new NetworkCredential(user, pw);
                    req.UsePassive = false;
                    // RequestStream에 데이타를 쓴다
                    req.ContentLength = data.Length;
                    Stream reqStream = req.GetRequestStream();
                    reqStream.Write(data, 0, data.Length);
                    reqStream.Close();

                    FtpWebResponse response = (FtpWebResponse)req.GetResponse();
                    Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);
                    response.Close();
                }
                XtraMessageBox.Show("이미지가 저장되었습니다.");

                DialogResult = DialogResult.OK;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
        }

        

        /// <summary>
        /// FTP 경로에 Upload
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public void FTPUpload(string directoryPath, string _FTPuserID, string _FTPpassword, byte[] data)
        {
            //업로드 위한 설정
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(directoryPath);
            req.Method = WebRequestMethods.Ftp.UploadFile;
            req.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            req.UsePassive = false;
            // RequestStream에 데이타를 쓴다
            req.ContentLength = data.Length;
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();

            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }

        /// <summary>
        /// FTP 경로에 Delete
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public void FTPDelete(string directoryPath, string _FTPuserID, string _FTPpassword)
        {
            FtpWebRequest requestFileDelete = (FtpWebRequest)WebRequest.Create(directoryPath);
            requestFileDelete.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            requestFileDelete.Method = WebRequestMethods.Ftp.DeleteFile;

            FtpWebResponse responseFileDelete = (FtpWebResponse)requestFileDelete.GetResponse();
        }

    /// <summary>
    /// FTP 경로의 디렉토리를 점검하고 없으면 생성
    /// </summary>
    /// <param name="directoryPath">디렉터리 경로 입니다.</param>
    public void FTPDirectioryCheck(string directoryPath, string _FTPuserID, string _FTPpassword)
        {
            string[] directoryPaths = directoryPath.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string[] result = new string[directoryPaths.Length - 1];
            for(int i = 0; i < result.Length; i++)
            {
                if(i == 0)
                {
                    result[i] = directoryPaths[i] + "//" + directoryPaths[i + 1];
                }
                else
                {
                    result[i] = directoryPaths[i + 1];
                }
            }

            string currentDirectory = string.Empty;
            foreach (string directory in result)
            {
                currentDirectory += string.Format("{0}/", directory);
                if (!IsExistDirectory(currentDirectory, _FTPuserID, _FTPpassword))
                {
                    MakeDirectory(currentDirectory, _FTPuserID, _FTPpassword);
                }
            }
        }

        private bool IsExistDirectory(string Directory, string _FTPuserID, string _FTPpassword)
        {
            try
            {
                var request = (FtpWebRequest)WebRequest.Create(Directory);
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);                 
                
                using (request.GetResponse())
                {
                    return true;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }

        private bool MakeDirectory(string Directory, string _FTPuserID, string _FTPpassword)
        {
            string URI = Directory;

            System.Net.FtpWebRequest ftp = WebRequest.Create(new Uri(URI)) as FtpWebRequest;
            ftp.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            ftp.UseBinary = true;
            ftp.UsePassive = true;
            ftp.Timeout = 10000;
            ftp.Method = System.Net.WebRequestMethods.Ftp.MakeDirectory;             

            try
            {
                string str = GetStringResponse(ftp);
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private string GetStringResponse(FtpWebRequest ftp)
        {
            string result = "";
            using (FtpWebResponse response = (FtpWebResponse)ftp.GetResponse())
            {
                long size = response.ContentLength;
                using (Stream datastream = response.GetResponseStream())
                {
                    if (datastream != null)
                    {
                        using (StreamReader sr = new StreamReader(datastream))
                        {
                            result = sr.ReadToEnd();
                            sr.Close();
                        }

                        datastream.Close();
                    }
                }

                response.Close();
            }

            return result;
        }

        private void BtnCancle_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BtnUpload1_1_Click(object sender, EventArgs e)
        {
            LoadToPictureBox(Pic1_1);
        }
        
        private void BtnUpload1_2_Click(object sender, EventArgs e)
        {
            LoadToPictureBox(Pic1_2);
        }

        private void BtnUpload1_3_Click(object sender, EventArgs e)
        {
            LoadToPictureBox(Pic1_3);
        }

        private void BtnUpload2_1_Click(object sender, EventArgs e)
        {
            LoadToPictureBox(Pic2_1);
        }

        private void BtnUpload2_2_Click(object sender, EventArgs e)
        {
            LoadToPictureBox(Pic2_2);
        }

        private void BtnUpload2_3_Click(object sender, EventArgs e)
        {
            LoadToPictureBox(Pic2_3);
        }

        string FileDirectory = string.Empty;
        private void LoadToPictureBox(PictureEdit pic)
        {
            try
            {
                XtraOpenFileDialog fileDlg = new XtraOpenFileDialog();
                string sFileName = string.Empty;

                if (string.IsNullOrEmpty(FileDirectory))
                    fileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                else
                    fileDlg.InitialDirectory = FileDirectory;

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileDirectory = fileDlg.FileName.Replace(fileDlg.SafeFileName, "");
                    sFileName = fileDlg.FileName;
                    pic.Image = Image.FromFile(sFileName);
                }
            }
            catch (Exception ex)
            {

            }
        }
        
    }
}