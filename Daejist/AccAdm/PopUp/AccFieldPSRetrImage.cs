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
using MySql.Data.MySqlClient;
using ComLib;
using System.IO;
using System.Net;

namespace AccAdm
{
    public partial class AccFieldPSRetrImage : DevExpress.XtraEditors.XtraForm
    {
        public AccFieldPSRetrImage()
        {
            InitializeComponent();
        }

        public string _JunpyoID { get; set; }
        public string _JDate { get; set; }

        private void AccFieldPSRetrImage_Load(object sender, EventArgs e)
        {
            string[] sJDateArr = _JDate?.ToString().Split(' ');
            string sJDate = sJDateArr[0];
            string[] strArr = sJDate.Split('-');
            //string ftpPath = @"ftp://192.168.0.202/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
            string ftpPath = @"ftp://"+ComnEtcFunc.FTP_IP+"/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
            string user = ComnEtcFunc.FTP_USER;
            string pw = ComnEtcFunc.FTP_PW;

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

            try
            {
                string[] filesInDirectory = null;
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    foreach (KeyValuePair<string, Image> item in dicPicture)
                    {
                        //해당 파일 Index
                        int findIndex = Array.FindIndex(filesInDirectory, i => i == string.Format("{0}_{1}.jpg", _JunpyoID, item.Key));
                        if (findIndex >= 0)
                        {
                            string fileName = filesInDirectory[findIndex];
                            dicCopy[item.Key] = DownloadFTPFile(string.Format(@"{0}\{1}", ftpPath, fileName), user, pw);
                        }
                    }
                }

                pictureEdit1.Image = dicCopy["1_1"];
                pictureEdit2.Image = dicCopy["1_2"];
                pictureEdit3.Image = dicCopy["1_3"];
                pictureEdit4.Image = dicCopy["2_1"];
                pictureEdit5.Image = dicCopy["2_2"];
                pictureEdit6.Image = dicCopy["2_3"];

                //dicPicture = null;

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                pictureEdit1.Image = null;
                pictureEdit2.Image = null;
                pictureEdit3.Image = null;
                pictureEdit4.Image = null;
                pictureEdit5.Image = null;
                pictureEdit6.Image = null;

                return;
            }

            #region[이전 코드]

            //try
            //{
            //    StringBuilder strSql = new StringBuilder();

            //    strSql.AppendLine(" SELECT J_DATE           ");
            //    strSql.AppendLine("   FROM MESURING         ");
            //    strSql.AppendLine("  WHERE JUNPYOID = " + junpyoID + "");

            //    DataTable dt = MySqlDb.GetDataTable(MySqlDb.dbCon, strSql.ToString());
            //    DataRow dtRow = dt.Rows[0];

            //    string[] jDateArr = dtRow["J_DATE"].ToString().Split(' ');
            //    string sJDate = jDateArr[0];
            //    string[] strArr = sJDate.Split('-');
            //    string ftpPath = "ftp://192.168.0.202/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
            //    string user = "Administrator";
            //    string pw = "Admin12345";
            //    string fileName = null;
            //    string[] fileNameArr = null;

            //    FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

            //    req1.Credentials = new NetworkCredential(user, pw);
            //    req1.Method = WebRequestMethods.Ftp.ListDirectory;

            //    using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
            //    {
            //        StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
            //        string strData = reader1.ReadToEnd();

            //        string[] filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            //        reader1.Close();

            //        for (int i = 0; i < filesInDirectory.Length; i++)
            //        {
            //            fileName = filesInDirectory[i];
            //            fileNameArr = fileName.Split('_');

            //            if (fileNameArr[2].Equals(junpyoID))
            //            {
            //                if (fileNameArr[3].Equals("1"))
            //                {
            //                    if (fileNameArr[4].Substring(0, 1).Equals("1"))
            //                    {
            //                        Stream stream = GetFtpImg(ftpPath+"/"+fileName, user, pw);
            //                        pictureEdit1.Image = Image.FromStream(stream);
            //                    }
            //                    else if (fileNameArr[4].Substring(0, 1).Equals("2"))
            //                    {
            //                        Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
            //                        pictureEdit2.Image = Image.FromStream(stream);
            //                    }
            //                    else if (fileNameArr[4].Substring(0, 1).Equals("3"))
            //                    {
            //                        Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
            //                        pictureEdit3.Image = Image.FromStream(stream);
            //                    }
            //                }
            //                else if (fileNameArr[3].Equals("2"))
            //                {
            //                    if (fileNameArr[4].Substring(0, 1).Equals("1"))
            //                    {
            //                        Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
            //                        pictureEdit4.Image = Image.FromStream(stream);
            //                    }
            //                    else if (fileNameArr[4].Substring(0, 1).Equals("2"))
            //                    {
            //                        Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
            //                        pictureEdit5.Image = Image.FromStream(stream);
            //                    }
            //                    else if (fileNameArr[4].Substring(0, 1).Equals("3"))
            //                    {
            //                        Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
            //                        pictureEdit6.Image = Image.FromStream(stream);
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            //catch (WebException)
            //{

            //}

            #endregion[이전 코드]
        }

        private Stream GetFtpImg(string ftpPath , string user, string pw)
        {
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(ftpPath);

            req.Credentials = new NetworkCredential(user, pw);
            req.Method = WebRequestMethods.Ftp.DownloadFile;

            FtpWebResponse resp = (FtpWebResponse)req.GetResponse();

            Stream stream = resp.GetResponseStream();

            return stream;
        }

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

        private void pictureEdit_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}