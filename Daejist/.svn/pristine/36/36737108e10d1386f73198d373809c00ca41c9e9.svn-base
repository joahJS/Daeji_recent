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
using System.Drawing.Imaging;
using System.IO;
using ComLib;
using System.Net;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Drawing.Printing;
namespace AccAdm
{
    public partial class AccMeasureDevPopImg : DevExpress.XtraEditors.XtraForm
    {
        public AccMeasureDevPopImg()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public string CDATE;
        public string _JUNPYOID;
        public string _sInitDir;
        public string _sSaveDir;
        public Form PARENT_FORM;
        Image image;
        string[] sArrImg_File;

        private void AccMeasureDevPopImg_Load(object sender, EventArgs e)
        {
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccMeasureDevImg");

            try
            {
                Cursor = Cursors.WaitCursor;
                
                GetMDate(_JUNPYOID);
                GetImagesFromFTP();

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GetMDate(string sJunpyoId)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();

           
            strSql.AppendLine("SELECT CONVERT(CHAR(8),CONVERT(DATETIME, ENT_DT),112) AS DT");
            strSql.AppendLine("  FROM mesure_CustomImg_info");
            strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0)
            {
                CDATE = dt.Rows[0]["DT"].ToString();
            }
            else
            {
                string d_temp = Convert.ToString(DateTime.Today.ToString());
                CDATE =  d_temp.Substring(0, 4) + d_temp.Substring(5, 2) + d_temp.Substring(8, 2);
            }
        }

        private void GetImagesFromFTP()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                //string sInitDir = string.Format(@"ftp://192.168.0.202/Custom_Images/{0}/{1}/{2}/{3}", CDATE.Substring(0, 4), CDATE.Substring(4, 2), CDATE.Substring(6, 2), _JUNPYOID);
                string sInitDir = string.Format(@"ftp://{0}/Custom_Images/{1}/{2}/{3}/{4}", ComnEtcFunc.FTP_IP, CDATE.Substring(0, 4), CDATE.Substring(4, 2), CDATE.Substring(6, 2), _JUNPYOID);
                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(sInitDir);
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                string[] filesInDirectory = null;
                Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();
                    
                    foreach (string filePath in filesInDirectory)
                    {
                        string[] filesCopy = filePath.Split('\\');
                        dicImages.Add(filesCopy[filesCopy.Length - 1], DownloadFTPFile(string.Format(@"{0}\{1}", sInitDir, filePath), user, pw));
                    }
                }

                DataTable dt = new DataTable();
                dt.TableName = "Table1";
                dt.Columns.Add("IMG", typeof(byte[]));
                dt.Columns.Add("FILENAME");

                foreach (KeyValuePair<string, Image> item in dicImages)
                {
                    MemoryStream ms = new MemoryStream(ImageToByteArray(item.Value));
                    image = Image.FromStream(ms);
                    picEdit.Image = (Image)image.Clone();
                    picEdit.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Squeeze;
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                return;
            }
            Cursor = Cursors.Default;
        }

        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            sArrImg_File = null;
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                sArrImg_File = dialog.FileNames;
                if (XtraMessageBox.Show(string.Format("{0}의 건에 대하여 저장을 진행하시겠습니까?", sArrImg_File.Length), "업로드유무", MessageBoxButtons.YesNo) == DialogResult.OK)
                    return;
            }
            else
            {
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                string sInitDir = string.Format(@"ftp://192.168.0.202/Custom_Images/{0}/{1}/{2}/{3}", CDATE.Substring(0, 4), CDATE.Substring(4, 2), CDATE.Substring(6, 2), _JUNPYOID);
                //string sInitDir = string.Format(@"ftp://{0}/Custom_Images/{1}/{2}/{3}/{4}", ComnEtcFunc.FTP_IP, CDATE.Substring(0, 4), CDATE.Substring(4, 2), CDATE.Substring(6, 2), _JUNPYOID);
                //FTP 서버
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;


                int i = 1;
                string sFileName = sArrImg_File[0];
                
                    string sFileOrgNm = Path.GetFileName(sFileName);
                    string ext = Path.GetExtension(sFileName);

                    Image img = Image.FromFile(sFileName);
                    byte[] reduce_data = SaveJpeg(sFileName, img, 10); //파일용량 줄이기

                    string sSaveDir = string.Format(@"{0}/{1}", sInitDir, _JUNPYOID + DateTime.Now.ToString("_yyMMddHHmmss") + ext);
                    FTPDirectioryCheck(sInitDir, user, pw);
                    FTPUpload(sSaveDir, user, pw, reduce_data);
                _sSaveDir = sSaveDir;
                UpdateImagesCount(sInitDir, sSaveDir);

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.ToString());
                return;
            }

            GetImagesFromFTP();

            string InsertSql, DelSql;

            DelSql = "DELETE FROM mesure_CustomImg_info WHERE JUNPYOID = '" + _JUNPYOID + "'  ";
            SqlCommand mwDel = new SqlCommand(DelSql, DBConn.dbCon);
            mwDel.ExecuteNonQuery();


            InsertSql = " INSERT INTO  mesure_CustomImg_info VALUES (@JUNPYOID, @ISPT_NO, @yymm, 1 ) ";
            SqlCommand mwInsert = new SqlCommand(InsertSql, DBConn.dbCon);
            mwInsert.Parameters.Add("@JUNPYOID", SqlDbType.VarChar).Value = _JUNPYOID;
            mwInsert.Parameters.Add("@ISPT_NO", SqlDbType.VarChar).Value = _sSaveDir;
            mwInsert.Parameters.Add("@yymm", SqlDbType.VarChar).Value = CDATE;

            mwInsert.ExecuteNonQuery();
            Cursor = Cursors.Default;

        }

        private void UpdateImagesCount(string sInitDir, string sArrImg_File)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(sInitDir);
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                string[] filesInDirectory = null;
                Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    Cursor = Cursors.WaitCursor;

                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE mesure_CustomImg_info ");
                    strSql.AppendLine("    SET IMG_CNT = @IMG_CNT ");
                    strSql.AppendLine("       ,ISPT_NO = @sArrImg_File ");
                    strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", _JUNPYOID);
                    cmd.Parameters.AddWithValue("@IMG_CNT", filesInDirectory.Length);
                    cmd.Parameters.AddWithValue("@sArrImg_File", sArrImg_File);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            Cursor = Cursors.Default;
        }

        /// <summary> 
        /// Saves an image as a jpeg image, with the given quality 
        /// </summary> 
        /// <param name="path"> Path to which the image would be saved. </param> 
        /// <param name="quality"> An integer from 0 to 100, with 100 being the highest quality. </param> 
        public byte[] SaveJpeg(string path, Image img, int quality)
        {
            if (quality < 0 || quality > 100)
                throw new ArgumentOutOfRangeException("quality must be between 0 and 100.");

            // Encoder parameter for image quality 
            EncoderParameter qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
            // JPEG image codec 
            ImageCodecInfo jpegCodec = GetEncoderInfo("image/jpeg");
            EncoderParameters encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = qualityParam;

            byte[] Data;
            using (var ms = new MemoryStream())
            {
                img.Save(ms, jpegCodec, encoderParams);
                Data = ms.ToArray();
            }
            return Data;
        }

        /// <summary> 
        /// Returns the image codec with the given mime type 
        /// </summary> 
        private ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            // Get image codecs for all image formats 
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();

            // Find the correct image codec 
            for (int i = 0; i < codecs.Length; i++)
                if (codecs[i].MimeType == mimeType)
                    return codecs[i];

            return null;
        }

        /// <summary>
        /// FTP 경로의 디렉토리를 점검하고 없으면 생성
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public void FTPDirectioryCheck(string directoryPath, string _FTPuserID, string _FTPpassword)
        {
            string[] directoryPaths = directoryPath.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string[] result = new string[directoryPaths.Length - 1];
            for (int i = 0; i < result.Length; i++)
            {
                if (i == 0)
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

        #region FTP 파일 삭제하기 - DeleteFTPFile(targetURI, userID, password)

        /// <summary>
        /// FTP 파일 삭제하기
        /// </summary>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <param name="targetURI">타겟 URI</param>
        /// <returns>처리 결과</returns>
        public bool DeleteFTPFile(string targetURI, string userID, string password)
        {
            try
            {
                FtpWebRequest ftpWebRequest = WebRequest.Create(targetURI) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(userID, password);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;
            }
            catch
            {
                return false;
            }

            return true;
        }

        #endregion


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
            req.UsePassive = true;
            // RequestStream에 데이타를 쓴다
            req.ContentLength = data.Length;
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();

            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }

        public void FTPDel(string directoryPath, string _FTPuserID, string _FTPpassword, byte[] data)
        {
            //삭제 위한 설정
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(directoryPath);
            req.Method = WebRequestMethods.Ftp.DeleteFile;
            //req.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            //req.UsePassive = true;
            //// RequestStream에 데이타를 쓴다
            //req.ContentLength = data.Length;
            //Stream reqStream = req.GetRequestStream();
            //reqStream.Write(data, 0, data.Length);
            //reqStream.Close();

            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
        #region 프린터
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            PrintDocument Pd = new PrintDocument();
            printDialog1.Document = Pd;
            Pd.PrintPage += new PrintPageEventHandler(PrintImage);
            if(printDialog1.ShowDialog() == DialogResult.OK)
            {
                Pd.Print();
            }
        }
        private  void PrintImage(object o, PrintPageEventArgs e)
        {
            Image img = picEdit.Image;
            Point loc = new Point(100, 100);
            e.Graphics.DrawImage(img, loc);
        }
        #endregion

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sMSG = string.Format("해당 계근 건에 대하여 삭제를 진행하시겠습니까?");
              
            if (XtraMessageBox.Show(sMSG, "이미지 삭제건", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();

            strSql.AppendLine(" SELECT ISPT_NO                                 ");
            strSql.AppendLine("    FROM mesure_CustomImg_info                  ");
            strSql.AppendLine("   WHERE JUNPYOID = '" + _JUNPYOID + "'         ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
            if (dt.Rows.Count > 0)
            {
                string sFileName = dt.Rows[0]["ISPT_NO"].ToString();

                string sInitDir = sFileName;
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;

                //Ftp 파일 삭제
                FtpWebRequest req = (FtpWebRequest)WebRequest.Create(sInitDir);
                req.Credentials = new NetworkCredential(user, pw);
                req.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse response = (FtpWebResponse)req.GetResponse();
                response.Close();
                picEdit.Image = null;
                //Table 내용삭제
                string DelSql;

                DelSql = "DELETE FROM mesure_CustomImg_info WHERE JUNPYOID = '" + _JUNPYOID + "'  ";
                SqlCommand mwDel = new SqlCommand(DelSql, DBConn.dbCon);
                mwDel.ExecuteNonQuery();
            }
            else
            {
                
            }


        }
    }
}