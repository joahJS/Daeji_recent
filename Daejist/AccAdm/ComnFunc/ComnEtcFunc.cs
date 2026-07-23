using ComLib;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Popbill;
using Popbill.Fax;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace AccAdm
{
    /*
     * 클래스설명 : 공통모듈
     * 생성자 : 고혜성 / 생성일시 : 2021-01-19
     * ----------------------------HISTORY-------------------------
     * 1. 수정일시: 2021-07-06
     *      수정자: 정은영
     *   변경 사항: (현업요청)엑셀양식추가
     *         ID : #001
     *        
     * 2. 수정일자: 2022-06-08
     *      수정자: 정은영
     *    변경사항: 1. 요일 주차 구하기 함수 추가
     *          ID: #002
     *          
     * 3. 수정일자: 2022-06-21
     *      수정자: 정은영
     *    변경사항: 1. 포커스 색변경 함수 추가
     *              2. 폼 상단아이콘 셋팅 함수추가
     *          ID: #003
     *          
     * 4. 수정일자: 2022-06-21
     *      수정자: 정은영
     *    변경사항: 1. repoGridLookup 설정 추가
     *          ID: #004
     *          
     * 5. 수정일자: 2022-09-27
     *      수정자: 정은영
     *          ID: #005
     *    변경사항: 1. 이전일 다음일 코드 추가
     *              2. 이전월 다음월 코드 추가
     * 
     */
    class ComnEtcFunc
    {
        public static string CashCode = "0101"; //계정코드 현금 초기세팅 값(HARD_CODING, 현금계정쓰는 로직에서 가져올 예정)
        //대지 ftp
        public static string FTP_IP = "192.168.0.202";
        public static string FTP_USER = "ftpuser";
        public static string FTP_PW = "Admin12345"; 
        public static string FTP_ROOT = "ftp://" + FTP_IP;

        public static string MESURING_IP = "192.168.0.164";
        public static bool IsFaxTest = false; //true - 개발용(테스트베드), false - 상업용(실서비스)
        public static string _INI_PATH = Application.StartupPath + @"\VersionCheck.ini";

        //암복호화 키
        public static string _SECRET_KEY = "skdjvopiehwgekgjpdovjdikwfjpoefjjeghwoegkjpo"; //44자 //주민번호용 키
        public static string _SECRET_KEY2 = "skdjvopiehwgekgjpdovjdikwfjpoefj"; //32자 //비밀번호용 키

        //public static Icon GetFavicon()
        //{
        //    Bitmap bitmap = null;
        //    Icon icon = null;
        //    bitmap = Properties.Resources.DJ_ICON; 
        //    if (bitmap != null)
        //        icon = Icon.FromHandle(bitmap.GetHicon()); 
        //    return icon;
        //}

        #region [ 스마트 1번가 접속 LOG API ]
        // 22.09.15 추가, 스마트 1번가 접속 로그 기록 API
        public static void LogAPI(string sGuBun, string sComip)
        {
            try
            {
                JObject data = new JObject();

                data.Add("crtfcKey", "$5$API$aQova2EnezC5li3vEJWFt.DUfoc3t0wSeqVn62NtBo7");
                data.Add("logDt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                data.Add("useSe", sGuBun);
                data.Add("sysUser", FmMainToolBar2.drUser["USRNM"]?.ToString());
                data.Add("conectIp", GetLocalIP());
                data.Add("dataUsgqty", "0");

                var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://log.smart-factory.kr/apisvc/sendLogDataJSON.do?logData=" + data.ToString());
                httpWebRequest.Method = "GET";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    JObject res = JObject.Parse(result);
                }
            }
            catch
            {

            }
        }
        #endregion

        public static string GetLayoutPath()
        {
            return Application.StartupPath + @"\xaml\" + FmMainToolBar2.UserID;
        }

        /*
         * From : 이번 달 1일
         * To : 오늘일자
         */
        public static void SetDateFromToValue(DateEdit ymdFrom, DateEdit ymdTo)
        {
            DateTime today = DateTime.Now.Date;
            ymdFrom.EditValue = today.AddDays(1 - today.Day);
            ymdTo.EditValue = today;
        }
        public static void SetDateFromFromValue(DateEdit ymdFrom, DateEdit ymdTo)
        {
            DateTime today = DateTime.Now.Date;
            ymdFrom.EditValue = today;
            ymdTo.EditValue = today;
        }
        public static void SetDateFromValue(DateEdit ymdFrom)
        {
            DateTime today = DateTime.Now.Date;
            ymdFrom.EditValue = today.AddDays(1 - today.Day);
        }
        public static void SetDateToValue(DateEdit ymdTo)
        {
            DateTime today = DateTime.Now.Date;
            ymdTo.EditValue = today;
        }
        public static void SetDateMonthValue(DateEdit ymdFrom)
        {
            DateTime today = DateTime.Now.Date;
            ymdFrom.EditValue = today.AddDays(1 - today.Month);
        }

        public static void YmdFromToValuesCheck(DateEdit ymdFrom, DateEdit ymdTo)
        {
            string sYmdFrom = ymdFrom.EditValue.ToString();
            string sYmdTo = ymdTo.EditValue.ToString();

            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("검색기간을 설정하세요.");
                ymdFrom.SelectAll();
                ymdFrom.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("검색기간을 설정하세요.");
                ymdTo.SelectAll();
                ymdTo.Focus();
                return;
            }
        }

        //#005
        public static string PrevDate(string sDate)
        {
            string result = string.Empty;
            if (DateTime.TryParse(sDate, out DateTime datetime))
            {
                datetime = datetime.AddDays(-1);
                result = datetime.ToString("yyyy-MM-dd");
            }

            return result;
        }

        public static string NextDate(string sDate)
        {
            string result = string.Empty;
            if (DateTime.TryParse(sDate, out DateTime datetime))
            {
                datetime = datetime.AddDays(1);
                result = datetime.ToString("yyyy-MM-dd");
            }

            return result;
        }

        public static string PrevMonth(string sDate)
        {
            string result = string.Empty;
            if (DateTime.TryParse(sDate, out DateTime datetime))
            {
                datetime = datetime.AddMonths(-1);
                result = datetime.ToString("yyyy-MM-dd");
            }

            return result;
        }

        public static string NextMonth(string sDate)
        {
            string result = string.Empty;
            if (DateTime.TryParse(sDate, out DateTime datetime))
            {
                datetime = datetime.AddMonths(1);
                result = datetime.ToString("yyyy-MM-dd");
            }

            return result;
        }

        public static string PrevYear(string sDate)
        {
            string result = string.Empty;
            if (DateTime.TryParse(sDate, out DateTime datetime))
            {
                datetime = datetime.AddYears(-1);
                result = datetime.ToString("yyyy-MM-dd");
            }

            return result;
        }

        public static string NextYear(string sDate)
        {
            string result = string.Empty;
            if (DateTime.TryParse(sDate, out DateTime datetime))
            {
                datetime = datetime.AddYears(1);
                result = datetime.ToString("yyyy-MM-dd");
            }

            return result;
        }
        //#005 end

        //#002
        #region 월 주차 구하기 - GetWeekOfMonth(sourceDate, cultureInfo)

        /// <summary>
        /// 월 주차 구하기
        /// </summary>
        /// <param name="sourceDate">소스 일자</param>
        /// <param name="cultureInfo">문화 정보</param>
        /// <returns>월 주차</returns>
        public static int GetWeekOfMonth(DateTime sourceDate, CultureInfo cultureInfo)
        {
            if (cultureInfo == null)
            {
                cultureInfo = CultureInfo.CurrentCulture;
            }

            DateTime firstDayOfMonth = DateTime.Parse(sourceDate.ToString("yyyy-MM-01"));

            int firstWeekOfMonth = GetWeekOfYear(firstDayOfMonth, cultureInfo);
            int sourceWeekOfMonth = GetWeekOfYear(sourceDate, cultureInfo);

            return (sourceWeekOfMonth - firstWeekOfMonth) + 1;
        }

        #endregion

        #region 월 주차 구하기 - GetWeekOfMonth(sourceDate)

        /// <summary>
        /// 월 주차 구하기
        /// </summary>
        /// <param name="sourceDate">소스 일자</param>
        /// <returns>월 주차</returns>
        public static int GetWeekOfMonth(DateTime sourceDate)
        {
            return GetWeekOfMonth(sourceDate, null);
        }

        #endregion

        #region 연도 주차 구하기 - GetWeekOfYear(sourceDate, cultureInfo)

        /// <summary>
        /// 연도 주차 구하기
        /// </summary>
        /// <param name="sourceDate">소스 일자</param>
        /// <param name="cultureInfo">문화 정보</param>
        /// <returns>연도 주차</returns>
        public static int GetWeekOfYear(DateTime sourceDate, CultureInfo cultureInfo)
        {
            if (cultureInfo == null)
            {
                cultureInfo = CultureInfo.CurrentCulture;
            }

            CalendarWeekRule calendarWeekRule = cultureInfo.DateTimeFormat.CalendarWeekRule;

            DayOfWeek firstDayOfWeek = cultureInfo.DateTimeFormat.FirstDayOfWeek;

            return cultureInfo.Calendar.GetWeekOfYear(sourceDate, calendarWeekRule, firstDayOfWeek);
        }

        #endregion

        #region 연도 주차 구하기 - GetWeekOfYear(sourceDate)

        /// <summary>
        /// 연도 주차 구하기
        /// </summary>
        /// <param name="sourceDate">소스 일자</param>
        /// <returns>연도 주차</returns>
        public static int GetWeekOfYear(DateTime sourceDate)
        {
            return GetWeekOfYear(sourceDate, null);
        }

        #endregion

        //#003
        public static void gp_SetColorFocused(Control sender)
        {
            // Control에 포커스 가면 색 변경 (SHS: 2020.12.02)
            foreach (Control ctl in sender.Controls)
            {
                if (ctl is TextEdit) { TextEdit tb = (TextEdit)ctl; tb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(255, 255, 128); }
                else
                if (ctl is MemoEdit) { MemoEdit cb = (MemoEdit)ctl; cb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(255, 255, 128); }
                else
                if (ctl is LookUpEdit) { LookUpEdit cb = (LookUpEdit)ctl; cb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(255, 255, 128); }
                else
                if (ctl is ComboBoxEdit) { ComboBoxEdit cb = (ComboBoxEdit)ctl; cb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(255, 255, 128); }
            }
        }

        //#003
        public static Icon GetFavicon()
        {
            Icon icon = null;
            icon = Properties.Resources.DJ_ICON;

            return icon;
        }
        
        public static void ExportExcelFile(string sFileName, GridControl grid)
        {
            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                //string sFileNM = sFileName + DateTime.Now.ToLongDateString().Replace(" ", "");
                string sFileNM = sFileName;
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    
                    grid.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
                }
                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        //엑셀 임시저장
        public static void ExportExcelFile(string sPath, string sFileName, GridControl grid)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(sPath);

                if(di.Exists == false)
                {
                    di.Create();
                }

                grid.ExportToXls(sPath + sFileName);
                Process.Start(sPath + sFileName);
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        #region[권한 적용]
        public static SqlConnection DbConn()
        {
            string strConn = "server = 112.216.189.178,12501; uid = pineit; pwd = pineit0401; database = daejierp";
            SqlConnection dbCon;
            try
            {
                dbCon = new SqlConnection(strConn);
                dbCon.Open();
                return dbCon;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public static void DbDisConn(SqlConnection dbCon1)
        {
            try
            {
                dbCon1.Close();
                dbCon1.Dispose();
                dbCon1 = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public static void SetMenuListByAuth(DataTable dt, BarSubItem Big_Basic, BarSubItem Big_Prod, BarSubItem Big_InOut, BarSubItem Big_Device, BarSubItem Big_QC
            , BarSubItem Big_HR, BarSubItem Big_ACC, BarSubItem Big_Report, BarSubItem Big_System, BarButtonItem[] SubCategory, string[] ArrSubItemName)
        {
            //대분류에 속한 컨트롤의 권한 체크를 위한 카운트 값, Loop 후 0일 시 해당 대분류는 Visible False;
            int Chk_Basic = 0;
            int Chk_Prod = 0;
            int Chk_InOut = 0;
            int Chk_Device = 0;
            int Chk_QC = 0;
            int Chk_HR = 0;
            int Chk_ACC = 0;
            int Chk_Report = 0;
            int Chk_System = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sPgmId = dt.Rows[i]["PGMID"]?.ToString();
                string sPgGrp = dt.Rows[i]["PGGRP"]?.ToString();
                string sUseYn = dt.Rows[i]["USE_Y"]?.ToString();

                if (sUseYn.Equals("Y"))
                {
                    //SubCategory 중 PGGRP 컬럼 값 매칭 후 Visible 변환
                    for (int j = 0; j < SubCategory.Length; j++)
                    {
                        if (ArrSubItemName[j].Equals(sPgmId))
                        {
                            SubCategory[j].Visibility = BarItemVisibility.Always;

                            //공통코드로 부터 코드값 HARD_CODING
                            if (sPgGrp.Equals("11"))
                            {
                                Chk_Basic++;
                            }
                            else if (sPgGrp.Equals("22"))
                            {
                                Chk_Prod++;
                            }
                            else if (sPgGrp.Equals("33"))
                            {
                                Chk_InOut++;
                            }
                            else if (sPgGrp.Equals("44"))
                            {
                                Chk_Device++;
                            }
                            else if (sPgGrp.Equals("55"))
                            {
                                Chk_QC++;
                            }
                            else if (sPgGrp.Equals("66"))
                            {
                                Chk_HR++;
                            }
                            else if (sPgGrp.Equals("77"))
                            {
                                Chk_ACC++;
                            }
                            else if (sPgGrp.Equals("88"))
                            {
                                Chk_Report++;
                            }
                            else if (sPgGrp.Equals("99"))
                            {
                                Chk_System++;
                            }
                        }
                    }
                }
            }

            //카운트 값 0일시 해당 서브Control 미존재로 간주, 해당 대분류 Control Visible False 설정
            if (Chk_Basic == 0)
            {
                Big_Basic.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_Basic.Visibility = BarItemVisibility.Always;
            }

            if (Chk_Prod == 0)
            {
                Big_Prod.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_Prod.Visibility = BarItemVisibility.Always;
            }

            if (Chk_InOut == 0)
            {
                Big_InOut.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_InOut.Visibility = BarItemVisibility.Always;
            }

            if (Chk_Device == 0)
            {
                Big_Device.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_Device.Visibility = BarItemVisibility.Always;
            }

            if (Chk_QC == 0)
            {
                Big_QC.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_QC.Visibility = BarItemVisibility.Always;
            }

            if (Chk_HR == 0)
            {
                Big_HR.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_HR.Visibility = BarItemVisibility.Always;
            }

            if (Chk_ACC == 0)
            {
                Big_ACC.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_ACC.Visibility = BarItemVisibility.Always;
            }

            if (Chk_Report == 0)
            {
                Big_Report.Visibility = BarItemVisibility.Never;
            }
            else
            {
                Big_Report.Visibility = BarItemVisibility.Always;
            }

            if (Chk_System == 0)
            {
                Big_System.Visibility = BarItemVisibility.Always;
            }
            else
            {
                Big_System.Visibility = BarItemVisibility.Always;
            }
        }
        #endregion[권한 적용]

        /// <summary>
        /// 권한정보를 가져오는 모듈
        /// </summary>
        /// <param name="sUsrCd">유저코드 </param>
        /// <returns> 첫번째 인자 + 두번째 인자 </returns>
        public static DataTable GetAuthInfo(string sUsrCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.*, B.PGGRP ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("   LEFT OUTER JOIN ZPGMLST B ");
            strSql.AppendLine("     ON A.PGMID = B.PGMID ");
            strSql.AppendLine("  WHERE A.USRCD = " + sUsrCd + " ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        public static string SendFexPopbil(string senderNum, string receiverNum, string receiverName,
                                            string strFileName, DateTime? reserveDT, string title, string requestNum)
        {
            // 연동신청시 발급받은 링크아이디로 수정.
            //string LinkID = "STEELNET";
            string LinkID = "LST";

            // 연동신청시 발급받은 비밀키로 수정.
            //string SecretKey = "Ljlyy4McNWmMIO+p77pE8QgDfe/AM4lQAVcGLvG30Jc=";
            string SecretKey = "M0XgPfNhJ0+c0KKTgjtgme9oT445+A0BYQiKZqzYAq0=";

            // 팩스 서비스 객체 선언
            FaxService faxService = new FaxService(LinkID, SecretKey);

            // 연동환경 설정값, true - 개발용(테스트베드), false - 상업용(실서비스)
            faxService.IsTest = ComnEtcFunc.IsFaxTest;

            // 발급된 토큰에 대한 IP 제한기능 사용여부, 권장(True)
            faxService.IPRestrictOnOff = true;

            // 광고팩스 전송여부
            bool adsYN = false;

            // 팝빌회원 사업자번호
            //string corpNum = "6062884137";
            string corpNum = "6068188502";

            // 팝빌회원 아이디
            //string userID = "daejist";

            /*
             * 2020-12-21 15:50 
             * 웹팩스 문의 끝에 UserID 세팅 값 받아옴
             */
            //string userID = "LST0821";
            FaxInfo faxInfo = new FaxInfo();
            string userID = faxInfo.SYNC_USER_ID;
            string receiptNum = null;

            try
            {
                receiptNum = faxService.SendFAX(corpNum, senderNum, receiverNum, receiverName,
                strFileName, reserveDT, userID, adsYN, title, requestNum);

                //MessageBox.Show("접수번호 : " + receiptNum, "팩스 전송");
            }
            catch (PopbillException ex)
            {
                MessageBox.Show("응답코드(code) : " + ex.code.ToString() + "\r\n" +
                                "응답메시지(message) : " + ex.Message, "팩스 전송");
            }

            return receiptNum;
        }

        public static void SetFormText(Form form)
        {
            string sFormText = FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME;
            form.Text += sFormText;
        }

        public static void SaveLayout(string sID, string sFormName, GridView[] arrGrdView)
        {
            for (int i = 0; i < arrGrdView.Length; i++)
            {
                string sFileName = string.Format("{0}{1}", sFormName, (i + 1).ToString());
                ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", sFileName, arrGrdView[i]);
            }
        }

        public static void SetLayout(string sID, string sFormName, GridView[] arrGrdView)
        {
            for (int i = 0; i < arrGrdView.Length; i++)
            {
                string sFileName = string.Format("{0}{1}", sFormName, (i + 1).ToString());
                ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", sFileName, arrGrdView[i]);
            }
        }

        public static DataTable GetDataTable(SqlConnection dc, string strSql, Dictionary<string, string> dicParams)
        {
            DataSet ds = new DataSet();

            try
            {
                SqlDataAdapter adpt = new SqlDataAdapter(strSql, dc);

                adpt.SelectCommand.Transaction = DBConn.dbTran;
                foreach (KeyValuePair<string, string> param in dicParams)
                {
                    adpt.SelectCommand.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                }

                adpt.Fill(ds);

                return ds.Tables[0];
            }
            catch (MySqlException sqlEx)
            {
                MessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }

        /// <summary>
        /// FTP 경로의 디렉토리를 점검하고 없으면 생성
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public static void FTPDirectioryCheck(string directoryPath, string _FTPuserID, string _FTPpassword)
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

        private static bool IsExistDirectory(string Directory, string _FTPuserID, string _FTPpassword)
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

        private static bool MakeDirectory(string Directory, string _FTPuserID, string _FTPpassword)
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

        private static string GetStringResponse(FtpWebRequest ftp)
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
        public static bool DeleteFTPFile(string targetURI, string userID, string password)
        {
            try
            {
                FtpWebRequest ftpWebRequest = WebRequest.Create(targetURI) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(userID, password);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
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

        public static Image DownloadFTPFile(string sourceFileURI, string user, string pw)
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

        /// <summary>
        /// FTP 파일 다운로드하기
        /// </summary>
        /// <param name="sourceFileURI">소스 파일 URI</param>
        /// <param name="targetFilePath">타겟 파일 경로</param>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <returns>처리 결과</returns>

        public static byte[] DownloadFTPFile_ByteArray(string sourceFileURI, string user, string pw)
        {
            byte[] byResult = new byte[16 * 1024];
            try
            {
                Uri sourceFileUri = new Uri(sourceFileURI);
                FtpWebRequest ftpWebRequest = WebRequest.Create(sourceFileUri) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(user, pw);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;

                Stream sourceStream = ftpWebResponse.GetResponseStream();
                byResult = streamToByteArray(sourceStream);

                sourceStream.Close();
            }
            catch (Exception ex)
            {
                return null;
            }

            return byResult;
        }

        /// <summary> 
        /// FTP 파일 다운로드하기 
        /// </summary> 
        /// <param name="sourceFileURI">소스 파일 URI</param> 
        /// <param name="targetFilePath">타겟 파일 경로</param> 
        /// <param name="user">사용자 ID</param> 
        /// <param name="pw">패스워드</param> 
        /// <returns>처리 결과</returns>
        public static bool FTPFileDownload(string sourceFileURI, string targetFilePath, string user, string pw)
        {
            try
            {
                Uri sourceFileUri = new Uri(sourceFileURI);
                FtpWebRequest ftpWebRequest = WebRequest.Create(sourceFileUri) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(user, pw);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;

                Stream sourceStream = ftpWebResponse.GetResponseStream();
                FileStream targetFileStream = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write);

                byte[] bufferByteArray = new byte[1024];

                while (true)
                {
                    int byteCount = sourceStream.Read(bufferByteArray, 0, bufferByteArray.Length);

                    if (byteCount == 0)
                    {
                        break;
                    }
                    targetFileStream.Write(bufferByteArray, 0, byteCount);
                }

                targetFileStream.Close();
                sourceStream.Close();

                return true;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return false;
            }
        }

        public static byte[] streamToByteArray(Stream input)
        {
            MemoryStream ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }
        #endregion

        /// <summary>
        /// FTP 경로에 Upload
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public static void FTPUpload(string directoryPath, string _FTPuserID, string _FTPpassword, byte[] data)
        {
            //업로드 위한 설정
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(directoryPath);
            req.Method = WebRequestMethods.Ftp.UploadFile;
            req.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            //req.UsePassive = false;
            req.UsePassive = true;//오픈서버 변경 후 ftp 오류로 인해 true로 변경 2022-09-29
            // RequestStream에 데이타를 쓴다
            req.ContentLength = data.Length;
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();

            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }

        public static byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }

        public static string GetLocalIP()
        {
            string myIP = "";
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    myIP = ip.ToString();
                }
            }

            return myIP;
        }

        //public static string InsertLog(string sPGM_ID, string sEDIT_KIND, string sEDIT_RMK)
        //{
        //    StringBuilder strSql = new StringBuilder();

        //    strSql.Clear();


        //}

        #region[DataTableToJsonObj]
        public static string DataTableToJsonObj(DataTable dt)
        {
            DataSet ds = new DataSet();
            ds.Merge(dt);
            StringBuilder JsonString = new StringBuilder();
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                JsonString.Append("[");
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    JsonString.Append("{");
                    for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    {
                        if (j < ds.Tables[0].Columns.Count - 1)
                        {
                            string sRowData = ds.Tables[0].Rows[i][j].ToString();

                            if (sRowData.Contains("\\"))
                            {
                                sRowData = sRowData.Replace("\\", "\\\\");
                            }

                            if (sRowData.Contains("\n"))
                            {
                                sRowData = sRowData.Replace("\n", "\\n");
                            }
                            
                            if (sRowData.Contains("\r"))
                            {
                                sRowData = sRowData.Replace("\r", "\\r");
                            }
                            
                            if (sRowData.Contains("\t"))
                            {
                                sRowData = sRowData.Replace("\t", "\\t");
                            }

                            if (sRowData.Contains("\""))
                            {
                                sRowData = sRowData.Replace("\"", "\\\"");
                            }

                            JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString() + "\":" + "\"" + sRowData + "\",");
                        }
                        else if (j == ds.Tables[0].Columns.Count - 1)
                        {
                            string sRowData = ds.Tables[0].Rows[i][j].ToString();

                            if (sRowData.Contains("\\"))
                            {
                                sRowData = sRowData.Replace("\\", "\\\\");
                            }

                            if (sRowData.Contains("\n"))
                            {
                                sRowData = sRowData.Replace("\n", "\\n");
                            }

                            if (sRowData.Contains("\r"))
                            {
                                sRowData = sRowData.Replace("\r", "\\r");
                            }

                            if (sRowData.Contains("\t"))
                            {
                                sRowData = sRowData.Replace("\t", "\\t");
                            }

                            if (sRowData.Contains("\""))
                            {
                                sRowData = sRowData.Replace("\"", "\\\"");
                            }

                            JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString() + "\":" + "\"" + sRowData + "\"");
                        }
                    }
                    if (i == ds.Tables[0].Rows.Count - 1)
                    {
                        JsonString.Append("}");
                    }
                    else
                    {
                        JsonString.Append("},");
                    }
                }
                JsonString.Append("]");
                return JsonString.ToString();
            }
            else
            {
                return null;
            }
        }

        public static string DataTableToJsonWithNewtonSoft(DataTable table)
        {
            string JSONstring = string.Empty;
            JSONstring = JsonConvert.SerializeObject(table);
            return JSONstring;
        }
        #endregion

        //#004
        /// <summary>
        /// LookupEdit에 DB데이터 바인딩
        /// 1. 테이블명이 REFFPF인 경우 columnName1에 코드값 컬럼명 columnName2에 ""
        /// 2. 테이블명이 REFFPF가 아닌 경우 columnName1에 코드값 컬럼명 columnName2에 보여줄 컬럼명
        /// </summary>
        /// <param name="lkup">룩업에디트</param>
        /// <param name="tableName">테이블명</param>
        /// <param name="columnName1"></param>
        /// <param name="columnName2"></param>
        public static void SetBoundLookUp(LookUpEdit lkup, string tableName, string columnName1, string columnName2)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" WITH ITEM_INFO AS ( ");

            strSql.AppendLine(" SELECT '' AS CD               ");
            strSql.AppendLine("      , '' AS NM               ");
            strSql.AppendLine("      , -1 AS SEQ              ");

            strSql.AppendLine(" UNION ALL                     ");

            if (tableName.Equals("COM_BASE_CD"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD                                 ");
                strSql.AppendLine("      , A.COM_NM AS NM                                 ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY A.SORT_SEQ) AS SEQ                                ");
                strSql.AppendLine("   FROM COM_BASE_CD A                                      ");
                strSql.AppendLine("  WHERE A.CD_GB = '" + columnName1 + "'               ");
            }
            else if (tableName.Equals("HR_EMP_BASIS"))
            {
                strSql.AppendLine(" SELECT EMP_ID AS CD                              ");
                strSql.AppendLine("      , EMP_NM AS NM                              ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("  FROM HR_EMP_BASIS                               ");
                strSql.AppendLine(" WHERE EMPL_GB = 'Y'                              ");
            }
            else
            {
                strSql.AppendLine(" SELECT A." + columnName1 + " AS CD                                 ");
                strSql.AppendLine("      , A." + columnName2 + " AS NM                                 ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY A." + columnName1 + ") AS SEQ    ");
                strSql.AppendLine("   FROM " + tableName + " A                                         ");
            }

            strSql.AppendLine(" )                            ");
            strSql.AppendLine(" SELECT CD, NM FROM ITEM_INFO ");
            strSql.AppendLine("  ORDER BY SEQ, CD            ");

            lkup.Properties.DataSource = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            lkup.Properties.ValueMember = "CD";
            lkup.Properties.DisplayMember = "NM";
            lkup.Properties.ShowHeader = false;
            lkup.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            lkup.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            lkup.Properties.ForceInitialize();
            lkup.Properties.PopulateColumns();
            lkup.Properties.Columns[0].Visible = false;
            lkup.Properties.UseDropDownRowsAsMaxCount = true;

            lkup.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Down)
                {
                    lkup.ShowPopup();
                }
            };

        }

        // 전체 GridLoopUpEdit Setting Method (공통)
        // 사용  예 : ComnTestFunc.SetBoundGridLookUp(RepoGLkupWLINE, "REFFPF", "WLINE", "");
        //            ComnTestFunc.SetBoundGridLookUp(RepoGLkupMACOD, "EQUIPF", "MACOD", "MANAM");
        // 파라미터 : 바인딩할 lookupedit명, 값이 있는 테이블 명, 코드값 컬럼명, 보여줄 컬럼명
        //            공통코드 테이블(REFFPF)의 경우 마지막 파라미터는 고려되지 않음. "" 추천.
        public static void SetBoundGridLookUp(RepositoryItemGridLookUpEdit repositoryLookUp, string tableName, string columnName1, string columnName2)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" WITH ITEM_INFO AS ( ");

            strSql.AppendLine(" SELECT '' AS CD               ");
            strSql.AppendLine("      , '' AS NM               ");
            strSql.AppendLine("      , -1 AS SEQ              ");

            strSql.AppendLine(" UNION ALL                     ");

            if (tableName.Equals("COM_BASE_CD"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD                                 ");
                strSql.AppendLine("      , A.COM_NM AS NM                                 ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY A.SORT_SEQ) AS SEQ                                ");
                strSql.AppendLine("   FROM COM_BASE_CD A                                      ");
                strSql.AppendLine("  WHERE A.CD_GB = '" + columnName1 + "'               ");
            }
            else if (tableName.Equals("HR_EMP_BASIS"))
            {
                strSql.AppendLine(" SELECT EMP_ID AS CD                              ");
                strSql.AppendLine("      , EMP_NM AS NM                              ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("  FROM HR_EMP_BASIS                               ");
                strSql.AppendLine(" WHERE EMPL_GB = 'Y'                              ");
            }
            else
            {
                strSql.AppendLine(" SELECT A." + columnName1 + " AS CD                                 ");
                strSql.AppendLine("      , A." + columnName2 + " AS NM                                 ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY A." + columnName1 + ") AS SEQ    ");
                strSql.AppendLine("   FROM " + tableName + " A                                         ");
            }

            strSql.AppendLine(" )                            ");
            strSql.AppendLine(" SELECT CD, NM FROM ITEM_INFO ");
            strSql.AppendLine("  ORDER BY SEQ                ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            dt.Columns["CD"].ColumnMapping = MappingType.Hidden; // CD 컬럼 안보이게 설정
            repositoryLookUp.DataSource = dt;
            repositoryLookUp.ValueMember = "CD";
            repositoryLookUp.DisplayMember = "NM";

            repositoryLookUp.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryLookUp.PopupFilterMode = PopupFilterMode.Contains;
            repositoryLookUp.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            repositoryLookUp.PopupView.OptionsFilter.AllowFilterIncrementalSearch = true;
            repositoryLookUp.ImmediatePopup = true;
            repositoryLookUp.View.OptionsView.ShowColumnHeaders = false;

            repositoryLookUp.PopupFormSize = new Size(0, dt.Rows.Count * 5); // 팝업 높이 설정
        }

        #region 암호화/복호화
        // 암호화 AES256
        public static string EncryptString(string InputText, string Password)
        {
            string EncryptedData = "";
            try
            {
                // Rihndael class를 선언하고, 초기화
                RijndaelManaged RijndaelCipher = new RijndaelManaged();

                RijndaelCipher.Padding = PaddingMode.PKCS7;
                RijndaelCipher.Mode = CipherMode.ECB;

                // 입력받은 문자열을 바이트 배열로 변환
                byte[] PlainText = System.Text.Encoding.Unicode.GetBytes(InputText);

                // 딕셔너리 공격을 대비해서 키를 더 풀기 어렵게 만들기 위해서 
                // Salt를 사용한다.
                byte[] Salt = Encoding.ASCII.GetBytes(Password.Length.ToString());

                // PasswordDeriveBytes 클래스를 사용해서 SecretKey를 얻는다.
                PasswordDeriveBytes SecretKey = new PasswordDeriveBytes(Password, Salt);

                // Create a encryptor from the existing SecretKey bytes.
                // encryptor 객체를 SecretKey로부터 만든다.
                // Secret Key에는 32바이트
                // (Rijndael의 디폴트인 256bit가 바로 32바이트입니다)를 사용하고, 
                // Initialization Vector로 16바이트
                // (역시 디폴트인 128비트가 바로 16바이트입니다)를 사용한다.
                ICryptoTransform Encryptor = RijndaelCipher.CreateEncryptor(SecretKey.GetBytes(32), SecretKey.GetBytes(16));

                // 메모리스트림 객체를 선언,초기화 
                MemoryStream memoryStream = new MemoryStream();

                // CryptoStream객체를 암호화된 데이터를 쓰기 위한 용도로 선언
                CryptoStream cryptoStream = new CryptoStream(memoryStream, Encryptor, CryptoStreamMode.Write);

                // 암호화 프로세스가 진행된다.
                cryptoStream.Write(PlainText, 0, PlainText.Length);

                // 암호화 종료
                cryptoStream.FlushFinalBlock();

                // 암호화된 데이터를 바이트 배열로 담는다.
                byte[] CipherBytes = memoryStream.ToArray();

                // 스트림 해제
                memoryStream.Close();
                cryptoStream.Close();

                // 암호화된 데이터를 Base64 인코딩된 문자열로 변환한다.
                EncryptedData = Convert.ToBase64String(CipherBytes);
            }
            catch { Console.Write("Faild Value"); }
            // 최종 결과를 리턴
            return EncryptedData;
        }

        // 복호화 AES256
        public static string DecryptString(string InputText, string Password)
        {
            string DecryptedData = "";   // 리턴값
            try
            {
                RijndaelManaged RijndaelCipher = new RijndaelManaged();

                RijndaelCipher.Padding = PaddingMode.PKCS7;
                RijndaelCipher.Mode = CipherMode.ECB;

                byte[] EncryptedData = Convert.FromBase64String(InputText);
                byte[] Salt = Encoding.ASCII.GetBytes(Password.Length.ToString());

                PasswordDeriveBytes SecretKey = new PasswordDeriveBytes(Password, Salt);

                // Decryptor 객체를 만든다.
                ICryptoTransform Decryptor = RijndaelCipher.CreateDecryptor(SecretKey.GetBytes(32), SecretKey.GetBytes(16));

                MemoryStream memoryStream = new MemoryStream(EncryptedData);

                // 데이터 읽기(복호화이므로) 용도로 cryptoStream객체를 선언, 초기화
                CryptoStream cryptoStream = new CryptoStream(memoryStream, Decryptor, CryptoStreamMode.Read);

                // 복호화된 데이터를 담을 바이트 배열을 선언한다.
                // 길이는 알 수 없지만, 일단 복호화되기 전의 데이터의 길이보다는
                // 길지 않을 것이기 때문에 그 길이로 선언한다.
                byte[] PlainText = new byte[EncryptedData.Length];

                // 복호화 시작
                int DecryptedCount = cryptoStream.Read(PlainText, 0, PlainText.Length);

                memoryStream.Close();
                cryptoStream.Close();

                // 복호화된 데이터를 문자열로 바꾼다.
                DecryptedData = Encoding.Unicode.GetString(PlainText, 0, DecryptedCount);

            }
            catch (Exception ex) { Console.Write(ex.Message); }
            // 최종 결과 리턴
            return DecryptedData;
        }

        /*
         * 패스워드는 웹 프로그램과 암복호화 맞춰야해서 아래 암복호화 사용
         */
        //암호화
        public static string Encrypt(string textToEncrypt, string key)
        {
            try
            {
                RijndaelManaged rijndaelCipher = new RijndaelManaged();

                rijndaelCipher.Mode = CipherMode.CBC;
                rijndaelCipher.Padding = PaddingMode.PKCS7;

                rijndaelCipher.KeySize = 128;
                rijndaelCipher.BlockSize = 128;

                byte[] pwdBytes = Encoding.UTF8.GetBytes(key);
                byte[] keyBytes = new byte[16];

                int len = pwdBytes.Length;
                if (len > keyBytes.Length)
                {
                    len = keyBytes.Length;
                }

                Array.Copy(pwdBytes, keyBytes, len);

                rijndaelCipher.Key = keyBytes;
                rijndaelCipher.IV = keyBytes;

                ICryptoTransform transform = rijndaelCipher.CreateEncryptor();

                byte[] plainText = Encoding.UTF8.GetBytes(textToEncrypt);

                return Convert.ToBase64String(transform.TransformFinalBlock(plainText, 0, plainText.Length));
            }
            catch
            {
                return null;
            }
        }

        //복호화
        public static string Decrypt(string textToDecrypt, string key)
        {
            try
            {
                RijndaelManaged rijndaelCipher = new RijndaelManaged();

                rijndaelCipher.Mode = CipherMode.CBC;
                rijndaelCipher.Padding = PaddingMode.PKCS7;

                rijndaelCipher.KeySize = 128;
                rijndaelCipher.BlockSize = 128;

                byte[] encryptedData = Convert.FromBase64String(textToDecrypt);
                byte[] pwdBytes = Encoding.UTF8.GetBytes(key);
                byte[] keyBytes = new byte[16];

                int len = pwdBytes.Length;

                if (len > keyBytes.Length)
                {
                    len = keyBytes.Length;
                }

                Array.Copy(pwdBytes, keyBytes, len);

                rijndaelCipher.Key = keyBytes;
                rijndaelCipher.IV = keyBytes;

                byte[] plainText = rijndaelCipher.CreateDecryptor().TransformFinalBlock(encryptedData, 0, encryptedData.Length);

                return Encoding.UTF8.GetString(plainText);
            }
            catch
            {
                return null;
            }
        }
        #endregion
    }

    public class FaxInfo
    {
        public string CorpNo;
        public string LinkID;
        public string SECRET_KEY;
        public string USER_ID; //일반회원ID
        public string SYNC_USER_ID; //연동회원ID

        public FaxInfo()
        {
            CorpNo = "6068188502";
            LinkID = "LST";
            SECRET_KEY = "M0XgPfNhJ0+c0KKTgjtgme9oT445+A0BYQiKZqzYAq0=";
            USER_ID = "daejist";
            //SYNC_USER_ID = "LST0821";
            SYNC_USER_ID = "daejist";
        }

        public string GetState(int state)
        {
            string sResult = "??";

            if(state == 0)
            {
                sResult = "접수";
            }
            else if(state == 1)
            {
                sResult = "변환중";
            }
            else if (state == 2)
            {
                sResult = "전송중";
            }
            else if (state == 3)
            {
                sResult = "처리완료";
            }
            else if (state == 4)
            {
                sResult = "취소";
            }

            return sResult;
        }

        public string GetResult(int state)
        {
            string sResult = "??";

            if (state == 100)
            {
                sResult = "전송성공";
            }
            else if (state == 300)
            {
                sResult = "발신번호 형식 오류";
            }
            else if (state == 301)
            {
                sResult = "발신번호 사전 미등록";
            }
            else if (state == 402)
            {
                sResult = "변환 실패";
            }
            else if (state == 405)
            {
                sResult = "비지원문서";
            }
            else if (state == 408)
            {
                sResult = "파일명오류(특수문자)";
            }
            else if (state == 413)
            {
                sResult = "지원장수초과";
            }
            else if (state == 414)
            {
                sResult = "팩스 변환실패";
            }
            else if (state == 415)
            {
                sResult = "팩스 변환 시간초과 - cron from db";
            }
            else if (state == 416)
            {
                sResult = "팩스파일 처리중 오류";
            }
            else if (state == 498)
            {
                sResult = "기타오류로 인한 변환실패";
            }
            else if (state == 499)
            {
                sResult = "변환후 과금실패";
            }
            else if (state == 502)
            {
                sResult = "부분완료";
            }
            else if (state == 503)
            {
                sResult = "통화중 (3번 재시도 후 최종결과)";
            }
            else if (state == 504)
            {
                sResult = "잘못된 수신번호";
            }
            else if (state == 505)
            {
                sResult = "응답 없음(받지않거나, 60초 초과)";
            }
            else if (state == 506)
            {
                sResult = "수화기 들었음 - 환불안됨";
            }
            else if (state == 507)
            {
                sResult = "수신 거부";
            }
            else if (state == 508)
            {
                sResult = "콜 혼잡";
            }
            else if (state == 509)
            {
                sResult = "차단요청번호";
            }
            else if (state == 511)
            {
                sResult = "중복번호";
            }
            else if (state == 512)
            {
                sResult = "시간초과";
            }
            else if (state == 516)
            {
                sResult = "팩스 기기 아님";
            }
            else if (state == 517)
            {
                sResult = "보이스 감지(ARS 또는 음성사서함) - 환불 안됨";
            }
            else if (state == 518)
            {
                sResult = "발신번호 오류";
            }
            else if (state == 800)
            {
                sResult = "080 수신거부 대상";
            }
            else if (state == 802)
            {
                sResult = "통합 수신거부 대상";
            }
            else if (state == 999)
            {
                sResult = "기타";
            }

            return sResult;
        }
    }

    public class SkinInfo
    {
        private string _USRCD;

        public SkinInfo()
        {

        }

        public SkinInfo(string sUsrCd)
        {
            this._USRCD = sUsrCd;
        }

        public void SetProgramSkin()
        {
            string sSKIN_NAME = GetSkinInfo();

            if (sSKIN_NAME.Equals("Visual Studio 2010"))
            {
                sSKIN_NAME = "VS2010";
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(sSKIN_NAME);
            }
            else if (sSKIN_NAME.Equals("Office 2013 White"))
            {
                sSKIN_NAME = "Office 2013";
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(sSKIN_NAME);
            }
            else if (sSKIN_NAME.Equals("McSkin"))
            {
                DevExpress.UserSkins.BonusSkins.Register();
                DevExpress.Skins.SkinManager.EnableFormSkins();
                DevExpress.LookAndFeel.UserLookAndFeel.Default.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Skin;
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName = "McSkin";
            }
            else
            {
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(sSKIN_NAME);
            }
        }

        private string GetSkinInfo()
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("USRCD", _USRCD);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("  ");
            strSql.AppendLine(" SELECT SKINNAME  ");
            strSql.AppendLine("   FROM ZUSRSKIN ");
            strSql.AppendLine("  WHERE USRCD = @USRCD ");

            DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
            string sSKIN_NAME = string.Empty;
            if (dt.Rows.Count > 0)
            {
                sSKIN_NAME = dt.Rows[0]["SKINNAME"]?.ToString();
            }

            if (string.IsNullOrEmpty(sSKIN_NAME))
            {
                sSKIN_NAME = "DevExpress Style";
            }

            return sSKIN_NAME;
        }
    }

    public class FileInfo_1
    {
        private string _FileName = "";
        private readonly string _Default_FTP_Path = @"ERP/Docs";
        private readonly string _Local_ETC_File_Path = string.Format(@"{0}\Format_File", Application.StartupPath);
        private readonly string _Local_ETC_File_Temp_Path = string.Format(@"{0}\Temp_File", Application.StartupPath);
        private Dictionary<string, string> _dicFiles = new Dictionary<string, string>();

        public FileInfo_1(string GB)
        {
            SetFileInfo();

            try
            {
                _FileName = _dicFiles[GB];
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("[Logic Error] FileInfo Class 'GB' 변수를 올바르게 세팅하세요");
            }
        }

        private void SetFileInfo()
        {
            _dicFiles.Add("1", "주간검수보고.xlsx");
            _dicFiles.Add("2", "월간검수보고.xlsx");
            //#001
            _dicFiles.Add("3", "영업실적목표.xlsx");
            _dicFiles.Add("4", "재고장.xlsx");
            _dicFiles.Add("5", "일일매입출현황(재고).xls");
            _dicFiles.Add("6", "영업일보.xlsx");
            _dicFiles.Add("7", "구매팀 실적.xlsx");
            _dicFiles.Add("8", "업무보고.xlsx");
        }

        private byte[] GetLocalFile()
        {
            byte[] byResult = null;
            try
            {
                string sFilePath = string.Format(@"{0}\{1}", _Local_ETC_File_Path, _FileName);
                FileInfo file = new FileInfo(sFilePath);
                if (file.Exists)
                {
                    byResult = File.ReadAllBytes(sFilePath);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("[GetLocalFile]" + ex.Message);
            }

            return byResult;
        }

        private byte[] GetFTP_File()
        {
            byte[] byResult = null;
            try
            {
                string sFTP_PATH = ComnEtcFunc.FTP_IP;

                //string ftpPath = string.Format(@"ftp://{0}/{1}\{2}", ComnEtcFunc.FTP_IP, _Default_FTP_Path, _FileName);
                //string ftpPath = @"ftp://192.168.0.202/" + _Default_FTP_Path + "/" + _FileName;
                string ftpPath = @"ftp://"+ ComnEtcFunc.FTP_IP + "/" + _Default_FTP_Path + "/" + _FileName;
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                
                try
                {
                    byResult = ComnEtcFunc.DownloadFTPFile_ByteArray(ftpPath, user, pw);
                }
                catch (Exception ex)
                {

                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show("[GetFTP_File]" + ex.Message);
            }

            return byResult;
        }

        public string[] CheckFileInfo()
        {
            string[] str = new string[2];

            bool bFTP_Exsits = true;
            byte[] by = GetFTP_File();
            if(by == null)
            {
                bFTP_Exsits = false;
            }

            bool bLocal_Exsits = true;
            if (!bFTP_Exsits)
            {
                by = GetLocalFile();
            }

            if (by == null)
                bLocal_Exsits = false;

            if (!bLocal_Exsits)
            {
                XtraMessageBox.Show("FTP 및 로컬에 어떠한 양식파일이 존재하지 않습니다.");
                return null;
            }
            
            DirectoryInfo directoryInfo = new DirectoryInfo(_Local_ETC_File_Path);
            if (!directoryInfo.Exists)
                directoryInfo.Create();

            string sPath = string.Format(@"{0}\{1}", _Local_ETC_File_Path, _FileName);
            File.WriteAllBytes(sPath, by);

            str[0] = sPath;

            DirectoryInfo directoryInfo_Temp = new DirectoryInfo(_Local_ETC_File_Temp_Path);
            if (!directoryInfo_Temp.Exists)
                directoryInfo_Temp.Create();

            string sPath_temp = string.Format(@"{0}\{1}", _Local_ETC_File_Temp_Path, _FileName);
            //File.WriteAllBytes(sPath_temp, by);
            str[1] = sPath_temp;

            return str;
        }
    }

    //public class SeparateSheetsExportHelper
    //{
    //    public SeparateSheetsExportHelper()
    //    {
    //    }
    //    internal static void Export(string fileName, XlsExportOptions xlsExportOptions)
    //    {
    //        XtraReport report = new XtraReport();
    //        DataSet ds = new nwindDataSet();
    //        DataSetTableAdapters.CategoriesTableAdapter cta = new WindowsApplication1.nwindDataSetTableAdapters.CategoriesTableAdapter();
    //        cta.Fill(ds.Categories);

    //        report.Category.Value = ds.Categories[0].CategoryID;
    //        report.PrintingSystem.ExportOptions.PrintPreview.ShowOptionsBeforeExport = false;
    //        report.CreateDocument();
    //        report.ExportToXls(fileName, xlsExportOptions);

    //        Excel.Application ObjExcel = new Excel.Application();
    //        Workbook ObjWorkBookGeneral;
    //        Worksheet ObjWorkSheet;
    //        Workbook ObjWorkBookTemp = ObjExcel.Workbooks.Open(Environment.CurrentDirectory + "\\" + fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
    //        ObjWorkBookGeneral = ObjExcel.Workbooks.Open(Environment.CurrentDirectory + "\\" + fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
    //        ((Worksheet)ObjWorkBookGeneral.Sheets[1]).Name = ds.Categories[0].Description.Length > 30 ? ds.Categories[0].Description.Substring(0, 30) : ds.Categories[0].Description;

    //        bool isObjWorkBookTempClosed = false;
    //        try
    //        {
    //            for (int i = 1; i < ds.Categories.Count; i++)
    //            {
    //                report.Category.Value = ds.Categories[i].CategoryID;
    //                report.CreateDocument();
    //                report.ExportToXls("temp_" + fileName + ".xls");
    //                ObjWorkBookTemp = ObjExcel.Workbooks.Open(Environment.CurrentDirectory + "\\temp_" + fileName + ".xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
    //                isObjWorkBookTempClosed = false;
    //                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookTemp.Sheets[1];
    //                ObjWorkSheet.Name = ds.Categories[i].Description.Length > 30 ? ds.Categories[i].Description.Substring(0, 30) : ds.Categories[i].Description;
    //                ObjWorkSheet.Copy(Type.Missing, ObjWorkBookGeneral.Sheets[ObjWorkBookGeneral.Sheets.Count]);
    //                ObjWorkBookTemp.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
    //                isObjWorkBookTempClosed = true;
    //            }
    //        }
    //        finally
    //        {
    //            if (!isObjWorkBookTempClosed)
    //                ObjWorkBookTemp.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
    //            ObjWorkBookGeneral.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
    //            ObjExcel.Quit();
    //            File.Delete(Environment.CurrentDirectory + "\\temp_" + fileName + ".xls");
    //        }
    //    }
    //}
}
