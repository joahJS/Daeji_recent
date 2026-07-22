using DevExpress.Utils.OAuth;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraLayout;
using Newtonsoft.Json;
using Daeji_MONITERING;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

class ComnMethod
{
    public static Dictionary<string, string> parameterDic = new Dictionary<string, string>();
    public static List<Parameter> parameterList = new List<Parameter>();
    
    #region [로그인]
    //public static bool Login(string ID, string PW)
    //{
    //    bool isOK = false;
    //    parameterDic.Clear();
    //    parameterDic.Add("USRID", ID);
    //    parameterDic.Add("USRPW", PW);
    //    DataTable dt = DBConn.GetDataTable(procedureDic[PROCEDURE_NAME.로그인], parameterDic);
    //    if (dt != null)
    //    {
    //        if (dt.Rows.Count > 0)
    //        {
    //            isOK = true;
    //            LoginUser.USRID = dt.Rows[0]["USRID"]?.ToString();
    //            LoginUser.USRNM = dt.Rows[0]["USRNM"]?.ToString();
    //            LoginUser.USRCD = dt.Rows[0]["USRCD"]?.ToString();
    //        }
    //    }
    //    return isOK;
    //}
    #endregion                
                            
    #region [LOG 기록]
    //public enum CONNECT_TYPE { 접속, 종료, 조회, 등록, 수정, 저장, 삭제, 엑셀, 출력, 접속실패 }
    //public static void SetLogInfo(string ProgramID, CONNECT_TYPE CON)
    //{
    //    parameterDic.Clear();
    //    parameterDic.Add("LOGGU", Enum.GetName(typeof(CONNECT_TYPE), CON)?.ToString());
    //    parameterDic.Add("USRCD", LoginUser.USRCD);
    //    parameterDic.Add("PGMID", ProgramID);
    //    parameterDic.Add("COMIP", DBConn.Client_IP);
    //    parameterDic.Add("EXENM", string.Format(@"{0}\{1}", Application.StartupPath, "SG_MONITERING.exe"));
    //    DBConn.ExecuteNonQuery(procedureDic[PROCEDURE_NAME.로그저장], parameterDic);
    //}
    //public static void SetLogInfo(string ProgramID, string ProgramName, CONNECT_TYPE CON)
    //{
    //    parameterDic.Clear();
    //    parameterDic.Add("LOGGU", Enum.GetName(typeof(CONNECT_TYPE), CON)?.ToString());
    //    parameterDic.Add("USRCD", LoginUser.USRCD);
    //    parameterDic.Add("PGMID", ProgramID);
    //    parameterDic.Add("PGMNM", ProgramName);
    //    parameterDic.Add("COMIP", DBConn.Client_IP);
    //    parameterDic.Add("EXENM", string.Format(@"{0}\{1}", Application.StartupPath, "SG_MONITERING.exe"));
    //    DBConn.ExecuteNonQuery(procedureDic[PROCEDURE_NAME.로그저장], parameterDic);
    //}
    //public static void SetLogInfo(string ProgramID, string ProgramName, string LogName)
    //{
    //    parameterDic.Clear();
    //    parameterDic.Add("LOGGU", LogName);
    //    parameterDic.Add("USRCD", LoginUser.USRCD);
    //    parameterDic.Add("PGMID", ProgramID);
    //    parameterDic.Add("PGMNM", ProgramName);
    //    parameterDic.Add("COMIP", DBConn.Client_IP);
    //    parameterDic.Add("EXENM", string.Format(@"{0}\{1}", Application.StartupPath, "SG_MONITERING.exe"));
    //    DBConn.ExecuteNonQuery(procedureDic[PROCEDURE_NAME.로그저장], parameterDic);
    //}
    #endregion

    #region [INI 저장 관련]
    public enum INI_NAME { USER, VERSION }
    public static Dictionary<INI_NAME, string> IniNameDic = new Dictionary<INI_NAME, string>
    {
        { INI_NAME.USER, "User.ini" },
        { INI_NAME.VERSION, "Version.ini" }
    };

    [DllImport("kernel32")]
    private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
    [DllImport("kernel32")]
    private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

    public static String GetIniValue(INI_NAME ini, String Section, String Key)
    {
        string path = string.Format(@"{0}\{1}", Application.StartupPath, IniNameDic[ini]);
        StringBuilder temp = new StringBuilder(255);
        int i = GetPrivateProfileString(Section, Key, "", temp, 255, path);
        return temp.ToString();
    }
    public static void SetIniValue(INI_NAME ini, String Section, String Key, String Value)
    {
        string path = string.Format(@"{0}\{1}", Application.StartupPath, IniNameDic[ini]);
        WritePrivateProfileString(Section, Key, Value, path);
    }
    #endregion         
                           
    #region [아이콘 & 메인이미지]
    public static Icon GetFavicon()
    {
        Bitmap bitmap = null;
        Icon icon = null;
        //if (BIZ_TYPE.Equals("ERCT"))
        //    bitmap = Properties.Resources.ERCT_FAVICON;
        //else if (BIZ_TYPE.Equals("DH"))
        //    bitmap = Properties.Resources.DH_FAVICON;
        //else if (BIZ_TYPE.Equals("SH"))
        //    bitmap = Properties.Resources.SH_FAVICON;
        //else if (BIZ_TYPE.Equals("DS"))
        //    bitmap = Properties.Resources.DS_FAVICON;

        if (bitmap != null)
            icon = Icon.FromHandle(bitmap.GetHicon());

        return icon;
    }

    public static Image GetMainImage()
    {
        Image img = null;
        Bitmap bitmap = null;
        //if (BIZ_TYPE.Equals("ERCT"))
        //    bitmap = Properties.Resources.ERCT_MAIN_IMG;
        //else if (BIZ_TYPE.Equals("DH"))
        //    bitmap = Properties.Resources.DH_MAIN_IMG;
        //else if (BIZ_TYPE.Equals("SH"))
        //    bitmap = Properties.Resources.SH_MAIN_IMG;
        //else if (BIZ_TYPE.Equals("DS"))
        //    bitmap = Properties.Resources.DS_MAIN_IMG;
        if (bitmap != null)
        {
            img = (Image)bitmap;
        }
        return img;
    }
    #endregion

    #region [레이아웃 저장 & 초기화]
    //public static void SetFormText(Form form)
    //{
    //    string sFormText = ComnString.SAVE_LAYOUT_LOADING_NAME;
    //    form.Text += sFormText;
    //}
    //public static void SaveLayout(string sFormName, GridView[] arrGrdView, LayoutControl layout)
    //{
    //    //GridView 저장
    //    for (int i = 0; i < arrGrdView.Length; i++)
    //    {
    //        string sFileName = string.Format("{0}{1}", sFormName, (i + 1).ToString());
    //        SaveGridViewLayout(LoginUser.USRID, sFileName, arrGrdView[i]);
    //    }
    //    //Layout 저장
    //    SaveGridViewLayout(LoginUser.USRID, sFormName, layout);
    //}
    //public static void SetLayout(string sFormName, GridView[] arrGrdView, LayoutControl layout)
    //{
    //    for (int i = 0; i < arrGrdView.Length; i++)
    //    {
    //        string sFileName = string.Format("{0}{1}", sFormName, (i + 1).ToString());
    //        SetGridViewLayout(LoginUser.USRID, "AccAdm", sFileName, arrGrdView[i]);
    //    }
    //}
    //public static void SaveGridViewLayout(string sId, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
    //{
    //    string path = Application.StartupPath + @"\xaml\" + sId;
    //    if (!Directory.Exists(path))
    //    {
    //        Directory.CreateDirectory(path);
    //    }
    //    string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
    //    view.SaveLayoutToXml(sFile);
    //}
    //public static void SaveGridViewLayout(string sId, string sClass, LayoutControl layout)
    //{
    //    string path = Application.StartupPath + @"\xaml\" + sId;
    //    if (!Directory.Exists(path))
    //    {
    //        Directory.CreateDirectory(path);
    //    }
    //    string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + "_Layout.xaml";
    //    layout.SaveLayoutToXml(sFile);
    //}
    //public static void SetGridViewLayout(string sId, string sProject, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
    //{
    //    string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
    //    if (!File.Exists(sFile))
    //    {
    //        sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
    //        if (!File.Exists(sFile)) return;
    //    }
    //    view.RestoreLayoutFromXml(sFile);
    //}
    //public static void SetGridViewLayout(string sId, string sProject, string sClass, LayoutControl layout)
    //{
    //    string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + "_Layout.xaml";
    //    if (!File.Exists(sFile))
    //    {
    //        sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + "_Layout.xaml";
    //        if (!File.Exists(sFile)) return;
    //    }
    //    layout.RestoreLayoutFromXml(sFile);
    //}
    //public static Form GetAssemblyForm(string strFormName)
    //{
    //    Form f = null;
    //    foreach (Type t in System.Reflection.Assembly.GetExecutingAssembly().GetTypes())
    //    {
    //        if (t.Name == strFormName)                     //프로젝트 내 폼 중에서 찾을 이름과 같으면...
    //        {
    //            object o = Activator.CreateInstance(t);    //인스턴스 개체 생성 
    //            f = o as Form;                                  //인스턴스 개체 폼 형식으로 캐스팅
    //        }
    //    }
    //    return f;
    //}
    #endregion

    #region [DateEdit 설정 관련 - 날짜 세팅, ShowClear = false + 날짜 체크(빈값)]
    public static void SetDateFromFromValue(DateEdit ymdFrom, DateEdit ymdTo)
    {
        DateTime today = DateTime.Now.Date;
        ymdFrom.EditValue = today;
        ymdTo.EditValue = today;
        ymdFrom.Properties.ShowClear = false;
        ymdTo.Properties.ShowClear = false;
    }
    public static void SetDateFromToValue(DateEdit ymdFrom, DateEdit ymdTo)
    {
        DateTime today = DateTime.Now.Date;
        ymdFrom.EditValue = today.AddDays(1 - today.Day);
        ymdTo.EditValue = today;
        ymdFrom.Properties.ShowClear = false;
        ymdTo.Properties.ShowClear = false;
    }
    public static void SetDateFromToValueThisYear(DateEdit ymdFrom, DateEdit ymdTo)
    {
        int year = DateTime.Now.Year;
        DateTime today = DateTime.Now.Date;
        DateTime firstDay = new DateTime(year, 1, 1);
        ymdFrom.EditValue = firstDay;
        ymdTo.EditValue = today;
        ymdFrom.Properties.ShowClear = false;
        ymdTo.Properties.ShowClear = false;
    }
    public static void SetDateFromValue(DateEdit ymdFrom)
    {
        DateTime today = DateTime.Now.Date;
        ymdFrom.EditValue = today.AddDays(1 - today.Day);
        ymdFrom.Properties.ShowClear = false;
    }
    public static void SetDateToValue(DateEdit ymdTo)
    {
        DateTime today = DateTime.Now.Date;
        ymdTo.EditValue = today;
        ymdTo.Properties.ShowClear = false;
    }
    public static void SetDateMonthValue(DateEdit ymdFrom)
    {
        DateTime today = DateTime.Now.Date;
        ymdFrom.EditValue = today.AddDays(1 - today.Month);
        ymdFrom.Properties.ShowClear = false;
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
    #endregion

    #region [현재 미사용]

    #region [버튼 관련]
    public enum BUTTON_NAME { 조회, 추가, 수정, 저장, 삭제, 출력, 엑셀, 닫기 }
    public static Dictionary<BUTTON_NAME, string> buttonNameDic = new Dictionary<BUTTON_NAME, string>
    {
        { BUTTON_NAME.조회, "조회 (F5)" }
    };
    public static Dictionary<BUTTON_NAME, string> buttonCodeDic = new Dictionary<BUTTON_NAME, string>
    {
        { BUTTON_NAME.조회, "bt_Retr" }
    };

    /// <summary>
    /// 버튼 종류에 따른 기본 세팅? - 활용방안에 의문
    /// </summary>
    /// <param name="simpleButton"></param>
    /// <param name="BTN"></param>
    public static void SetComponent(SimpleButton simpleButton, BUTTON_NAME BTN)
    {
        simpleButton.Name = buttonCodeDic[BTN];
        simpleButton.Size = new System.Drawing.Size(100, 30);
        simpleButton.Text = buttonNameDic[BTN];
    }
    #endregion






    #endregion
    
    #region [Lookup 관련]
    public enum LookUpName { 사용자, 거래처, 제품, 공통코드 }
    public static void SetLookUp(LookUpEdit lkup, LookUpName lookUpName, string RCDTP)
    {
        if (lookUpName.Equals(LookUpName.사용자))
        {
            ComnEtcFunc.SetBoundLookUp(lkup, "zUSRLST", "USRCD", "USRNM");
        }
        else if (lookUpName.Equals(LookUpName.거래처))
        {
            ComnEtcFunc.SetBoundLookUp(lkup, "CVMAST", "CVCOD", "CVNAM");
        }
        else if (lookUpName.Equals(LookUpName.제품))
        {
            ComnEtcFunc.SetBoundLookUp(lkup, "ITEMAS", "ITCOD", "ITNAM");
        }
        else if (lookUpName.Equals(LookUpName.공통코드))
        {
            ComnEtcFunc.SetBoundLookUp(lkup, "REFFPF", RCDTP, "");
        }
    }
    public static void SetGridLookUp(RepositoryItemGridLookUpEdit lkup, LookUpName lookUpName, string RCDTP)
    {
        if (lookUpName.Equals(LookUpName.사용자))
        {
            ComnEtcFunc.SetBoundGridLookUp(lkup, "zUSRLST", "USRCD", "USRNM");
        }
        else if (lookUpName.Equals(LookUpName.거래처))
        {
            ComnEtcFunc.SetBoundGridLookUp(lkup, "CVMAST", "CVCOD", "CVNAM");
        }
        else if (lookUpName.Equals(LookUpName.제품))
        {
            ComnEtcFunc.SetBoundGridLookUp(lkup, "ITEMAS", "ITCOD", "ITNAM");
        }
        else if (lookUpName.Equals(LookUpName.공통코드))
        {
            ComnEtcFunc.SetBoundGridLookUp(lkup, "REFFPF", RCDTP, "");
        }
    }
    public static void SetBoundLookUp_Mall(LookUpEdit lkup)
    {
        StringBuilder strSql = new StringBuilder();
        strSql.Clear();
        strSql.AppendLine(" ");
        strSql.AppendLine(" WITH ITEM_INFO AS (                                 ");
        strSql.AppendLine(" SELECT '' AS CD                                     ");
        strSql.AppendLine("      , '' AS NM                                     ");
        strSql.AppendLine("      , -1 AS SEQ                                    ");
        strSql.AppendLine(" UNION ALL                                           ");
        strSql.AppendLine(" SELECT A.CVCOD AS CD                                ");
        strSql.AppendLine("      , A.CVNAM AS NM                                ");
        strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY A.CVCOD) AS SEQ   ");
        strSql.AppendLine("   FROM CVMAST A                                     ");
        strSql.AppendLine("  WHERE A.CVGU = '5'                                 ");
        strSql.AppendLine(" )                                                   ");
        strSql.AppendLine(" SELECT CD, NM FROM ITEM_INFO                        ");
        strSql.AppendLine("  ORDER BY SEQ, CD                                   ");

        lkup.Properties.DataSource = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        lkup.Properties.ValueMember = "CD";
        lkup.Properties.DisplayMember = "NM";
        lkup.Properties.ShowHeader = false;
        lkup.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
        lkup.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
        lkup.ItemIndex = 1;

    }
    #endregion

    #region [JSON 관련]
    public static string DataTableToJsonWithNewtonSoft(DataTable table)
    {
        string JSONstring = string.Empty;
        JSONstring = JsonConvert.SerializeObject(table);
        return JSONstring;
    } 
    #endregion

    #region [FTP 업로드-다운로드]
    /// <summary>
    /// FT{ 디렉토리에 같은 이름의 파일 존재 유무 검사 메서드
    /// </summary>
    /// <param name="addr">FTP 디렉토리 경로</param>
    /// <param name="FtpId">FTP ID</param>
    /// <param name="FtpPw">FTP PASS</param>
    /// <returns>FTP 파일존재 여부</returns>
    public static bool FtpFileExists(string addr, string FtpId, string FtpPw)
    {
        bool IsExists = true;
        FtpWebRequest reqFTP = null;
        FtpWebResponse respFTP = null;

        try
        {
            reqFTP = (FtpWebRequest)WebRequest.Create(addr);
            reqFTP.Credentials = new NetworkCredential(FtpId, FtpPw);
            reqFTP.Method = WebRequestMethods.Ftp.GetFileSize;
            respFTP = (FtpWebResponse)reqFTP.GetResponse();

            if (respFTP.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
            {
                IsExists = false;
            }
        }
        catch
        {
            IsExists = false;
        }
        finally
        {
            if (reqFTP != null)
            {
                reqFTP = null;
            }
            if (respFTP != null)
            {
                respFTP = null;
            }
        }
        return IsExists;
    }

    public static bool FTP_CheckDir(string dir)
    {
        string str = string.Empty;
        bool _isExist = false;

        try
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(dir);
            request.Credentials = new NetworkCredential(ComnString.FTP_ID, ComnString.FTP_PW);
            request.Method = WebRequestMethods.Ftp.ListDirectory;
            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                response.Close();
                _isExist = true;
            }
        }
        catch (WebException e)
        {
            FtpWebResponse response = (FtpWebResponse)e.Response;
            if (response.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
            {
                str = "폴더 없음";
            }
            else if (e.Status == WebExceptionStatus.ProtocolError)
            {
                str = string.Format("상태코드 : {0}", ((HttpWebResponse)e.Response).StatusCode);
                str += string.Format("\r\n상태설명 : {0}", ((HttpWebResponse)e.Response).StatusDescription);
            }
            else
            {
                str = "Error : " + e.Message;
            }
            return _isExist;
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show(ex.Message);
        }
        return _isExist;
    }

    public static void FTP_MakeDir(string dir)
    {
        try
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(dir);
            request.Credentials = new NetworkCredential(ComnString.FTP_ID, ComnString.FTP_PW);
            request.Method = WebRequestMethods.Ftp.MakeDirectory;
            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                response.Close();
            }
        }
        catch (WebException ex)
        {
            throw ex;
        }
    }


    /// <summary>
    /// 저장 시에 호출하는 FTP 업로드 메서드
    /// </summary>
    /// <param name="paths"></param>
    /// <param name="DIR"></param>
    /// <param name="SLINO"></param>
    public static void FTP_Upload_Files(Dictionary<string, string[]> paths, string DIR, string SLINO)
    {
        try
        {
            string FTP_DIR = DIR + @"/" + SLINO;
            if (!ComnMethod.FTP_CheckDir(FTP_DIR))
            {
                ComnMethod.FTP_MakeDir(FTP_DIR);
            }
            foreach (string[] path in paths.Values)
            {
                for (int i = 0; i < path.Length; i++)
                {
                    string sFileName = Path.GetFileName(path[i]);
                    //Read the contents of the file into a stream
                    var fileStream = new FileStream(path[i], FileMode.Open);
                    byte[] FILE = new byte[fileStream.Length];
                    fileStream.Read(FILE, 0, FILE.Length);
                    string FtpFilePath = FTP_DIR + @"/" + sFileName;
                    if (ComnMethod.FtpFileExists(FtpFilePath, ComnString.FTP_ID, ComnString.FTP_PW))
                    {
                        if (XtraMessageBox.Show("이미 있는 파일입니다. 덮어쓸까요?" + "\r\n파일명 : " + sFileName
                            , "중복 파일", MessageBoxButtons.YesNo) != DialogResult.Yes)
                        {
                            continue;
                        }
                    }
                    ComnMethod.FTPUpload(FtpFilePath, ComnString.FTP_ID, ComnString.FTP_PW, FILE);
                }
            }
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show(ex.Message);
            return;
        }
    }

    public static void FTP_Download_Files(string dir, string[] fileArr)
    {
        //try
        //{
        //    using (XtraFolderBrowserDialog dialog = new XtraFolderBrowserDialog())
        //    {
        //        dialog.Title = "저장할 폴더 선택";
        //        if (dialog.ShowDialog() == DialogResult.OK)
        //        {
        //            string FTP_dir = string.Empty;
        //            string COM_dir = string.Empty;
        //            for (int i = 0; i < fileArr.Length; i++)
        //            {
        //                FTP_dir = dir + @"/" + fileArr[i];
        //                COM_dir = dialog.SelectedPath + @"/" + fileArr[i];

        //                // WebRequest.Create로 Http,Ftp,File Request 객체를 모두 생성할 수 있다.
        //                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(FTP_dir);
        //                request.Method = WebRequestMethods.Ftp.DownloadFile;
        //                request.Credentials = new NetworkCredential(ComnString.FTP_ID, ComnString.FTP_PW);
        //                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
        //                {
        //                    Stream stream = response.GetResponseStream();
        //                    string data; // 결과값 문자열 (Binary로도 읽을 수 있다)                
        //                    using (StreamReader reader = new StreamReader(stream))
        //                    {
        //                        data = reader.ReadToEnd();
        //                    }
        //                    File.WriteAllText(COM_dir, data);
        //                }
        //            }
        //            XtraMessageBox.Show("저장되었습니다.");
        //        }
        //    }
        //}
        //catch (Exception ex)
        //{
        //    XtraMessageBox.Show(ex.Message);
        //}
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="directoryPath">FTP 경로</param>
    /// <param name="_FTPuserID">ID</param>
    /// <param name="_FTPpassword">PASS</param>
    /// <param name="data">DATA</param>
    public static void FTPUpload(string directoryPath, string _FTPuserID, string _FTPpassword, byte[] data)
    {
        //업로드 위한 설정
        try
        {
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(directoryPath);
            req.Method = WebRequestMethods.Ftp.UploadFile;
            req.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            req.ContentLength = data.Length;

            // RequestStream에 데이타를 쓴다
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();
            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// FTP 경로의 디렉토리를 점검하고 없으면 생성
    /// </summary>
    /// <param name="directoryPath">FTP 디렉토리 경로</param>
    /// <param name="_FTPuserID">FTP ID</param>
    /// <param name="_FTPpassword">FTP PASS</param>
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

    /// <summary>
    /// FTP 디렉토리 존재 여부 확인
    /// </summary>
    /// <param name="Directory">FTP 디렉토리 경로</param>
    /// <param name="_FTPuserID">FTP ID</param>
    /// <param name="_FTPpassword">FTP PASS</param>
    /// <returns>경로 존재 여부</returns>
    public static bool IsExistDirectory(string Directory, string _FTPuserID, string _FTPpassword)
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

    /// <summary>
    /// FTP 디렉토리 생성
    /// </summary>
    /// <param name="Directory">FTP 디렉토리 경로</param>
    /// <param name="_FTPuserID">FTP ID</param>
    /// <param name="_FTPpassword">FTP PASS</param>
    /// <returns></returns>
    public static bool MakeDirectory(string Directory, string _FTPuserID, string _FTPpassword)
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
    public static string GetStringResponse(FtpWebRequest ftp)
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
    #endregion

}