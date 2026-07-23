using DevExpress.XtraBars;
using DevExpress.XtraBars.Navigation;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraLayout;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Net.NetworkInformation;
using Newtonsoft.Json.Linq;
/*
* 작성일자 :
*   작성자 :
* ---------------------HISTORY-----------------------
* 수정일자 : 
*   수정자 : 
*   참조ID : 
* 수정내용 : 
*/
namespace Daeji_MONITERING
{
    class ComnEtcFunc
    {
        public static StringBuilder strSql = new StringBuilder();
        public static string EXS_ID = "";
        public static string V_CHK = "";
        public static string BIZ_TYPE = "PRMS";
        public static string VERSION = "";

        public static string FTP_IP = "112.216.189.178:8054";
        public static string FTP_PW = "pineit0401";
        public static string FTP_ID = "steelnet";

        //API 정보
        public static string _URI = "https://log.smart-factory.kr/apisvc/sendLogDataJSON.do?logData=";
        public static string _CRTKEY = "$5$API$IScg.UVhMAATjWXSKsfwBFPSOniSGCWwYEybmHhPnVD";
        public static string _LOGDT = "2021-10-29 16:15:01.778";
        public static string _USESE = "조회";
        public static string _SYSUSER = "소나무";
        public static string _CONECTIP = "49.166.208.48";
        public static string _DATAUSGQTY = "0";

        #region [공용 DataRow]
        // 공용 폼에서 사용하는 DataRow 속성
        public static DataRow row { get; set; }
        #endregion


        public static Icon GetFavicon()
        {
            Bitmap bitmap = null;
            Icon icon = null;
            if (BIZ_TYPE.Equals("PRMS"))
                //bitmap = Properties.Resources.QS_MAIN;

            if (bitmap != null)
                icon = Icon.FromHandle(bitmap.GetHicon());

            return icon;
        }

        public static Image GetMainImage()
        {
            Image img = null;
            Bitmap bitmap = null;

            if (BIZ_TYPE.Equals("PRMS"))
                //bitmap = Properties.Resources.QS_MAIN3;

            if (bitmap != null)
            {
                img = (Image)bitmap;
            }

            return img;
        }


        public static Image GetLoginImage()
        {
            Image img = null;
            Bitmap bitmap = null;

            if (BIZ_TYPE.Equals("PRMS"))
                //bitmap = Properties.Resources.DAESUNG_LOGIN;

            if (bitmap != null)
            {
                img = (Image)bitmap;
            }

            return img;
        }

        public static string GetSubjectName()
        {
            string sSbj = string.Empty;
            if (BIZ_TYPE.Equals("SAFEGD"))
                sSbj = "스마트가드레일";
            return sSbj;
        }

        public static string GetSubjectEngName()
        {
            string sSbj = string.Empty;
            if (BIZ_TYPE.Equals("SAFEGD"))
                sSbj = "SMART_GARDRAIL";
            return sSbj;
        }

        public static string GetDBName()
        {
            string DB = string.Empty;
            if (BIZ_TYPE.Equals("SAFEGD"))
                DB = "SAFEGD";
            return DB;
        }

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

        public static void ExportExcelFile(string sFileName, GridControl grid)
        {
            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = sFileName + DateTime.Now.ToLongDateString().Replace(" ", "");
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

        public static void ExportExcelFile(string sFileName, GridControl grid, string cmd, string sName, string sText)
        {
            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = sFileName + DateTime.Now.ToLongDateString().Replace(" ", "");
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    grid.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
                    GetLog(cmd, EXS_ID, Client_IP, sName, sText);
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

        public static void ExportPdfFile(string sFileName, GridControl grid)
        {
            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = sFileName + DateTime.Now.ToLongDateString().Replace(" ", "");
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Pdf files (*.pdf)|.pdf;";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    grid.ExportToPdf(FileName);
                    Process.Start(FileName);
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

        public static Dictionary<string, string> MakeDictionary(string[] sParamName, string[] sValue)
        {
            Dictionary<string, string> param = new Dictionary<string, string>();

            for (int i = 0; i < sParamName.Length; i++)
            {
                param.Add(sParamName[i], sValue[i]);
            }

            return param;
        }

       

        public static bool GetAuthInfo(string sCmd, string sUsrId, string sPgmId)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("CMD", sCmd);
            dicParams.Add("USRID", sUsrId);
            dicParams.Add("PGMID", sPgmId);

            DataTable dt = GetInfo(dicParams, "GET_AUTH_RST");
            if (dt.Rows.Count > 0)
            {
                string sRst = dt.Rows[0]["RST"]?.ToString();
                string sMsg = dt.Rows[0]["MSG"]?.ToString();
                if (sRst.Equals("Y"))
                    return false;
                else
                {
                    XtraMessageBox.Show(sMsg);
                    return true;
                }
            }
            else
            {
                string sWord = string.Empty;
                if (sCmd.Equals("USE"))
                    sWord = "사용";
                else if (sCmd.Equals("ADD"))
                    sWord = "추가";
                else if (sCmd.Equals("UPD"))
                    sWord = "수정";
                else if (sCmd.Equals("DEL"))
                    sWord = "삭제";
                else if (sCmd.Equals("PRT"))
                    sWord = "출력";
                else if (sCmd.Equals("XLS"))
                    sWord = "엑셀";
                XtraMessageBox.Show(string.Format("해당 사용자는 {0} 권한이 없습니다.", sWord));
                return true;
            }
        }

        public static bool GetPassInfo(string sCmd, string sUsrId, string sPgmId)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("CMD", sCmd);
            dicParams.Add("USRID", sUsrId);
            dicParams.Add("PGMID", sPgmId);

            DataTable dt = GetInfo(dicParams, "GET_AUTH_RST");
            if (dt.Rows.Count > 0)
            {
                XtraMessageBox.Show("패스워드 변경 후 사용가능합니다.");
                return true;
            }
            else
            {
                
                return false;
            }
        }

        // DB에 로그 저장
        public static bool GetLog(string sCmd, string sUsrId, string sComip, string sPgmId, string sPgmNm)
        {
            string sUsrCd = string.Empty;
            string sUsrNm = string.Empty;
            string sMacAd = ConvertMac(NetworkInterface.GetAllNetworkInterfaces()[0].GetPhysicalAddress().ToString());
            string sComNm = SystemInformation.ComputerName;

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Clear();
            dicParams.Add("CMD", "USERINFO");
            dicParams.Add("USRCD", sUsrId);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_LOGHIST", dicParams);
            if (dt.Rows.Count > 0)
            {
                sUsrCd = dt.Rows[0]["USRCD"].ToString();
                sUsrNm = dt.Rows[0]["USRNM"].ToString();
            }

            try
            {
                SqlCommand sqlCmd = new SqlCommand("DP_LOGHIST", DBConn.dbCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("CMD", sCmd);
                sqlCmd.Parameters.AddWithValue("USRCD", sUsrCd);
                sqlCmd.Parameters.AddWithValue("USRNM", sUsrNm);
                sqlCmd.Parameters.AddWithValue("COMIP", sComip);
                sqlCmd.Parameters.AddWithValue("MACAD", sMacAd);
                sqlCmd.Parameters.AddWithValue("COMNM", sComNm);
                sqlCmd.Parameters.AddWithValue("EXENM", string.Format(@"{0}\{1}", Application.StartupPath, "PRMS.exe"));
                sqlCmd.Parameters.AddWithValue("PGMID", sPgmId);
                sqlCmd.Parameters.AddWithValue("PGMNM", sPgmNm);

                SqlDataAdapter adp = new SqlDataAdapter(sqlCmd);
                DataSet ds = new DataSet();
                adp.Fill(ds);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

            return true;
        }

        // MAC주소 사이 하이픈('-') 추가
        public static string ConvertMac(string Mac)
        {
            string temp = Mac.Substring(0, 2) + "-" + Mac.Substring(2, 2) + "-" + Mac.Substring(4, 2) + "-"
                + Mac.Substring(6, 2) + "-" + Mac.Substring(8, 2) + "-" + Mac.Substring(10, 2);
            return temp;
        }

        /*
        public static bool GetLog(string sCmd, string sUsrId, string sPgmId, string sComip)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("CMD", sCmd);
            dicParams.Add("USRCD", sUsrId);
            dicParams.Add("PGMID", sPgmId);
            dicParams.Add("COMIP", sComip);
            dicParams.Add("EXENM", string.Format(@"{0}\{1}", Application.StartupPath, "DAESUNG_TEC.exe"));

            DataTable dt = DBConn.GetDataTable_1(DBConn.dbCon, "DP_LOGHIST", dicParams);
            //if (dt.Rows.Count > 0)
            //{
            //    string sRst = dt.Rows[0]["RST"]?.ToString();
            //    string sMsg = dt.Rows[0]["MSG"]?.ToString();
            //    if (sRst.Equals("Y"))
            //        return false;
            //    else
            //    {
            //        XtraMessageBox.Show(sMsg);
            //        return true;
            //    }
            //}
            //else
            //{
            return true;
            //}
        }
        */
      

        public static DataTable GetAuthInfo(string sUsrCd, string sProcedureName)
        {
            string[] sField = new string[2];
            string[] sValue = new string[sField.Length];

            sField[0] = "CMD";
            sValue[0] = "AUTH";

            sField[1] = "USRCD";
            sValue[1] = sUsrCd;

            return DBConn.GetDataTableByProcedure(DBConn.dbCon, sProcedureName, sField, sValue);
        }

        public static bool CheckEditableAuth(string sUsrCd, string sPgmId, string sProcedureName, string sEdit_Kind)
        {
            try
            {
                string sColumnsName = string.Empty;
                if (sEdit_Kind.Equals("A")) //추가
                {
                    sColumnsName = "ADD_Y";
                }
                else if (sEdit_Kind.Equals("U"))    //수정
                {
                    sColumnsName = "UPD_Y";
                }
                else if (sEdit_Kind.Equals("D"))    //삭제
                {
                    sColumnsName = "DEL_Y";
                }

                if (string.IsNullOrEmpty(sColumnsName))
                {
                    XtraMessageBox.Show("시스템 ERROR : ComnEtcFunc.CheckEditableAuth 참조, sColumnsName 변수 관련 에러");
                    return true;
                }

                string[] sField = new string[3];
                string[] sValue = new string[sField.Length];

                sField[0] = "CMD";
                sValue[0] = "AUTH_CHK";

                sField[1] = "USRCD";
                sValue[1] = sUsrCd;

                sField[2] = "PGMID";
                sValue[2] = sPgmId;

                string sResult = string.Empty;
                DataTable dt = DBConn.GetDataTableByProcedure(DBConn.dbCon, sProcedureName, sField, sValue);
                if (dt.Rows.Count > 0)
                {
                    sResult = dt.Rows[0][sColumnsName]?.ToString();
                }
                else
                {
                    XtraMessageBox.Show("해당 권한정보가 존재하지 않습니다. \r\n 시스템 ERROR : ComnEtc.CheckEditableAuth 참조, DataTable.Rows.Count == 0");
                    return true;
                }

                if (sResult.Equals("Y"))
                    return false;
                else
                    return true;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return true;
            }
        }

        public static void SetLookUpEdit(LookUpEdit lkup, DataTable dt, string sValue, string sDisplay, string sSetIdx)
        {
            lkup.Properties.DataSource = dt;
            lkup.Properties.ValueMember = sValue;
            lkup.Properties.DisplayMember = sDisplay;

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }

        public static void SetGridLookUpEdit(RepositoryItemGridLookUpEdit repositoryLookUp, DataTable dataTable, GridControl gridControl, GridColumn gridColumn, string sValue, string sDisplay, string sNullText)
        {
            GridView view = new GridView();

            repositoryLookUp.DataSource = dataTable;
            repositoryLookUp.ValueMember = sValue;
            repositoryLookUp.DisplayMember = sDisplay;

            gridControl.RepositoryItems.Add(repositoryLookUp);
            gridColumn.ColumnEdit = repositoryLookUp;
            repositoryLookUp.NullText = sNullText;

            repositoryLookUp.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryLookUp.PopupFilterMode = PopupFilterMode.Contains;
            repositoryLookUp.PopupView.OptionsFilter.AllowFilterIncrementalSearch = true;
            repositoryLookUp.ImmediatePopup = true;
            repositoryLookUp.View.OptionsView.ShowColumnHeaders = false;

        }

        // 모든 입력폼 리셋 (현재는 TextEdit, LookUpEdit, PictureEdit)
        public static void ClearAllForm(Control control)
        {
            foreach (Control ctrl in control.Controls)
            {
                TextEdit te = ctrl as TextEdit;
                if (te != null)
                {
                    te.ResetText();
                }

                if (ctrl is LookUpEdit)
                {
                    LookUpEdit lue = (LookUpEdit)ctrl;
                    if (lue != null)
                    {
                        lue.EditValue = string.Empty;
                    }
                }

                if (ctrl is PictureEdit)
                {
                    PictureEdit pe = (PictureEdit)ctrl;
                    if (pe != null)
                    {
                        pe.Image = null;
                        pe.Tag = null;
                    }
                }

                ClearAllForm(ctrl);
            }
        }

        // 룩업 일부(CD) 컬럼 배제 (모든 룩업 적용)
        public static void LkupCdClear(Control control)
        {
            foreach (Control ctrl in control.Controls)
            {
                LookUpEdit lue = ctrl as LookUpEdit;

                if (lue != null)
                {
                    lue.Properties.ForceInitialize();
                    lue.Properties.PopulateColumns();
                    lue.Properties.Columns[0].Visible = false;
                }
                LkupCdClear(ctrl);
            }
        }

        // 룩업 일부(CD) 컬럼 배제 (특정 룩업 적용)
        public static void LkupCdClear_One(LookUpEdit lue)
        {
            if (lue != null)
            {
                lue.Properties.ForceInitialize();
                lue.Properties.PopulateColumns();
                lue.Properties.Columns[0].Visible = false;
            }
        }

        public static void SelectRowByParam2(GridControl gv, string[] arrColName, string[] arrColVal)
        {
            if (arrColName[0] == null || arrColName[1] == null)
                return;
            if (arrColVal[0] == null || arrColVal[1] == null)
                return;

            ColumnView view = gv.MainView as ColumnView;
            GridColumn col1 = view.Columns[arrColName[0]];
            GridColumn col2 = view.Columns[arrColName[1]];
            if (col1 == null || col2 == null) return;

            view.OptionsSelection.MultiSelect = true;
            view.ClearSelection();
            int rowHandle = -1;

            while (rowHandle != GridControl.InvalidRowHandle)
            {
                rowHandle = view.LocateByDisplayText(rowHandle + 1, col1, arrColVal[0]);
                break;
            }

            for (int i = rowHandle; i < view.RowCount; i++)
            {
                string sVal = view.GetRowCellValue(i, arrColName[1])?.ToString();
                if (sVal.Equals(arrColVal[1]))
                {
                    view.SelectRow(i);
                    view.FocusedRowHandle = i;
                    break;
                }
            }
        }

        public static DataTable GetInfo(Dictionary<string, string> dicParams, string sProcedureId)
        {
            return DBConn.GetDataTable(DBConn.dbCon, sProcedureId, dicParams);
        }
        public static DataTable GetInfo1(Dictionary<string, string> dicParams, string sProcedureId)
        {
            return DBConn.GetDataTable(sProcedureId, dicParams);
        }
        public static class UserInfo
        {
            public static string USRCD;
            public static string USRID;
            public static string USRNM;
            public static string JCODE;
            public static string DEPTCD;
            public static string DEPTCDNM;
            public static string JKWICD;
            public static string JKWICDNM;
            public static string JKMUCD;
            public static string JKMUCDNM;
            public static string WKLOCA;
            public static string WKLOCANM;
            public static string TELNO;
            public static string CELNO;
            public static string EMAIL;
            public static string ADDR1;
            public static string ADDR2;
            public static string EMPID;
        }

        public static Dictionary<int, GridColumn> EssentialValueCheck(GridView view, GridColumn[] gridCol, GridColumn[] ApplyGridCol, string sGb, int iStartIdx, int LastIndex)
        {
            Dictionary<int, GridColumn> result = new Dictionary<int, GridColumn>();
            bool bYn = true;

            if (iStartIdx > view.RowCount - 1)
                return result;

            for (int i = iStartIdx; i < LastIndex; i++)
            {
                foreach (GridColumn col in gridCol)
                {
                    if (sGb.Equals("1"))
                    {
                        int findIndex = Array.FindIndex(gridCol, idx => idx == col);

                        string sVal = view.GetRowCellValue(i, col)?.ToString();
                        int iVal = 0;
                        int.TryParse(sVal, out iVal);
                        if (string.IsNullOrEmpty(sVal))
                        {
                            string sMSG = string.Format("{0}번째 행의 {1}을 입력하세요.", (i + 1), ApplyGridCol[findIndex].Caption);
                            XtraMessageBox.Show(sMSG);
                            result.Add(i, ApplyGridCol[findIndex]);
                            bYn = false;
                            view.Focus();
                            break;
                        }
                    }
                    else
                    {
                        int findIndex = Array.FindIndex(gridCol, idx => idx == col);

                        string sVal = view.GetRowCellValue(i, col)?.ToString();
                        double iVal = 0;
                        double.TryParse(sVal, out iVal);
                        if (iVal == 0)
                        {
                            string sMSG = string.Format("{0}번째 행의 {1}를 0 이상으로 입력하세요.", (i + 1), ApplyGridCol[findIndex].Caption);
                            XtraMessageBox.Show(sMSG);
                            result.Add(i, ApplyGridCol[findIndex]);
                            bYn = false;
                            view.Focus();
                            break;
                        }
                    }
                }

                if (!bYn)
                    break;
            }

            return result;
        }
        // 2021-08-26 추가
        public static bool CheckPassword(string pw)
        {
            /// 영 대/소문자 1개 이상, 숫자 1개 이상,  특수문자 1개 이상, 비밀번호 길이 9자 이상 체크
            // ^ : 라인의 처음
            // ?= : 전방 탐색
            // .* :  새 라인을 제외한 하나 이상 문자
            // [a-z] : a~z까지 영소문자,   [0-9] : 0~9까지 숫자,   [\W] : 특수문자열(!@#$...)
            // .{8,} : 8자리 이상
            // $ : 라인의 마지막
            Regex rxPassword = new Regex(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*[0-9])(?=.*[\W]).{9,}$");
            //Regex rxPassword = new Regex(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*[0-9]).{8,}$");
            // 입력한 Pw가 형식에 맞으면 true, 아니면 False 반환
            return rxPassword.IsMatch(pw);

            /// 특수문자 자체 사용방법
            // \^ : ^
            // \. : .
            // \[ : [
            // \$ : $
            // \( : (
            // \) : )
            // \| : |
            // \* : *
            // \+ : +
            // \? : ?
            // \{ : {
            // \\ : \
            // \n : 줄넘김 문자
            // \r : 리턴 문자
            // \w : 알파벳과 _ (언더바)
            // \W : 알파벳과 _ 가 아닌 것
            // \s : 빈 공간(space)
            // \S : 빈 공간이 아닌 것
            // \d : 숫자
            // \D : 숫자가 아닌 것
            // \b : 단어와 단어 사이의 경계
            // \B : 단어 사이의 경계가 아닌 것
            // \t : Tab 문자
            // \xnn : 16진수 nn에 해당하는 문자
        }
        public static void gp_SetBtnEnabled(Control sender)
        {
            string Chk = ComnEtcFunc.V_CHK;
            if (Chk.Equals("Y"))
            {
                foreach (Control ctl in sender.Controls)
                {
                    if (ctl is SimpleButton)
                    {
                        SimpleButton bt = (SimpleButton)ctl;
                        string str = bt.Tag?.ToString();
                        if (string.IsNullOrEmpty(str))
                        {
                            bt.Enabled﻿ = false;
                        }
                    }

                }
            }
        }

        /// <summary>
        /// 바이트 배열 -> 이미지
        /// </summary>
        /// <param name="byteArrayIn"></param>
        /// <returns></returns>
        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            Image returnImage = null;
            try
            {
                System.IO.MemoryStream ms = new MemoryStream(byteArrayIn, 0, byteArrayIn.Length);
                ms.Write(byteArrayIn, 0, byteArrayIn.Length);
                returnImage = Image.FromStream(ms, true);//Exception occurs here
            }
            catch { }
            return returnImage;
        }

        /// <summary>
        /// 이미지 -> 바이트 배열
        /// </summary>
        /// <param name="imageIn"></param>
        /// <returns></returns>
        //public static byte[] ImageToByteArray(System.Drawing.Image imageIn)
        //{
        //    using (var ms = new MemoryStream())
        //    {
        //        imageIn.Save(ms, imageIn.RawFormat);
        //        return ms.ToArray();
        //    }
        //}


        /// <summary> 
        /// Saves an image as a jpeg image, with the given quality 
        /// </summary> 
        /// <param name="path"> Path to which the image would be saved. </param> 
        /// <param name="quality"> An integer from 0 to 100, with 100 being the highest quality. </param> 
        public static byte[] SaveJpeg(string path, Image img, int quality)
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
        private static ImageCodecInfo GetEncoderInfo(string mimeType)
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
        public static byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
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

                while (true) {
                    int byteCount = sourceStream.Read(bufferByteArray, 0, bufferByteArray.Length);

                    if (byteCount == 0) {
                        break;
                    } targetFileStream.Write(bufferByteArray, 0, byteCount);
                }

                targetFileStream.Close();
                sourceStream.Close();

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
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
            req.UsePassive = false;
            // RequestStream에 데이타를 쓴다
            req.ContentLength = data.Length;
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();

            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }

        public static void SetFormText(Form form)
        {
            //string sFormText = MN001F00.SAVE_LAYOUT_LOADING_NAME;
            //form.Text += sFormText;
        }

        public static void SaveLayout(string sFormName, GridView[] arrGrdView, LayoutControl layout)
        {
            //GridView 저장
            for (int i = 0; i < arrGrdView.Length; i++)
            {
                string sFileName = string.Format("{0}{1}", sFormName, (i + 1).ToString());
                SaveGridViewLayout(EXS_ID, sFileName, arrGrdView[i]);
            }

            //Layout 저장
            SaveGridViewLayout(EXS_ID, sFormName, layout);
        }

        public static void SetLayout(string sFormName, GridView[] arrGrdView, LayoutControl layout)
        {
            for (int i = 0; i < arrGrdView.Length; i++)
            {
                string sFileName = string.Format("{0}{1}", sFormName, (i + 1).ToString());
                SetGridViewLayout(EXS_ID, "AccAdm", sFileName, arrGrdView[i]);
            }


        }
        public string Get_MyIP()
        {
            string strHostName = "";
            strHostName = System.Net.Dns.GetHostName();
            IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);
            IPAddress[] addr = ipEntry.AddressList; return addr[addr.Length - 1].ToString();

        }
        public static string Client_IP
        {
            get
            {
                IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
                string ClientIP = string.Empty;
                for (int i = 0; i < host.AddressList.Length; i++)
                {
                    if (host.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
                    {
                        ClientIP = host.AddressList[i].ToString();
                    }
                }
                return ClientIP;
            }
        }

        public static JObject GetAPIJSON()
        {
            JObject data = new JObject();

            data.Add("crtfcKey", _CRTKEY);
            data.Add("logDt", _LOGDT);
            data.Add("useSe", _USESE);
            data.Add("sysUser", _SYSUSER);
            data.Add("conectIp", _CONECTIP);
            data.Add("dataUsgqty", _DATAUSGQTY);

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(_URI + data.ToString());
            httpWebRequest.Method = "GET";

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();

                JObject res = JObject.Parse(result);
                MessageBox.Show(res["result"]["recptnRslt"].ToString());

                return res;
            }
        }

        #region [GridView Layout Setting]

        public static void SetGridViewLayout(string sId, string sProject, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            if (!File.Exists(sFile))
            {
                //sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sClass + ".xml";
                sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
                if (!File.Exists(sFile)) return;
            }
            view.RestoreLayoutFromXml(sFile);
        }

        public static void SetGridViewLayout(string sId, string sProject, string sClass, LayoutControl layout)
        {
            string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + "_Layout.xaml";
            if (!File.Exists(sFile))
            {
                //sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sClass + ".xml";
                sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + "_Layout.xaml";
                if (!File.Exists(sFile)) return;
            }
            layout.RestoreLayoutFromXml(sFile);
        }

        #endregion [GridView Layout Setting]

        #region [GridView Layout Save]

        public static void SaveGridViewLayout(string sId, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            //string path = @"C:\STLNT\" + sProject + @"\xaml\" + sId;
            string path = Application.StartupPath + @"\xaml\" + sId;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //string sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            view.SaveLayoutToXml(sFile);
        }

        public static void SaveGridViewLayout(string sId, string sClass, LayoutControl layout)
        {
            //string path = @"C:\STLNT\" + sProject + @"\xaml\" + sId;
            string path = Application.StartupPath + @"\xaml\" + sId;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //string sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + "_Layout.xaml";
            layout.SaveLayoutToXml(sFile);
        }

        #endregion [GridView Layout Save] 

        #region [Skin 적용]
        /*
         * #00001 
         */
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
                    //DevExpress.UserSkins.BonusSkins.Register();
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

                dicParams.Add("CMD", "SKIN3");
                dicParams.Add("USER", _USRCD);

                DataTable dt = DBConn.GetDataTable("DP_MN003F00", dicParams);
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
        #endregion


        #region[SetLookUpData]

        // LoopUpEdit 세팅 메소드
        // 사용  예 : ComnTestFunc.SetBoundLookUp(LkupHOUSE, "HOUSF", "HOUSE", "HOUNM"); 
        //            ComnTestFunc.SetBoundLookUp(LkupWLINE, "REFFPF", "WLINE", "");
        // 파라미터 : 바인딩할 lookupedit명, 값이 있는 테이블 명, 코드값 컬럼명, 보여줄 컬럼명
        //            공통코드 테이블(REFFPF)의 경우 마지막 파라미터는 고려되지 않음. "" 추천.
        public static void SetBoundLookUp(LookUpEdit lkup, string tableName, string columnName1, string columnName2)
        {
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" WITH ITEM_INFO AS ( ");

            strSql.AppendLine(" SELECT '' AS CD               ");
            strSql.AppendLine("      , '' AS NM               ");
            strSql.AppendLine("      , -1 AS SEQ              ");

            strSql.AppendLine(" UNION ALL                     ");

            if (tableName.Equals("REFFPF"))
            {
                strSql.AppendLine(" SELECT A.REFNO AS CD                                 ");
                strSql.AppendLine("      , A.REFNM AS NM                                 ");
                strSql.AppendLine("      , A.PRTSQ AS SEQ                                ");
                strSql.AppendLine("   FROM REFFPF A                                      ");
                strSql.AppendLine("  WHERE A.RCDTP = '" + columnName1 + "'               ");
                strSql.AppendLine("    AND A.REFNO <> '?'                                ");
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

        }

        public static void SetBoundLookUpWithTotal(LookUpEdit lkup, string tableName, string columnName1, string columnName2)
        {
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" WITH ITEM_INFO AS ( ");

            strSql.AppendLine(" SELECT '' AS CD               ");
            strSql.AppendLine("      , '전체' AS NM           ");
            strSql.AppendLine("      , -1 AS SEQ              ");

            strSql.AppendLine(" UNION ALL                     ");

            if (tableName.Equals("REFFPF"))
            {
                strSql.AppendLine(" SELECT A.REFNO AS CD                                 ");
                strSql.AppendLine("      , A.REFNM AS NM                                 ");
                strSql.AppendLine("      , A.PRTSQ AS SEQ                                ");
                strSql.AppendLine("   FROM REFFPF A                                      ");
                strSql.AppendLine("  WHERE A.RCDTP = '" + columnName1 + "'               ");
                strSql.AppendLine("    AND A.REFNO <> '?'                                ");
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

        }

        // GridLoopUpEdit 세팅 메소드
        // 사용  예 : ComnTestFunc.SetBoundGridLookUp(RepoGLkupWLINE, "REFFPF", "WLINE", "");
        //            ComnTestFunc.SetBoundGridLookUp(RepoGLkupMACOD, "EQUIPF", "MACOD", "MANAM");
        // 파라미터 : 바인딩할 lookupedit명, 값이 있는 테이블 명, 코드값 컬럼명, 보여줄 컬럼명
        //            공통코드 테이블(REFFPF)의 경우 마지막 파라미터는 고려되지 않음. "" 추천.
        public static void SetBoundGridLookUp(RepositoryItemGridLookUpEdit repositoryLookUp, string tableName, string columnName1, string columnName2)
        {
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" WITH ITEM_INFO AS ( ");

            strSql.AppendLine(" SELECT '' AS CD               ");
            strSql.AppendLine("      , '' AS NM               ");
            strSql.AppendLine("      , 0 AS SEQ               ");

            strSql.AppendLine(" UNION ALL                     ");

            if (tableName.Equals("REFFPF"))
            {
                strSql.AppendLine(" SELECT A.REFNO AS CD                                 ");
                strSql.AppendLine("      , A.REFNM AS NM                                 ");
                strSql.AppendLine("      , A.PRTSQ AS SEQ                                ");
                strSql.AppendLine("   FROM REFFPF A                                      ");
                strSql.AppendLine("  WHERE A.RCDTP = '" + columnName1 + "'               ");
                strSql.AppendLine("    AND A.REFNO <> '?'                                ");
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


            repositoryLookUp.DataSource = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            repositoryLookUp.ValueMember = "CD";
            repositoryLookUp.DisplayMember = "NM";

            repositoryLookUp.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryLookUp.PopupFilterMode = PopupFilterMode.Contains;
            repositoryLookUp.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            repositoryLookUp.PopupView.OptionsFilter.AllowFilterIncrementalSearch = true;
            repositoryLookUp.ImmediatePopup = true;
            repositoryLookUp.View.OptionsView.ShowColumnHeaders = false;

        }

        #endregion

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
                            JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString() + "\":" + "\"" + ds.Tables[0].Rows[i][j].ToString() + "\",");
                        }
                        else if (j == ds.Tables[0].Columns.Count - 1)
                        {
                            JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString() + "\":" + "\"" + ds.Tables[0].Rows[i][j].ToString() + "\"");
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
        #endregion

        public static void gp_SetColorFocused(Control sender)
        {
            // Control에 포커스 가면 색 변경 (SHS: 2020.12.02)
            foreach (Control ctl in sender.Controls)
            {
                if (ctl is TextEdit) { TextEdit tb = (TextEdit)ctl; tb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow; }
                else
                if (ctl is MemoEdit) { MemoEdit cb = (MemoEdit)ctl; cb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow; }
                else
                if (ctl is LookUpEdit) { LookUpEdit cb = (LookUpEdit)ctl; cb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow; }
            }
        }

        public static void gp_ResetCompo(Control sender)
        {
            // 컨트롤 초기화 (SHS: 2020.11.24)
            foreach (Control ctl in sender.Controls)
            {
                if (ctl is TextBox) { TextBox tb = (TextBox)ctl; if (string.IsNullOrEmpty(tb.Tag?.ToString())) { tb.Text = string.Empty; } }
                else
                if (ctl is TextEdit) { TextEdit tb = (TextEdit)ctl; if (string.IsNullOrEmpty(tb.Tag?.ToString())) { tb.Text = string.Empty; } }
                else
                if (ctl is MemoEdit) { MemoEdit me = (MemoEdit)ctl; if (string.IsNullOrEmpty(me.Tag?.ToString())) { me.Text = string.Empty; } }
                else
                if (ctl is DateEdit) { DateEdit me = (DateEdit)ctl; if (string.IsNullOrEmpty(me.Tag?.ToString())) { me.Text = string.Empty; } }
                else
                if (ctl is ButtonEdit) { ButtonEdit me = (ButtonEdit)ctl; if (string.IsNullOrEmpty(me.Tag?.ToString())) { me.Text = string.Empty; } }
                else
                if (ctl is LookUpEdit) { LookUpEdit cb = (LookUpEdit)ctl; if (string.IsNullOrEmpty(cb.Tag?.ToString())) { cb.SelectedText = string.Empty; } }
            }
        }

        // 폼 별 컨트롤 규칙 적용
        public static void InitControllerRule(Control control)
        {
            foreach (Control ctrl in control.Controls)
            {
                if (ctrl is DateEdit)
                {
                    DateEdit dt = (DateEdit)ctrl;
                    dt.Properties.ShowClear = false;
                    dt.EnterMoveNextControl = true;
                    dt.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret;
                    dt.Properties.Mask.UseMaskAsDisplayFormat = true;
                    dt.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow;
                }
                else if (ctrl is ComboBoxEdit)
                {
                    ComboBoxEdit cb = (ComboBoxEdit)ctrl;
                    cb.EnterMoveNextControl = true;
                    cb.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
                    cb.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow;
                }
                else if (ctrl is TextEdit)
                {
                    TextEdit tx = (TextEdit)ctrl;
                    tx.EnterMoveNextControl = true;
                    tx.ImeMode = ImeMode.Hangul;
                    tx.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow;
                }
                else if (ctrl is ButtonEdit)
                {
                    ButtonEdit be = (ButtonEdit)ctrl;
                    be.EnterMoveNextControl = true;
                    be.ImeMode = ImeMode.Hangul;
                    be.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow;
                }
                else if (ctrl is LookUpEdit)
                {
                    LookUpEdit lk = (LookUpEdit)ctrl;
                    lk.EnterMoveNextControl = true;
                    lk.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow;
                }
                else if (ctrl is RadioGroup)
                {
                    RadioGroup rg = (RadioGroup)ctrl;
                    rg.EnterMoveNextControl = true;
                }
                else if (ctrl is MemoEdit)
                {
                    MemoEdit me = (MemoEdit)ctrl;
                    me.ImeMode = ImeMode.Hangul;
                    me.Properties.AppearanceFocused.BackColor = System.Drawing.Color.LightYellow;
                    //me.EnterMoveNextControl = true;
                }
            }
        }
    }
}
