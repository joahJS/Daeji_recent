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
using ComLib;
using System.Net;
using System.IO;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 */
namespace AccAdm
{
    public partial class PD02001F01 : DevExpress.XtraEditors.XtraForm
    {
        public PD02001F01()
        {
            InitializeComponent();
        }

        public GridView[] arrGrdView;
        private void PD02001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout2.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl2.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout3.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl3.RestoreLayoutFromXml(sFile);
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("등록일자를 바르게 입력하세요.");
                DateEditFrom.SelectAll();
                DateEditFrom.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("등록일자를 바르게 입력하세요.");
                DateEditTo.SelectAll();
                DateEditTo.Focus();
                return;
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", sYmdFrom);
            dicParams.Add("DATE_T", sYmdTo);
            dicParams.Add("FIND_IDX", CboFindSbj.SelectedIndex.ToString());
            dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());
            dicParams.Add("ITNL_YN", RdgbItnlYn.EditValue?.ToString());

            GridControl grdList = new GridControl();
            GridView grdView = new GridView();
            DataTable dt = new DataTable();
            if (TabControl.SelectedTabPage == TabPageGumsu)
            {
                grdList = GridRetr;
                grdView = GridViewRetr;
                dt = GetInfo(dicParams);
            }
            else
            {
                grdList = GridRetr2;
                grdView = GridViewRetr2;
                dt = GetSummary(dicParams);
            }


            grdList.DataSource = dt;
            if (grdView.RowCount > 0)
            {
                grdView.Focus();
            }
            else if(grdView.RowCount == 0)
            {
                if (string.IsNullOrEmpty(TxtFindWord.EditValue?.ToString()))
                {
                    TxtFindWord.SelectAll();
                    TxtFindWord.Focus();
                    return;
                }
                else
                {
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }
            }

            
        }


        private DataTable GetInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A1.MAKENO_LM ");
            strSql.AppendLine("      , DATE_FORMAT(A.MDATE, '%Y-%m-%d') AS MDATE ");
            strSql.AppendLine("      , B.EMP_NM ");
            strSql.AppendLine("      , A1.M4_CARNO ");
            strSql.AppendLine("      , A1.M4_CVNAM ");
            strSql.AppendLine("      , A1.M4_GRADE ");
            strSql.AppendLine("      , A1.M4_WGT ");
            strSql.AppendLine("      , A1.M4_MINUS ");
            strSql.AppendLine("      , A1.M4_EVIDENCE ");
            strSql.AppendLine("      , A1.M4_WGT_ADMT ");
            strSql.AppendLine("      , IFNULL(A1.M4_ITNL_YN, 'N') AS M4_ITNL_YN ");
            strSql.AppendLine("      , A1.M4_ISPT_OPN ");
            strSql.AppendLine("      , IFNULL(A1.IMG_CNT, 0) AS IMG_CNT ");
            strSql.AppendLine("      , A2.SIGN1 ");
            strSql.AppendLine("      , A2.SIGN2 ");
            strSql.AppendLine("      , A2.SIGN3 ");
            strSql.AppendLine("      , B1.USRNM AS ENT_NM ");
            strSql.AppendLine("      , A1.ENT_DT ");
            strSql.AppendLine("      , B2.USRNM AS MFY_NM ");
            strSql.AppendLine("      , A1.MFY_DT ");
            strSql.AppendLine("   FROM MAKE_M A ");
            strSql.AppendLine("   LEFT JOIN MAKE_4 A1 ");
            strSql.AppendLine("     ON A.MAKENO = A1.MAKENO  ");
            strSql.AppendLine("   LEFT JOIN MAKE_S A2 ");
            strSql.AppendLine("     ON A.MDATE = A2.MDATE ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B  ");
            strSql.AppendLine("     ON A.MUSER_ID = B.EMP_ID ");
            strSql.AppendLine("   LEFT JOIN ZUSRLST B1 ");
            strSql.AppendLine("     ON A1.ENT_ID = B1.USRCD ");
            strSql.AppendLine("   LEFT JOIN ZUSRLST B2 ");
            strSql.AppendLine("     ON A1.MFY_ID = B2.USRCD ");
            strSql.AppendLine("  WHERE A.MDATE BETWEEN @DATE_F AND @DATE_T  ");
            strSql.AppendLine("   AND A1.MAKENO IS NOT NULL ");
            strSql.AppendLine("   AND ((('" + dicParams["FIND_WORD"] + "' = '' OR '" + dicParams["FIND_WORD"] + "' IS NULL) AND 1 = 1) ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@FIND_IDX = '0' AND A1.M4_CVNAM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@FIND_IDX = '1' AND A1.M4_GRADE LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@FIND_IDX = '2' AND A1.M4_EVIDENCE LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@FIND_IDX	= '3' AND A1.M4_WGT_ADMT LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@FIND_IDX	= '4' AND A1.M4_ISPT_OPN LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@FIND_IDX = '5' AND B.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@FIND_IDX	= '6' AND A1.M4_CARNO LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("   AND ((@ITNL_YN = 'ALL' AND 1 = 1) ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@ITNL_YN = 'Y' AND A1.M4_ITNL_YN = 'Y') ");
            strSql.AppendLine("        OR  ");
            strSql.AppendLine("        (@ITNL_YN = 'N' AND A1.M4_ITNL_YN <> 'Y')) ");
            strSql.AppendLine(" ORDER BY A.MDATE, A.MAKENO, A1.MAKENO_LM ");
            
            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetSummary(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT '스크랩' AS GB #스크랩 ");
            strSql.AppendLine("      , A1.`M4_CVCOD`  ");
            strSql.AppendLine("      , B.DEALER_NM ");
            strSql.AppendLine("      , SUM(A1.M4_WGT) AS WGT ");
            strSql.AppendLine("      , SUM(A1.M4_MINUS) AS MINUS ");
            strSql.AppendLine("      , C.IN_CNT ");
            strSql.AppendLine("      , IFNULL(C.CHAGAM, 0) AS CHAGAM ");
            strSql.AppendLine("      , IFNULL(C.CNT, 0) AS CNT ");
            strSql.AppendLine("      , CASE WHEN COUNT(CASE WHEN A1.M4_ITNL_YN = 'Y' THEN 1 END) > 0 THEN '유' ELSE '무' END AS ITNL_GB ");
            strSql.AppendLine("   FROM MAKE_M A  ");
            strSql.AppendLine("   LEFT JOIN MAKE_4 A1  ");
            strSql.AppendLine("     ON A.MAKENO = A1.MAKENO   ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B ");
            strSql.AppendLine("     ON A1.M4_CVCOD = B.DEALER_CD ");
            strSql.AppendLine("   LEFT JOIN ( SELECT SUM(A.ICHAGAM) AS CHAGAM ");
            strSql.AppendLine("                    , COUNT(*) AS IN_CNT ");
            strSql.AppendLine("                    , COUNT(CASE WHEN A.ICHAGAM > 0 THEN 1 END) AS CNT ");
            strSql.AppendLine("                    , A.MAIPCHERID ");
            strSql.AppendLine("                 FROM MESURING A ");
            strSql.AppendLine("                 LEFT JOIN JAJAE B ");
            strSql.AppendLine("                   ON A.J_SERIAL = B.J_SERIAL ");
            strSql.AppendLine("                WHERE KERATYPE = '입고' ");
            strSql.AppendLine("                  AND A.J_DATE BETWEEN DATE_FORMAT(@DATE_F, '%Y-%m-%d') AND DATE_FORMAT(@DATE_T, '%Y-%m-%d') ");
            strSql.AppendLine("                  AND B.DAEGUBUN IN ('고철A', '고철B') ");
            strSql.AppendLine("                GROUP BY A.MAIPCHERID ) C ");
            strSql.AppendLine("     ON A1.M4_CVCOD = C.MAIPCHERID ");
            strSql.AppendLine("   LEFT JOIN JAJAE D ");
            strSql.AppendLine("     ON A1.M4_GRADE_CD = D.J_SERIAL ");
            strSql.AppendLine("  WHERE A1.M4_CVCOD IS NOT NULL ");
            strSql.AppendLine("    AND A.MDATE BETWEEN DATE_FORMAT(@DATE_F, '%Y%m%d') AND DATE_FORMAT(@DATE_T, '%Y%m%d') ");
            strSql.AppendLine("    AND D.DAEGUBUN IN ('고철A', '고철B') ");
            strSql.AppendLine("  GROUP BY A1.M4_CVCOD ");
            strSql.AppendLine("  UNION ALL ");
            strSql.AppendLine("  SELECT '슈레더' AS GB #슈레더 ");
            strSql.AppendLine("      , A1.`M4_CVCOD`  ");
            strSql.AppendLine("      , B.DEALER_NM ");
            strSql.AppendLine("      , SUM(A1.M4_WGT) AS WGT ");
            strSql.AppendLine("      , SUM(A1.M4_MINUS) AS MINUS ");
            strSql.AppendLine("      , C.IN_CNT ");
            strSql.AppendLine("      , IFNULL(C.CHAGAM, 0) AS CHAGAM ");
            strSql.AppendLine("      , IFNULL(C.CNT, 0) AS CNT ");
            strSql.AppendLine("      , CASE WHEN COUNT(CASE WHEN A1.M4_ITNL_YN = 'Y' THEN 1 END) > 0 THEN '유' ELSE '무' END AS ITNL_GB ");
            strSql.AppendLine("   FROM MAKE_M A  ");
            strSql.AppendLine("   LEFT JOIN MAKE_4 A1  ");
            strSql.AppendLine("     ON A.MAKENO = A1.MAKENO   ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B ");
            strSql.AppendLine("     ON A1.M4_CVCOD = B.DEALER_CD ");
            strSql.AppendLine("   LEFT JOIN ( SELECT SUM(A.ICHAGAM) AS CHAGAM ");
            strSql.AppendLine("                    , COUNT(*) AS IN_CNT ");
            strSql.AppendLine("                    , COUNT(CASE WHEN A.ICHAGAM > 0 THEN 1 END) AS CNT ");
            strSql.AppendLine("                    , A.MAIPCHERID ");
            strSql.AppendLine("                 FROM MESURING A ");
            strSql.AppendLine("                 LEFT JOIN JAJAE B ");
            strSql.AppendLine("                   ON A.J_SERIAL = B.J_SERIAL ");
            strSql.AppendLine("                WHERE KERATYPE = '입고' ");
            strSql.AppendLine("                  AND A.J_DATE BETWEEN DATE_FORMAT(@DATE_F, '%Y-%m-%d') AND DATE_FORMAT(@DATE_T, '%Y-%m-%d') ");
            strSql.AppendLine("                  AND B.DAEGUBUN IN ('슈레더') ");
            strSql.AppendLine("                GROUP BY A.MAIPCHERID ) C ");
            strSql.AppendLine("     ON A1.M4_CVCOD = C.MAIPCHERID ");
            strSql.AppendLine("   LEFT JOIN JAJAE D ");
            strSql.AppendLine("     ON A1.M4_GRADE_CD = D.J_SERIAL ");
            strSql.AppendLine("  WHERE A1.M4_CVCOD IS NOT NULL ");
            strSql.AppendLine("    AND A.MDATE BETWEEN DATE_FORMAT(@DATE_F, '%Y%m%d') AND DATE_FORMAT(@DATE_T, '%Y%m%d') ");
            strSql.AppendLine("    AND D.DAEGUBUN IN ('슈레더') ");
            strSql.AppendLine("  GROUP BY A1.M4_CVCOD ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            ComnEtcFunc.ExportExcelFile(string.Format("{0}_", this.Text), GridRetr);
        }

        public string RST_DT;
        private void BtnReport_Click(object sender, EventArgs e)
        {
            PD02001F02 frm = new PD02001F02();
            frm.P_PD02001F01 = this;
            if(frm.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            DateTime dt = DateTime.Parse(RST_DT);
            int iWeekCnt = GetWeekCnt(dt);

            DateTime FirstDateOfWeek = GetFirstDateOfWeek(dt.Year, iWeekCnt);
            DateTime LastDateOfWeek = FirstDateOfWeek.AddDays(6);
            string sWeek = string.Format("{0}   ~   {1}   {2}월 주차{3}", FirstDateOfWeek.ToString("yyyy-MM-dd"), LastDateOfWeek.ToString("yyyy-MM-dd"), FirstDateOfWeek.Month, CalcWeekNumberFromDate(dt.Year, dt.Month, dt.Day).ToString());
            
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("DATE", sWeek);

            
        }

        // 내가 만든 것.
        private static int GetWeekCnt(DateTime dt)
        {
            int week = Enum.GetValues(typeof(DayOfWeek)).Length;
            int dayOffset = (int)dt.AddDays(-(dt.Day - 1)).DayOfWeek;
            int weekCnt = (dt.Day + dayOffset) / week;
            weekCnt += ((dt.Day + dayOffset) % week) > 0 ? 1 : 0;
            return weekCnt;
        }
        
        // 이건 해당 년도에 주차 수를 구한다.
        public int GetWeeksOfYear(DateTime date)
        {
            System.Globalization.CultureInfo cult_info = System.Globalization.CultureInfo.CreateSpecificCulture("ko");
            System.Globalization.Calendar cal = cult_info.Calendar;
            int weekNo = cal.GetWeekOfYear(date, cult_info.DateTimeFormat.CalendarWeekRule, cult_info.DateTimeFormat.FirstDayOfWeek);
            int week1day = cal.GetWeekOfYear(date.AddDays(-(date.Day + 1)), cult_info.DateTimeFormat.CalendarWeekRule, cult_info.DateTimeFormat.FirstDayOfWeek);
            return weekNo - week1day + 1;
        }

        /// <summary>
        /// 특정 주차 시작일 구하기
        /// </summary>
        /// <param name="year">연도</param>
        /// <param name="week">주차</param>
        /// <returns>특정 주차 시작일</returns>
        public DateTime GetFirstDateOfWeek(int year, int week)
        {
            DateTime firstDateOfYear = new DateTime(year, 1, 1);
            DateTime firstDateOfFirstWeek = firstDateOfYear.AddDays(7 - (int)(firstDateOfYear.DayOfWeek) + 1);
            return firstDateOfFirstWeek.AddDays(7 * (week - 1));
        }

        private static int CalcWeekNumberFromDate(int year, int month, int day)
        {
            int iWeekNumber = 0;
            DateTime dTime = new DateTime(year, month, day);
            string FirstOfMonth = DateTime.Now.ToString("yyyy-MM-01");
            DateTime dt = new DateTime(); DateTime.TryParse(FirstOfMonth, out dt);
            System.Globalization.CultureInfo myCl = new System.Globalization.CultureInfo("ko-KR");
            System.Globalization.Calendar myCal = myCl.Calendar;
            System.Globalization.CalendarWeekRule myCWR = myCl.DateTimeFormat.CalendarWeekRule;
            DayOfWeek myFirstDOW = myCl.DateTimeFormat.FirstDayOfWeek;
            
            //오늘 몇주차
            int WeekOfToday = myCal.GetWeekOfYear(dTime, myCWR, myFirstDOW);
            
            //오늘이 있는달 첫날짜 주차
            int WeekOfFirstday = myCal.GetWeekOfYear(dt, myCWR, myFirstDOW);
            int WeekOfMonth = WeekOfToday - WeekOfFirstday;
            iWeekNumber = WeekOfMonth + 1;

            return iWeekNumber;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void PD02001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RdgbItnlYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void CboFindSbj_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(e.FocusedRowHandle < 0)
            {
                if(GridImg.DataSource != null)
                {
                    DataTable dt = (DataTable)GridImg.DataSource;
                    dt.Rows.Clear();
                    GridImg.DataSource = dt;
                }
                return;
            }

            //불필요한 FTP 서버의 접속을 줄이기 위하여 이미지카운트가 0 인경우 아래 로직 수행되지 않도록 구현
            string sImgCnt = GridViewRetr.GetFocusedRowCellValue(GridCol1ImgCnt)?.ToString();
            int iImgCnt = string.IsNullOrEmpty(sImgCnt) ? 0 : Convert.ToInt32(sImgCnt);

            if (iImgCnt == 0)
            {
                //이미지 그리드 초기화
                if (GridImg.DataSource != null)
                {
                    DataTable dt = (DataTable)GridImg.DataSource;
                    dt.Rows.Clear();
                    GridImg.DataSource = dt;
                }
                return;
            }


            string sMDATE = GridViewRetr.GetFocusedRowCellValue(GridCol1MDate)?.ToString().Replace("-", "");
            string sMAKENO = GridViewRetr.GetFocusedRowCellValue(GridCol1MakeNo)?.ToString();
            string sMAKENO_LM = GridViewRetr.GetFocusedRowCellValue(GridCol1MakeNoLn)?.ToString();

            GetImagesFromFTP(sMDATE, sMAKENO, sMAKENO_LM);

        }

        private void GetImagesFromFTP(string MDATE, string MAKENO, string MAKENO_LN)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                //string sInitDir = string.Format(@"ftp://192.168.0.202/Gumsu_Images/{0}/{1}/{2}/{3}_{4}", MDATE.Substring(0, 4), MDATE.Substring(4, 2), MDATE.Substring(6, 2), MAKENO, MAKENO_LN);
                string sInitDir = string.Format(@"ftp://"+ComnEtcFunc.FTP_IP+"/Gumsu_Images/{0}/{1}/{2}/{3}_{4}", MDATE.Substring(0, 4), MDATE.Substring(4, 2), MDATE.Substring(6, 2), MAKENO, MAKENO_LN);

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
                        dicImages.Add(filesCopy[filesCopy.Length - 1], ComnEtcFunc.DownloadFTPFile(string.Format(@"{0}\{1}", sInitDir, filePath), user, pw));
                    }
                }

                DataTable dt = new DataTable();
                dt.TableName = "Table1";
                dt.Columns.Add("IMAGE", typeof(byte[]));
                dt.Columns.Add("FILE_NAME");

                foreach (KeyValuePair<string, Image> item in dicImages)
                {
                    DataRow row = dt.NewRow();
                    row["FILE_NAME"] = item.Key;
                    row["IMAGE"] = ComnEtcFunc.ImageToByteArray(item.Value);
                    dt.Rows.Add(row);
                }

                GridImg.DataSource = dt;

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                return;
            }
        }

        private void TabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if(e.Page == TabPageGumsu)
            {
                LayoutFindIdx.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutFindWord.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutItnl.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else
            {
                LayoutFindIdx.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutFindWord.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutItnl.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }

            BtnRetr.PerformClick();
        }

        private void PD02001F01_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                string path = ComnEtcFunc.GetLayoutPath();
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");
                layoutControl2.SaveLayoutToXml(path + @"\" + this.Name + "_Layout2.xaml");
                layoutControl3.SaveLayoutToXml(path + @"\" + this.Name + "_Layout3.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}