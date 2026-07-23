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
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using DevExpress.Spreadsheet;
using System.Data.SqlClient;
using System.Diagnostics;
using DevExpress.XtraSplashScreen;
using System.Threading;
using DevExpress.Utils.Menu;

/*
 * 수정일자: 2023-01-04
 * 수정자  : 정은영
 * ID      : #0001
 * 내용    : (현업요청)
 *           1. 상단 검색 하단검색으로 이동. 문서구분 및 상태표시 검색제거
 *           2. 상단 조회 내용: 내 결재건, 결재진행중인건 표시
 *           3. 하단 조회 내용: 완료,최종반려 부서별 조회. 경영관리부는 전체조회. 단, 잠금은 결재라인만
 *           4. 폼 시작시 바로 조회
 *           5. 대표이사 전체 다 보도록 추가
 * 
 * 수정일자 : 2023-01-10
 * 수정자   : 정은영
 * ID       : #0002
 * 내용     : (현업요청)
 *            1. 하단 기본조회 작성자 명으로 변경
 *            2. 최근작성 문서 순으로 정렬
 *            
 * 수정일자 : 2023-01-13
 * 수정자   : 정은영
 * ID       : #0003
 * 내용     : (현업요청)
 *            1. 하단 결재일 추가 해서 최종 승인일 표시
 *            2. 최종승인일 내림차순정렬로 변경
 *            
 *            
 * 수정일자 : 2023-02-23
 * 수정자   : 정은영
 * ID       : #0004
 * 내용     : (현업요청)
 *            1. 하단 조회 작성자명으로 검색 되는거 제거
 *       
 */
namespace AccAdm
{
    public partial class RptApplSystem : DevExpress.XtraEditors.XtraForm
    {
        public RptApplSystem()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_SI003F01";

        //FTP 서버
        private string sInitDir = ComnEtcFunc.FTP_ROOT + @"/ERP/AttFile/"; //첨부파일
        private string user = ComnEtcFunc.FTP_USER;
        private string pw = ComnEtcFunc.FTP_PW;

        //임시 문서파일저장 경로
        private string tempDoctPath = Application.StartupPath + @"/tempDoct/" + FmMainToolBar2.drUser["USRCD"];
        //임시 파일명
        private string _tempFileName = string.Empty;

        //파일 임시저장 폴더
        private string attFolder = Application.StartupPath + "/docttemp/" + FmMainToolBar2.drUser["USRCD"] + "/";

        //
        public string _SLINO = string.Empty;
        public int _FINDOBJ;
        public string _FINDWORD;

        private void RptApplSystem_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            SetLoadFormLayout();

            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);

            //#0002 //#0004 제거 요청으로 주석처리
            //TxtFindWord.EditValue = FmMainToolBar2.drUser["USRNM"];
        }

        private void RptApplSystem_Shown(object sender, EventArgs e)
        {
            //#0001
            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr, GridViewRetr3 };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
        }
        #endregion

        #region LOOKUP DATA
        private DataTable GetLookupData(string sGB)
        {
            StringBuilder strSql = new StringBuilder();

            if (sGB.Equals("0"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT DOCTP AS CD                                  ");
                strSql.AppendLine("      , DOCNM AS NM                                  ");
                strSql.AppendLine("   FROM DOCT_K                                       ");
            }

            if (sGB.Equals("1")) //문서번호
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT SLINO AS CD");
                strSql.AppendLine("   FROM DOCT_M     ");
            }

            if (sGB.Equals("2"))//제목
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT STITL AS CD");
                strSql.AppendLine("   FROM DOCT_M     ");
            }

            if (sGB.Equals("3"))//작성자
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT A2.USRNM AS CD      ");
                strSql.AppendLine("   FROM DOCT_M A1           ");
                strSql.AppendLine("   LEFT JOIN ZUSRLST A2     ");
                strSql.AppendLine("     ON A1.PLNCD = A2.USRCD ");
                strSql.AppendLine("  GROUP BY A2.USRNM         ");
            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        #endregion

        #region event
        public void SearchFun()
        {
            if (!string.IsNullOrEmpty(_FINDWORD))
            {
                CboFindObj.SelectedIndex = _FINDOBJ;
                TxtFindWord.EditValue = _FINDWORD;

                _SLINO = _FINDWORD;

                BtnRetr.PerformClick();

                TxtFindWord.EditValue = "";
                _SLINO = string.Empty;
            }
        }

        //단축키
        private void RptApplSystem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        #region 조회
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            GetGridRetr3();
            GetGridRetr1();
        }

        //결재대기문서 조회
        private void GetGridRetr1()
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                int focseIdx = GridViewRetr.FocusedRowHandle;

                string sUsrcd = FmMainToolBar2.drUser["USRCD"]?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                #region 2023-01-04 이전코드 
                //strSql.AppendLine(" WITH TEMP1 AS(                       ");
                //strSql.AppendLine("        SELECT SLINO                  ");
                //strSql.AppendLine("             , GUBUN                  ");
                //strSql.AppendLine("             , MIN(SEQNO) AS SEQNO    ");
                //strSql.AppendLine("          FROM DOCT_SIGN              ");
                //strSql.AppendLine("         WHERE APPTP = '3'            ");
                //strSql.AppendLine("           AND GUBUN = '0'            ");
                //strSql.AppendLine("         GROUP BY SLINO, GUBUN        ");
                //strSql.AppendLine(" ), TEMP2 AS(--다음결재자                         ");
                //strSql.AppendLine("     SELECT A1.SLINO                  ");
                //strSql.AppendLine("          , A1.GUBUN                  ");
                //strSql.AppendLine("          , A1.SEQNO                  ");
                //strSql.AppendLine("          , A2.PLNCD                  ");
                //strSql.AppendLine("       FROM TEMP1 A1                  ");
                //strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                //strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                //strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                //strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                //strSql.AppendLine(" ), TEMP3 AS(                         ");
                //strSql.AppendLine("     SELECT SLINO                     ");
                //strSql.AppendLine("          , GUBUN                     ");
                //strSql.AppendLine("          , MAX(SEQNO) AS SEQNO       ");
                //strSql.AppendLine("       FROM DOCT_SIGN                 ");
                //strSql.AppendLine("      WHERE APPTP = '1'               ");
                //strSql.AppendLine("        AND GUBUN = '0'               ");
                //strSql.AppendLine("      GROUP BY SLINO, GUBUN           ");
                //strSql.AppendLine(" ), TEMP4 AS(--이전결재자                         ");
                //strSql.AppendLine("     SELECT A1.SLINO                  ");
                //strSql.AppendLine("          , A1.GUBUN                  ");
                //strSql.AppendLine("          , A1.SEQNO                  ");
                //strSql.AppendLine("          , A2.PLNCD                  ");
                //strSql.AppendLine("       FROM TEMP3 A1                  ");
                //strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                //strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                //strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                //strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                //strSql.AppendLine(" ), TEMP6 AS(                         ");
                //strSql.AppendLine("     SELECT SLINO                     ");
                //strSql.AppendLine("          , GUBUN                     ");
                //strSql.AppendLine("          , MAX(SEQNO) AS SEQNO       ");
                //strSql.AppendLine("       FROM DOCT_SIGN                 ");
                //strSql.AppendLine("      WHERE GUBUN = '0'               ");
                //strSql.AppendLine("      GROUP BY SLINO, GUBUN           ");
                //strSql.AppendLine(" ), TEMP7 AS(--최종결재자                         ");
                //strSql.AppendLine("     SELECT A1.SLINO                  ");
                //strSql.AppendLine("          , A1.GUBUN                  ");
                //strSql.AppendLine("          , A1.SEQNO                  ");
                //strSql.AppendLine("          , A2.PLNCD                  ");
                //strSql.AppendLine("       FROM TEMP6 A1                  ");
                //strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                //strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                //strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                //strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                //strSql.AppendLine(" ), TEMP8 AS(                         ");
                //strSql.AppendLine("     SELECT A1.SLINO                  ");
                //strSql.AppendLine("          , A1.TDATE                  ");
                //strSql.AppendLine("          , A1.PLNCD                  ");
                //strSql.AppendLine("          , A1.DEPTCD                 ");
                //strSql.AppendLine("          , A1.DLOCK                  ");
                //strSql.AppendLine("          , B1.USRNM                  ");
                //strSql.AppendLine("          , B2.DEPT_CD                ");
                //strSql.AppendLine("          , B4.DEPT_NM                ");
                //strSql.AppendLine("          , B2.GRADE_CD               ");
                //strSql.AppendLine("          , B3.COM_NM AS JIKWI        ");
                //strSql.AppendLine("          , A1.STITL                  ");
                //strSql.AppendLine("          , A1.DOCTP                  ");
                //strSql.AppendLine("          , A2.COM_NM AS DOCNM        ");
                //strSql.AppendLine("          , A1.STATE                  ");
                //strSql.AppendLine("          , A3.COM_NM AS STATENM      ");
                //strSql.AppendLine("          , A1.FILNM                  ");
                //strSql.AppendLine("          , A1.RESON                  ");
                //strSql.AppendLine("          , A1.CUSER                  ");
                //strSql.AppendLine("          , A1.CDATE                  ");
                //strSql.AppendLine("          , A1.MUSER                  ");
                //strSql.AppendLine("          , A1.MDATE                  ");
                //strSql.AppendLine("          , REPLACE(STUFF(( SELECT '-' + C.USRNM + '(' + F.COM_NM + ')'          ");
                //strSql.AppendLine("                              FROM DOCT_SIGN D                                   ");
                //strSql.AppendLine("                              LEFT JOIN ZUSRLST C                                ");
                //strSql.AppendLine("                                ON D.PLNCD = C.USRCD                             ");
                //strSql.AppendLine("                              LEFT JOIN COM_BASE_CD F                            ");
                //strSql.AppendLine("                                ON D.APPTP = F.COM_CD                            ");
                //strSql.AppendLine("                               AND CD_GB = 'APPTP'                               ");
                //strSql.AppendLine("                             WHERE A1.SLINO = D.SLINO                            ");
                //strSql.AppendLine("                               AND D.GUBUN = '0'                                 ");
                //strSql.AppendLine("                               FOR XML PATH('')), 1, 1, ''),'-',' -> ') AS GLINE ");
                //strSql.AppendLine("          , A4.PLNCD AS NXTCD      ");
                //strSql.AppendLine("          , A5.PLNCD AS PRPCD      ");
                //strSql.AppendLine("          , A6.PLNCD AS FINCD      ");
                //strSql.AppendLine("       FROM DOCT_M A1              ");
                //strSql.AppendLine("       LEFT JOIN COM_BASE_CD A2    ");
                //strSql.AppendLine("         ON A1.DOCTP = A2.COM_CD   ");
                //strSql.AppendLine("        AND A2.CD_GB = 'DOCTP'     ");
                //strSql.AppendLine("       LEFT JOIN COM_BASE_CD A3    ");
                //strSql.AppendLine("         ON A1.STATE = A3.COM_CD   ");
                //strSql.AppendLine("        AND A3.CD_GB = 'RTSTATE'   ");
                //strSql.AppendLine("       LEFT JOIN TEMP2 A4          ");
                //strSql.AppendLine("         ON A1.SLINO = A4.SLINO    ");
                //strSql.AppendLine("       LEFT JOIN TEMP4 A5          ");
                //strSql.AppendLine("         ON A1.SLINO = A5.SLINO    ");
                //strSql.AppendLine("       LEFT JOIN TEMP7 A6          ");
                //strSql.AppendLine("         ON A1.SLINO = A6.SLINO    ");
                //strSql.AppendLine("       LEFT JOIN ZUSRLST B1        ");
                //strSql.AppendLine("         ON A1.PLNCD = B1.USRCD    ");
                //strSql.AppendLine("       LEFT JOIN HR_EMP_BASIS B2   ");
                //strSql.AppendLine("         ON B1.INSANO = B2.EMP_ID  ");
                //strSql.AppendLine("       LEFT JOIN COM_BASE_CD B3    ");
                //strSql.AppendLine("         ON B2.GRADE_CD = B3.COM_CD");
                //strSql.AppendLine("        AND B3.CD_GB = 'GRADE_CD'  ");
                //strSql.AppendLine("       LEFT JOIN ACC_DEPT_CD B4    ");
                //strSql.AppendLine("         ON B2.DEPT_CD = B4.DEPT_CD");
                //strSql.AppendLine("      WHERE 1 = 1                  ");
                //strSql.AppendLine("        AND A1.TDATE BETWEEN '" + fdate + "' AND '" + tdate + "'");
                //if (!string.IsNullOrEmpty(findword))
                //{
                //    if (findIdx == 0)
                //    {
                //        strSql.AppendLine("AND A1.SLINO LIKE '%' + '" + findword + "' + '%'");
                //    }
                //    else if (findIdx == 1)
                //    {
                //        strSql.AppendLine("AND A1.STITL LIKE '%' + '" + findword + "' + '%'");
                //    }
                //    else if (findIdx == 2)
                //    {
                //        strSql.AppendLine("AND B1.USRNM LIKE '%' + '" + findword + "' + '%'");
                //    }
                //    else if (findIdx == 3)
                //    {
                //        strSql.AppendLine("AND A2.COM_NM LIKE '%' + '" + findword + "' + '%'");
                //    }
                //}

                //strSql.AppendLine("        AND A1.STATE != '4'                                "); //완료가 아닌거만 조회 (요청) 2022-12-07

                //strSql.AppendLine(" ), TEMP9 AS(                                             ");
                //strSql.AppendLine("    --본인 글 및 결재라인인거 (최종결재자 제거요청 2022-12-07)                             ");
                //strSql.AppendLine("    SELECT A1.*                                           ");
                //strSql.AppendLine("      FROM TEMP8 A1                                       ");
                //strSql.AppendLine("     WHERE(A1.PLNCD = '" + sUsrcd + "' OR A1.NXTCD = '" + sUsrcd + "' OR A1.PRPCD = '" + sUsrcd + "' /*OR A1.FINCD = '" + sUsrcd + "'*/)    ");
                //strSql.AppendLine(" )                           ");
                //strSql.AppendLine("                             ");
                //strSql.AppendLine(" SELECT DISTINCT A1.*        ");
                //strSql.AppendLine("   FROM TEMP9 A1             ");
                //strSql.AppendLine("  ORDER BY A1.TDATE, A1.SLINO");
                #endregion

                //#0001
                strSql.AppendLine(" WITH TEMP1 AS(                       ");
                strSql.AppendLine("        SELECT SLINO                  ");
                strSql.AppendLine("             , GUBUN                  ");
                strSql.AppendLine("             , MIN(SEQNO) AS SEQNO    ");
                strSql.AppendLine("          FROM DOCT_SIGN              ");
                strSql.AppendLine("         WHERE APPTP = '3'            ");
                strSql.AppendLine("           AND GUBUN = '0'            ");
                strSql.AppendLine("         GROUP BY SLINO, GUBUN        ");
                strSql.AppendLine(" ), TEMP2 AS(--다음결재자                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.GUBUN                  ");
                strSql.AppendLine("          , A1.SEQNO                  ");
                strSql.AppendLine("          , A2.PLNCD                  ");
                strSql.AppendLine("       FROM TEMP1 A1                  ");
                strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                strSql.AppendLine(" ), TEMP3 AS(                         ");
                strSql.AppendLine("     SELECT SLINO                     ");
                strSql.AppendLine("          , GUBUN                     ");
                strSql.AppendLine("          , MAX(SEQNO) AS SEQNO       ");
                strSql.AppendLine("       FROM DOCT_SIGN                 ");
                strSql.AppendLine("      WHERE APPTP = '1'               ");
                strSql.AppendLine("        AND GUBUN = '0'               ");
                strSql.AppendLine("      GROUP BY SLINO, GUBUN           ");
                strSql.AppendLine(" ), TEMP4 AS(--이전결재자                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.GUBUN                  ");
                strSql.AppendLine("          , A1.SEQNO                  ");
                strSql.AppendLine("          , A2.PLNCD                  ");
                strSql.AppendLine("       FROM TEMP3 A1                  ");
                strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                strSql.AppendLine(" ), TEMP6 AS(                         ");
                strSql.AppendLine("     SELECT SLINO                     ");
                strSql.AppendLine("          , GUBUN                     ");
                strSql.AppendLine("          , MAX(SEQNO) AS SEQNO       ");
                strSql.AppendLine("       FROM DOCT_SIGN                 ");
                strSql.AppendLine("      WHERE GUBUN = '0'               ");
                strSql.AppendLine("      GROUP BY SLINO, GUBUN           ");
                strSql.AppendLine(" ), TEMP7 AS(--최종결재자                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.GUBUN                  ");
                strSql.AppendLine("          , A1.SEQNO                  ");
                strSql.AppendLine("          , A2.PLNCD                  ");
                strSql.AppendLine("       FROM TEMP6 A1                  ");
                strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                strSql.AppendLine(" ), TEMP8 AS(                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.TDATE                  ");
                strSql.AppendLine("          , A1.PLNCD                  ");
                strSql.AppendLine("          , A1.DEPTCD                 ");
                strSql.AppendLine("          , A1.DLOCK                  ");
                strSql.AppendLine("          , B1.USRNM                  ");
                strSql.AppendLine("          , B2.DEPT_CD                ");
                strSql.AppendLine("          , B4.DEPT_NM                ");
                strSql.AppendLine("          , B2.GRADE_CD               ");
                strSql.AppendLine("          , B3.COM_NM AS JIKWI        ");
                strSql.AppendLine("          , A1.STITL                  ");
                strSql.AppendLine("          , A1.DOCTP                  ");
                strSql.AppendLine("          , A2.DOCNM                  ");
                strSql.AppendLine("          , A1.STATE                  ");
                strSql.AppendLine("          , A3.COM_NM AS STATENM      ");
                strSql.AppendLine("          , A1.FILNM                  ");
                strSql.AppendLine("          , A1.RESON                  ");
                strSql.AppendLine("          , A1.CUSER                  ");
                strSql.AppendLine("          , A1.CDATE                  ");
                strSql.AppendLine("          , A1.MUSER                  ");
                strSql.AppendLine("          , A1.MDATE                  ");
                strSql.AppendLine("          , A1.CUSTOM                 ");
                strSql.AppendLine("          , A1.PAY                    ");
                strSql.AppendLine("          , REPLACE(STUFF(( SELECT '-' + C.USRNM + '(' + F.COM_NM + ')'          ");
                strSql.AppendLine("                              FROM DOCT_SIGN D                                   ");
                strSql.AppendLine("                              LEFT JOIN ZUSRLST C                                ");
                strSql.AppendLine("                                ON D.PLNCD = C.USRCD                             ");
                strSql.AppendLine("                              LEFT JOIN COM_BASE_CD F                            ");
                strSql.AppendLine("                                ON D.APPTP = F.COM_CD                            ");
                strSql.AppendLine("                               AND CD_GB = 'APPTP'                               ");
                strSql.AppendLine("                             WHERE A1.SLINO = D.SLINO                            ");
                strSql.AppendLine("                               AND D.GUBUN = '0'                                 ");
                strSql.AppendLine("                               FOR XML PATH('')), 1, 1, ''),'-',' -> ') AS GLINE ");
                //
                strSql.AppendLine("          , STUFF((SELECT '/' + CONVERT(VARCHAR, D.PLNCD)        ");
                strSql.AppendLine("              FROM DOCT_SIGN D                                   ");
                strSql.AppendLine("              LEFT JOIN ZUSRLST C                                ");
                strSql.AppendLine("                ON D.PLNCD = C.USRCD                             ");
                strSql.AppendLine("             WHERE A1.SLINO = D.SLINO                            ");
                strSql.AppendLine("               AND D.GUBUN = '0'                                 ");
                strSql.AppendLine("               FOR XML PATH('')), 1, 1, '') AS GLINE_CD          ");
                //
                strSql.AppendLine("          , STUFF((SELECT '/' + CONVERT(VARCHAR, D.PLNCD)                                              ");
                strSql.AppendLine("                     FROM DOCT_SIGN D                                                                  ");
                strSql.AppendLine("                     LEFT JOIN ZUSRLST C                                                               ");
                strSql.AppendLine("                       ON D.PLNCD = C.USRCD                                                            ");
                strSql.AppendLine("                    WHERE A1.SLINO = D.SLINO                                                           ");
                strSql.AppendLine("                      AND D.GUBUN = '0'                                                                ");
                strSql.AppendLine("                      AND D.APPTP IN ('1', '2')                                                        ");
                strSql.AppendLine("                      FOR XML PATH('')), 1, 1, '') AS GLINE_CD2 --결재라인에서 결재 완료,반려 한사람만 ");

                strSql.AppendLine("          , STUFF((SELECT ',' + C.USRNM   ");
                strSql.AppendLine("                     FROM DOCT_SIGN D                            ");
                strSql.AppendLine("                     LEFT JOIN ZUSRLST C                         ");
                strSql.AppendLine("                       ON D.PLNCD = C.USRCD                      ");
                strSql.AppendLine("                    WHERE A1.SLINO = D.SLINO                     ");
                strSql.AppendLine("                      AND D.GUBUN = '1'                          ");
                strSql.AppendLine("                      FOR XML PATH('')), 1, 1, '') AS CHAM       ");
                strSql.AppendLine("          , STUFF((SELECT '/' + CONVERT(VARCHAR, D.PLNCD)        ");
                strSql.AppendLine("                        FROM DOCT_SIGN D                         ");
                strSql.AppendLine("                        LEFT JOIN ZUSRLST C                      ");
                strSql.AppendLine("                          ON D.PLNCD = C.USRCD                   ");
                strSql.AppendLine("                       WHERE A1.SLINO = D.SLINO                  ");
                strSql.AppendLine("                         AND D.GUBUN = '1'                       ");
                strSql.AppendLine("                         FOR XML PATH('')), 1, 1, '') AS CHAM_CD ");
                //
                strSql.AppendLine("          , A4.PLNCD AS NXTCD      ");
                strSql.AppendLine("          , A5.PLNCD AS PRPCD      ");
                strSql.AppendLine("          , A6.PLNCD AS FINCD      ");
                strSql.AppendLine("       FROM DOCT_M A1              ");
                strSql.AppendLine("       LEFT JOIN DOCT_K A2         ");
                strSql.AppendLine("         ON A1.DOCTP = A2.DOCTP    "); 
                strSql.AppendLine("       LEFT JOIN COM_BASE_CD A3    ");
                strSql.AppendLine("         ON A1.STATE = A3.COM_CD   ");
                strSql.AppendLine("        AND A3.CD_GB = 'RTSTATE'   ");
                strSql.AppendLine("       LEFT JOIN TEMP2 A4          ");
                strSql.AppendLine("         ON A1.SLINO = A4.SLINO    ");
                strSql.AppendLine("       LEFT JOIN TEMP4 A5          ");
                strSql.AppendLine("         ON A1.SLINO = A5.SLINO    ");
                strSql.AppendLine("       LEFT JOIN TEMP7 A6          ");
                strSql.AppendLine("         ON A1.SLINO = A6.SLINO    ");
                strSql.AppendLine("       LEFT JOIN ZUSRLST B1        ");
                strSql.AppendLine("         ON A1.PLNCD = B1.USRCD    ");
                strSql.AppendLine("       LEFT JOIN HR_EMP_BASIS B2   ");
                strSql.AppendLine("         ON B1.INSANO = B2.EMP_ID  ");
                strSql.AppendLine("       LEFT JOIN COM_BASE_CD B3    ");
                strSql.AppendLine("         ON B2.GRADE_CD = B3.COM_CD");
                strSql.AppendLine("        AND B3.CD_GB = 'GRADE_CD'  ");
                strSql.AppendLine("       LEFT JOIN ACC_DEPT_CD B4    ");
                strSql.AppendLine("         ON B2.DEPT_CD = B4.DEPT_CD");
                strSql.AppendLine("      WHERE 1 = 1                  ");
                strSql.AppendLine("        AND A1.STATE NOT IN ('4','5')                              ");
                strSql.AppendLine(" ), TEMP9 AS(                                            ");
                strSql.AppendLine("    --결재라인인거  ");
                strSql.AppendLine("    SELECT A1.*                                          ");
                strSql.AppendLine("      FROM TEMP8 A1                                      ");
                //strSql.AppendLine("      LEFT JOIN( --결재라인 전체                         ");
                //strSql.AppendLine("                 SELECT A1.SLINO                         ");
                //strSql.AppendLine("                      , STUFF((SELECT ',' + CONVERT(VARCHAR, PLNCD)    ");
                //strSql.AppendLine("                                 FROM DOCT_SIGN                        ");
                //strSql.AppendLine("                                WHERE SLINO = A1.SLINO                 ");
                //strSql.AppendLine("                                  AND APPTP <> '3'");
                //strSql.AppendLine("                                  FOR XML PATH('')),1,1,'') AS SINLIN  ");
                //strSql.AppendLine("                   FROM DOCT_SIGN A1                                   ");
                //strSql.AppendLine("                  GROUP BY SLINO ) A2                                  ");
                //strSql.AppendLine("        ON A1.SLINO = A2.SLINO                                          ");
                //strSql.AppendLine("       AND (A2.SINLIN LIKE '%" + sUsrcd + "%' OR A1.NXTCD = '" + sUsrcd + "')                            ");
                //strSql.AppendLine("     WHERE A2.SLINO IS NOT NULL                                         ");
                strSql.AppendLine("      WHERE (A1.PLNCD = '" + sUsrcd + "' OR A1.GLINE_CD2 LIKE '%" + sUsrcd + "%' OR A1.NXTCD = '" + sUsrcd + "' OR A1.CHAM_CD LIKE '%" + sUsrcd + "%')");
                strSql.AppendLine(" )                                                                      ");
                strSql.AppendLine("                                                                        ");
                strSql.AppendLine(" SELECT DISTINCT A1.*                                                   ");
                strSql.AppendLine("   FROM TEMP9 A1                                                        ");
                strSql.AppendLine("  ORDER BY A1.TDATE, A1.SLINO                                           ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                GridRetr.DataSource = dt;

                if (!string.IsNullOrEmpty(_SLINO))
                {
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColSlino, _SLINO);
                    _SLINO = string.Empty;

                    //SetBtnEnables();
                }

                if (dt != null && dt.Rows.Count > 0 && focseIdx == 0)
                {
                    GridViewRetr_FocusedRowChanged(null, null);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        //부서결재문서 조회
        private void GetGridRetr3()
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                int focseIdx = GridViewRetr.FocusedRowHandle;

                string fdate = DateFrom.DateTime.ToString("yyyy-MM-dd");
                string tdate = DateTo.DateTime.ToString("yyyy-MM-dd");
                int findIdx = CboFindObj.SelectedIndex;
                string findword = TxtFindWord.EditValue?.ToString();
                string sUsrcd = FmMainToolBar2.drUser["USRCD"]?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" WITH TEMP1 AS(                       ");
                strSql.AppendLine("        SELECT SLINO                  ");
                strSql.AppendLine("             , GUBUN                  ");
                strSql.AppendLine("             , MIN(SEQNO) AS SEQNO    ");
                strSql.AppendLine("          FROM DOCT_SIGN              ");
                strSql.AppendLine("         WHERE APPTP = '3'            ");
                strSql.AppendLine("           AND GUBUN = '0'            ");
                strSql.AppendLine("         GROUP BY SLINO, GUBUN        ");
                strSql.AppendLine(" ), TEMP2 AS(--다음결재자                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.GUBUN                  ");
                strSql.AppendLine("          , A1.SEQNO                  ");
                strSql.AppendLine("          , A2.PLNCD                  ");
                strSql.AppendLine("       FROM TEMP1 A1                  ");
                strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                strSql.AppendLine(" ), TEMP3 AS(                         ");
                strSql.AppendLine("     SELECT SLINO                     ");
                strSql.AppendLine("          , GUBUN                     ");
                strSql.AppendLine("          , MAX(SEQNO) AS SEQNO       ");
                strSql.AppendLine("       FROM DOCT_SIGN                 ");
                strSql.AppendLine("      WHERE APPTP = '1'               ");
                strSql.AppendLine("        AND GUBUN = '0'               ");
                strSql.AppendLine("      GROUP BY SLINO, GUBUN           ");
                strSql.AppendLine(" ), TEMP4 AS(--이전결재자                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.GUBUN                  ");
                strSql.AppendLine("          , A1.SEQNO                  ");
                strSql.AppendLine("          , A2.PLNCD                  ");
                strSql.AppendLine("       FROM TEMP3 A1                  ");
                strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                strSql.AppendLine(" ), TEMP6 AS(                         ");
                strSql.AppendLine("     SELECT SLINO                     ");
                strSql.AppendLine("          , GUBUN                     ");
                strSql.AppendLine("          , MAX(SEQNO) AS SEQNO       ");
                strSql.AppendLine("       FROM DOCT_SIGN                 ");
                strSql.AppendLine("      WHERE GUBUN = '0'               ");
                strSql.AppendLine("      GROUP BY SLINO, GUBUN           ");
                strSql.AppendLine(" ), TEMP7 AS(--최종결재자                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A1.GUBUN                  ");
                strSql.AppendLine("          , A1.SEQNO                  ");
                strSql.AppendLine("          , A2.PLNCD                  ");
                strSql.AppendLine("          , A2.APPDT                  "); //#0003
                strSql.AppendLine("       FROM TEMP6 A1                  ");
                strSql.AppendLine("       LEFT JOIN DOCT_SIGN A2         ");
                strSql.AppendLine("         ON A1.SLINO = A2.SLINO       ");
                strSql.AppendLine("        AND A1.GUBUN = A2.GUBUN       ");
                strSql.AppendLine("        AND A1.SEQNO = A2.SEQNO       ");
                strSql.AppendLine(" ), TEMP8 AS(                         ");
                strSql.AppendLine("     SELECT A1.SLINO                  ");
                strSql.AppendLine("          , A6.APPDT                  "); //#0003
                strSql.AppendLine("          , A1.TDATE                  ");
                strSql.AppendLine("          , A1.PLNCD                  ");
                strSql.AppendLine("          , A1.DEPTCD                 ");
                strSql.AppendLine("          , A1.DLOCK                  ");
                strSql.AppendLine("          , A1.CCHEK                  ");
                strSql.AppendLine("          , B1.USRNM                  ");
                strSql.AppendLine("          , B2.DEPT_CD                ");
                strSql.AppendLine("          , B4.DEPT_NM                ");
                strSql.AppendLine("          , B2.GRADE_CD               ");
                strSql.AppendLine("          , B3.COM_NM AS JIKWI        ");
                strSql.AppendLine("          , A1.STITL                  ");
                strSql.AppendLine("          , A1.DOCTP                  ");
                strSql.AppendLine("          , A2.DOCNM        ");
                strSql.AppendLine("          , A1.STATE                  ");
                strSql.AppendLine("          , A3.COM_NM AS STATENM      ");
                strSql.AppendLine("          , A1.FILNM                  ");
                strSql.AppendLine("          , A1.RESON                  ");
                strSql.AppendLine("          , A1.CUSER                  ");
                strSql.AppendLine("          , A1.CDATE                  ");
                strSql.AppendLine("          , A1.MUSER                  ");
                strSql.AppendLine("          , A1.MDATE                  ");
                strSql.AppendLine("          , A1.CUSTOM                 ");
                strSql.AppendLine("          , A1.PAY                    ");
                strSql.AppendLine("          , REPLACE(STUFF(( SELECT '-' + C.USRNM + '(' + F.COM_NM + ')'          ");
                strSql.AppendLine("                              FROM DOCT_SIGN D                                   ");
                strSql.AppendLine("                              LEFT JOIN ZUSRLST C                                ");
                strSql.AppendLine("                                ON D.PLNCD = C.USRCD                             ");
                strSql.AppendLine("                              LEFT JOIN COM_BASE_CD F                            ");
                strSql.AppendLine("                                ON D.APPTP = F.COM_CD                            ");
                strSql.AppendLine("                               AND CD_GB = 'APPTP'                               ");
                strSql.AppendLine("                             WHERE A1.SLINO = D.SLINO                            ");
                strSql.AppendLine("                               AND D.GUBUN = '0'                                 ");
                strSql.AppendLine("                               FOR XML PATH('')), 1, 1, ''),'-',' -> ') AS GLINE ");
                strSql.AppendLine("          , STUFF((SELECT '/' + CONVERT(VARCHAR, D.PLNCD)");
                strSql.AppendLine("                        FROM DOCT_SIGN D                 ");
                strSql.AppendLine("                        LEFT JOIN ZUSRLST C              ");
                strSql.AppendLine("                          ON D.PLNCD = C.USRCD           ");
                strSql.AppendLine("                       WHERE A1.SLINO = D.SLINO          ");
                strSql.AppendLine("                         AND D.GUBUN = '0'               ");
                strSql.AppendLine("                         FOR XML PATH('')), 1, 1, '') AS GLINE_CD");
                //
                strSql.AppendLine("          , STUFF((SELECT ',' + C.USRNM ");
                strSql.AppendLine("                     FROM DOCT_SIGN D                          ");
                strSql.AppendLine("                     LEFT JOIN ZUSRLST C                       ");
                strSql.AppendLine("                       ON D.PLNCD = C.USRCD                    ");
                strSql.AppendLine("                    WHERE A1.SLINO = D.SLINO                   ");
                strSql.AppendLine("                      AND D.GUBUN = '1'                        ");
                strSql.AppendLine("                      FOR XML PATH('')), 1, 1, '') AS CHAM     ");
                strSql.AppendLine("          , STUFF((SELECT '/' + CONVERT(VARCHAR, D.PLNCD)");
                strSql.AppendLine("                        FROM DOCT_SIGN D                 ");
                strSql.AppendLine("                        LEFT JOIN ZUSRLST C              ");
                strSql.AppendLine("                          ON D.PLNCD = C.USRCD           ");
                strSql.AppendLine("                       WHERE A1.SLINO = D.SLINO          ");
                strSql.AppendLine("                         AND D.GUBUN = '1'               ");
                strSql.AppendLine("                         FOR XML PATH('')), 1, 1, '') AS CHAM_CD");
                //
                strSql.AppendLine("          , A4.PLNCD AS NXTCD      ");
                strSql.AppendLine("          , A5.PLNCD AS PRPCD      ");
                strSql.AppendLine("          , A6.PLNCD AS FINCD      ");
                strSql.AppendLine("       FROM DOCT_M A1              ");
                strSql.AppendLine("       LEFT JOIN DOCT_K A2     ");
                strSql.AppendLine("         ON A1.DOCTP = A2.DOCTP");
                strSql.AppendLine("       LEFT JOIN COM_BASE_CD A3    ");
                strSql.AppendLine("         ON A1.STATE = A3.COM_CD   ");
                strSql.AppendLine("        AND A3.CD_GB = 'RTSTATE'   ");
                strSql.AppendLine("       LEFT JOIN TEMP2 A4          ");
                strSql.AppendLine("         ON A1.SLINO = A4.SLINO    ");
                strSql.AppendLine("       LEFT JOIN TEMP4 A5          ");
                strSql.AppendLine("         ON A1.SLINO = A5.SLINO    ");
                strSql.AppendLine("       LEFT JOIN TEMP7 A6          ");
                strSql.AppendLine("         ON A1.SLINO = A6.SLINO    ");
                strSql.AppendLine("       LEFT JOIN ZUSRLST B1        ");
                strSql.AppendLine("         ON A1.PLNCD = B1.USRCD    ");
                strSql.AppendLine("       LEFT JOIN HR_EMP_BASIS B2   ");
                strSql.AppendLine("         ON B1.INSANO = B2.EMP_ID  ");
                strSql.AppendLine("       LEFT JOIN COM_BASE_CD B3    ");
                strSql.AppendLine("         ON B2.GRADE_CD = B3.COM_CD");
                strSql.AppendLine("        AND B3.CD_GB = 'GRADE_CD'  ");
                strSql.AppendLine("       LEFT JOIN ACC_DEPT_CD B4    ");
                strSql.AppendLine("         ON B2.DEPT_CD = B4.DEPT_CD");
                strSql.AppendLine("      WHERE 1 = 1                  ");
                strSql.AppendLine("        AND A1.TDATE BETWEEN '" + fdate + "' AND '" + tdate + "'");
                strSql.AppendLine(" ), TEMP9 AS(                                             ");
                strSql.AppendLine("    --부서 전체(본인문서포함)                             ");
                strSql.AppendLine("    SELECT A1.*                                           ");
                strSql.AppendLine("      FROM TEMP8 A1                                       ");
                strSql.AppendLine("     WHERE A1.STATE IN('4','5')");//완료인거만 조회(요청) 2022-12-07

                //#0001
                if (!sUsrcd.Equals("1001"))
                {
                    strSql.AppendLine("       AND ((((SELECT Z2.DEPT_CD FROM ZUSRLST Z1 LEFT JOIN HR_EMP_BASIS Z2 ON Z1.INSANO = Z2.EMP_ID WHERE Z1.USRCD = '" + sUsrcd + "') = '2000') OR ((SELECT Z2.DEPT_CD FROM ZUSRLST Z1 LEFT JOIN HR_EMP_BASIS Z2 ON Z1.INSANO = Z2.EMP_ID WHERE Z1.USRCD =  '" + sUsrcd + "') = '1000') AND 1 = 1)");
                    strSql.AppendLine("           OR((SELECT Z2.DEPT_CD FROM ZUSRLST Z1 LEFT JOIN HR_EMP_BASIS Z2 ON Z1.INSANO = Z2.EMP_ID WHERE Z1.USRCD = '" + sUsrcd + "') != '2000'           ");
                    strSql.AppendLine("               AND A1.DEPTCD = (SELECT Z2.DEPT_CD FROM ZUSRLST Z1 LEFT JOIN HR_EMP_BASIS Z2 ON Z1.INSANO = Z2.EMP_ID WHERE Z1.USRCD = '" + sUsrcd + "') AND A1.DLOCK = 'N'))");
                }

                strSql.AppendLine(" )                           ");
                strSql.AppendLine("                             ");
                strSql.AppendLine(" SELECT DISTINCT A1.*        ");
                strSql.AppendLine("   FROM TEMP9 A1             ");
                //#0001
                if (!string.IsNullOrEmpty(findword))
                {
                    if (findIdx == 0)
                    {
                        strSql.AppendLine("WHERE A1.SLINO LIKE '%' + '" + findword + "' + '%'");
                    }
                    else if (findIdx == 1)
                    {
                        strSql.AppendLine("WHERE A1.STITL LIKE '%' + '" + findword + "' + '%'");
                    }
                    else if (findIdx == 2)
                    {
                        strSql.AppendLine("WHERE A1.USRNM LIKE '%' + '" + findword + "' + '%'");
                    }
                    else if (findIdx == 3)
                    {
                        strSql.AppendLine("WHERE A1.DOCNM LIKE '%' + '" + findword + "' + '%'");
                    }
                }
                //#0002 //#0003
                strSql.AppendLine("  ORDER BY A1.APPDT DESC, A1.TDATE DESC, A1.SLINO DESC");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                GridRetr3.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sSlino = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();

            GetDtlData(sSlino);
        }

        private void GridViewRetr3_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sSlino = GridViewRetr3.GetFocusedRowCellValue("SLINO")?.ToString();

            GetDtlData(sSlino);
        }

        private void GetDtlData(string sSlino)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                StringBuilder strSql = new StringBuilder();

                #region 첨부파일 기능 제거 요청 2022-07-19
                //첨부
                //strSql.Clear();
                //strSql.AppendLine(" SELECT SLINO      ");
                //strSql.AppendLine("      , SEQNO      ");
                //strSql.AppendLine("      , FILNM      ");
                //strSql.AppendLine("   FROM DOCT_F     ");
                //strSql.AppendLine("  WHERE SLINO = '" + sSlino + "' ");

                //DataTable dtAtt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                //GridRetr2.DataSource = dtAtt;
                #endregion

                #region 엑셀파일가져오기
                if (!string.IsNullOrEmpty(sSlino))
                {
                    strSql.Clear();
                    strSql.AppendLine(" SELECT EXCFIL     ");
                    strSql.AppendLine("   FROM DOCT_M     ");
                    strSql.AppendLine("  WHERE SLINO = '" + sSlino + "' ");

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    byte[] DownloadFile = null;

                    if (dt != null && !DBNull.Value.Equals(dt.Rows[0]["EXCFIL"]))
                    {
                        DownloadFile = (byte[])dt.Rows[0]["EXCFIL"];
                    }

                    if (DownloadFile == null)
                    {
                        int sheetCnt = spread.ActiveWorksheet.Workbook.Worksheets.Count;

                        if (!spread.Document.Worksheets[0].Name.Equals("EmptySheet1"))
                        {
                            spread.ActiveWorksheet.Workbook.Worksheets.Add("EmptySheet1");

                            for (int i = 0; i < sheetCnt; i++)
                            {
                                spread.ActiveWorksheet.Workbook.Worksheets.RemoveAt(0);
                            }
                        }
                    }
                    else
                    {
                        spread.LoadDocument(DownloadFile, DocumentFormat.Xlsm);
                    }
                }
                else
                {
                    int sheetCnt = spread.ActiveWorksheet.Workbook.Worksheets.Count;

                    if (!spread.Document.Worksheets[0].Name.Equals("EmptySheet1"))
                    {
                        spread.ActiveWorksheet.Workbook.Worksheets.Add("EmptySheet1");

                        for (int i = 0; i < sheetCnt; i++)
                        {
                            spread.ActiveWorksheet.Workbook.Worksheets.RemoveAt(0);
                        }
                    }
                }
                #endregion

                #region 사인 가져오기

                //사인 리셋
                ResetSing();

                if (!string.IsNullOrEmpty(sSlino))
                {
                    strSql.Clear();
                    strSql.AppendLine(" SELECT A1.SEQNO                       ");
                    strSql.AppendLine("      , B2.EMP_NM                      ");
                    strSql.AppendLine("      , A1.APPTP ");
                    strSql.AppendLine("      , CONVERT(VARCHAR(5),CONVERT(DATE,A1.APPDT),101) AS APPDT");
                    strSql.AppendLine("      , B3.COM_NM AS JIKWI             ");
                    strSql.AppendLine("      , B2.SINIMG                      ");
                    strSql.AppendLine("   FROM DOCT_SIGN A1                   ");
                    strSql.AppendLine("   LEFT JOIN ZUSRLST B1                ");
                    strSql.AppendLine("     ON A1.PLNCD = B1.USRCD            ");
                    strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B2           ");
                    strSql.AppendLine("     ON B1.INSANO = B2.EMP_ID          ");
                    strSql.AppendLine("   LEFT JOIN COM_BASE_CD B3            ");
                    strSql.AppendLine("     ON B2.GRADE_CD = B3.COM_CD        ");
                    strSql.AppendLine("    AND B3.CD_GB = 'GRADE_CD'          ");
                    strSql.AppendLine("  WHERE A1.SLINO = '" + sSlino + "'");
                    strSql.AppendLine("    AND A1.GUBUN = '0'                 ");
                    strSql.AppendLine("  ORDER BY A1.SEQNO                    ");

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    if (dt != null)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string sNm = dt.Rows[i]["EMP_NM"]?.ToString();
                            string sAppdt = dt.Rows[i]["APPDT"]?.ToString();
                            string sAPPTP = dt.Rows[i]["APPTP"]?.ToString();
                            string sJikwi = dt.Rows[i]["JIKWI"]?.ToString();
                            byte[] bSINIMG = null;
                            Image iSingimg = null;

                            if (dt != null && !DBNull.Value.Equals(dt.Rows[i]["SINIMG"]))
                            {
                                bSINIMG = (byte[])dt.Rows[i]["SINIMG"];

                                MemoryStream ms = new MemoryStream(bSINIMG);

                                iSingimg = Image.FromStream(ms);
                            }

                            switch (i)
                            {
                                case 0:
                                    laNm0.Text = sJikwi;
                                    laDt0.Text = sAppdt;
                                    if (!string.IsNullOrEmpty(sAPPTP) && sAPPTP.Equals("1"))//승인일때
                                        picSign0.EditValue = iSingimg;
                                    break;
                                case 1:
                                    laNm1.Text = sJikwi;
                                    laDt1.Text = sAppdt;
                                    if (!string.IsNullOrEmpty(sAPPTP) && sAPPTP.Equals("1"))
                                        picSign1.EditValue = iSingimg;
                                    break;
                                case 2:
                                    laNm2.Text = sJikwi;
                                    laDt2.Text = sAppdt;
                                    if (!string.IsNullOrEmpty(sAPPTP) && sAPPTP.Equals("1"))
                                        picSign2.EditValue = iSingimg;
                                    break;
                                case 3:
                                    laNm3.Text = sJikwi;
                                    laDt3.Text = sAppdt;
                                    if (!string.IsNullOrEmpty(sAPPTP) && sAPPTP.Equals("1"))
                                        picSign3.EditValue = iSingimg;
                                    break;
                                case 4:
                                    laNm4.Text = sJikwi;
                                    laDt4.Text = sAppdt;
                                    if (!string.IsNullOrEmpty(sAPPTP) && sAPPTP.Equals("1"))
                                        picSign4.EditValue = iSingimg;
                                    break;
                                case 5:
                                    laNm5.Text = sJikwi;
                                    laDt5.Text = sAppdt;
                                    if (!string.IsNullOrEmpty(sAPPTP) && sAPPTP.Equals("1"))
                                        picSign5.EditValue = iSingimg;
                                    break;
                            }
                        }
                        SetSignLayout(dt.Rows.Count);
                    }

                }
                #endregion

                SetBtnEnables();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }

        //사인 표시
        private void SetSignLayout(int cnt)
        {
            switch (cnt)
            {
                case 1:
                    laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;
                case 2:
                    laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;
                case 3:
                    laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;
                case 4:
                    laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;
                case 5:
                    laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;
                case 6:
                    laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    break;
            }
        }

        //사인리셋
        private void ResetSing()
        {
            laNm0.Text = "";
            laDt0.Text = "";
            picSign0.EditValue = null;

            laNm1.Text = "";
            laDt1.Text = "";
            picSign1.EditValue = null;

            laNm2.Text = "";
            laDt2.Text = "";
            picSign2.EditValue = null;

            laNm3.Text = "";
            laDt3.Text = "";
            picSign3.EditValue = null;

            laNm4.Text = "";
            laDt4.Text = "";
            picSign4.EditValue = null;

            laNm5.Text = "";
            laDt5.Text = "";
            picSign5.EditValue = null;

            laySign1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            laySign2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            laySign3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            laySign4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            laySign5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            laySign6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void SetBtnEnables()
        {
            string sNXTCD = GridViewRetr.GetFocusedRowCellValue("NXTCD")?.ToString();
            string sPRPCD = GridViewRetr.GetFocusedRowCellValue("PRPCD")?.ToString();
            string sPLNCD = GridViewRetr.GetFocusedRowCellValue("PLNCD")?.ToString();
            string sState = GridViewRetr.GetFocusedRowCellValue("STATE")?.ToString();

            string sUsrcd = FmMainToolBar2.drUser["USRCD"]?.ToString();

            #region 승인,반려 버튼 활성화/비활성화
            //다음 결재자와 접속자가 같을때 승인버튼 활성화, 승인취소 비활성화
            if (!string.IsNullOrEmpty(sNXTCD) && sNXTCD.Equals(sUsrcd))
            {
                layoutControlItem13.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;//승인버튼 활성화
                layoutControlItem14.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;//반려활성화
                //layoutControlItem20.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;//승인취소 비활성화
            }
            else
            {
                layoutControlItem13.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;//승인버튼 비활성화
                layoutControlItem14.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;//반려 비활성화
            }

            //이전 결재자와 접속자가 같을때 승인취소 활성화
            if (!string.IsNullOrEmpty(sPRPCD) && sPRPCD.Equals(sUsrcd))
            {
                layoutControlItem20.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;//승인취소 활성화
            }
            else
            {
                layoutControlItem20.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;//승인취소 비활성화
            }
            #endregion

            #region 본인글만 삭제
            if (!string.IsNullOrEmpty(sPLNCD) && sPLNCD.Equals(sUsrcd))
            {
                //작성중인 글만 삭제 가능
                if (sState.Equals("1"))
                {
                    BtnDel.Enabled = true;
                }
                else
                {
                    BtnDel.Enabled = false;
                }
            }
            else
            {
                BtnDel.Enabled = false;
            }
            #endregion
        }
        #endregion

        #region 추가
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            RptApplSystemP1 fm = new RptApplSystemP1();

            //fm.Owner = this;  //창 우선순위로 인해 변경
            fm.AddModifyGb = "ADD";
            fm.DataRowSendEvent += new RptApplSystemP1.SendDataHandler(GetAddData);
            fm.Show();
            //if (fm.ShowDialog() == DialogResult.OK)
            //{
            //    BtnRetr.PerformClick();
            //}
            //else
            //{
            //    //ExcelCloseAndDelete();
            //}

        }
        
        private void GetAddData(Dictionary<string, object> dicParams)
        {
            _SLINO = dicParams["SLINO"]?.ToString();
            BtnRetr.PerformClick();
        }
        #endregion

        #region 수정
        private void GridViewRetr_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                RptApplSystemP1 fm = new RptApplSystemP1();

                //fm.Owner = this;
                fm.AddModifyGb = "MOD";
                fm._SLINO = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();
                fm.DataRowSendEvent += new RptApplSystemP1.SendDataHandler(GetAddData);
                fm.Show();
                //if (fm.ShowDialog() == DialogResult.OK)
                //{
                //    BtnRetr.PerformClick();
                //}
            }
            else if (e.Clicks == 1)
            {
                string sSlino = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();

                GetDtlData(sSlino);
            }
        }
        #endregion

        #region 삭제
        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sSlino = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();
            string sDoctp = GridViewRetr.GetFocusedRowCellValue("DOCTP")?.ToString();

            if (string.IsNullOrEmpty(sSlino))
            {
                XtraMessageBox.Show("삭제할 항목이 없습니다.");
                return;
            }
            try
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT STATE     ");
                strSql.AppendLine("   FROM DOCT_M    ");
                strSql.AppendLine("  WHERE SLINO = '"+ sSlino + "'");

                DataTable chkState = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if(chkState != null)
                {
                    string sState = chkState.Rows[0]["STATE"]?.ToString();

                    if (sState.Equals("2"))
                    {
                        XtraMessageBox.Show("결재 진행중인 건은 삭제가 불가능 합니다.");
                        return;
                    }
                    else if (sState.Equals("3"))
                    {
                        XtraMessageBox.Show("결재 반려된 건은 삭제가 불가능 합니다.");
                        return;
                    }
                    else if (sState.Equals("4"))
                    {
                        XtraMessageBox.Show("결재 완료된 건은 삭제가 불가능 합니다.");
                        return;
                    }
                }

                int idx = GridViewRetr.GetFocusedDataSourceRowIndex();

                if (XtraMessageBox.Show("문서번호 : " + sSlino +
                     " \r\n선택된 항목을 삭제하시겠습니까? \r\n삭제한 데이터는 복구할 수 없습니다."
                   , "결재문서 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                #region 첨부파일 기능 삭제 요청 2022-07-19
                //strSql.Clear();
                //strSql.AppendLine("  SELECT SLINO ");
                //strSql.AppendLine("       , SEQNO  ");
                //strSql.AppendLine("       , FILNM  ");
                //strSql.AppendLine("    FROM DOCT_F ");
                //strSql.AppendLine("   WHERE SLINO = '" + sSlino + "'");

                //DataTable dtAtt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                //DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                //SqlCommand cmd = DBConn.dbCon.CreateCommand();
                //cmd.Transaction = DBConn.dbTran;

                //if (dtAtt != null && dtAtt.Rows.Count > 0)
                //{
                //    string AttchPath = sInitDir + "/" + sSlino +"/";

                //    for (int i = 0; i < dtAtt.Rows.Count; i++)
                //    {
                //        string sSeqno = dtAtt.Rows[i]["SEQNO", DataRowVersion.Original]?.ToString();
                //        string sFilnm = dtAtt.Rows[i]["FILNM", DataRowVersion.Original]?.ToString();

                //        strSql.Clear();
                //        strSql.AppendLine(" DELETE FROM DOCT_F");
                //        strSql.AppendLine("  WHERE SLINO = '" + sSlino + "' ");
                //        strSql.AppendLine("    AND SEQNO =" + sSeqno + " ");

                //        cmd.CommandType = CommandType.Text;
                //        cmd.CommandText = strSql.ToString();
                //        cmd.ExecuteNonQuery();

                //        if (!string.IsNullOrEmpty(sFilnm))
                //        {
                //            ComnEtcFunc.DeleteFTPFile(AttchPath + sFilnm, user, pw);
                //        }
                //    }
                //}
                #endregion

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM ALAMMGT                 ");
                strSql.AppendLine("  WHERE ALKEY1 = '"+ sSlino + "'");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM DOCT_M");
                strSql.AppendLine("     WHERE SLINO = '" + sSlino + "'  ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM DOCT_SIGN");
                strSql.AppendLine("  WHERE SLINO = '" + sSlino + "'");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = idx - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        #region 승인
        private void BtnAppl_Click(object sender, EventArgs e)
        {
            try
            {
                int idx = GridViewRetr.GetFocusedDataSourceRowIndex();
                string sDoctp = GridViewRetr.GetFocusedRowCellValue("DOCTP")?.ToString();
                string sSLINO = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();
                string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();
                string sState = string.Empty;

                if (string.IsNullOrEmpty(sSLINO))
                {
                    XtraMessageBox.Show("승인 처리할 문서를 선택해 주세요.");
                    return;
                }

                if (XtraMessageBox.Show("문서번호 : " + sSLINO +
                 " \r\n선택된 항목을 승인하시겠습니까?"
               , "결재문서 승인여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT STATE     ");
                strSql.AppendLine("   FROM DOCT_M    ");
                strSql.AppendLine("  WHERE SLINO = '" + sSLINO + "'");

                DataTable chkState = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (chkState != null)
                {
                    sState = chkState.Rows[0]["STATE"]?.ToString();

                    if (sState.Equals("4"))
                    {
                        XtraMessageBox.Show("이미 결재 승인이 완료 된 건 입니다.");
                        return;
                    }
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                
                dicParams.Clear();
                dicParams.Add("CMD", "SAVEOK");
                dicParams.Add("SLINO", sSLINO);
                dicParams.Add("SUSER", sUser);

                DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if(dtResult != null)
                {
                    string sResult = dtResult.Rows[0]["RESULT"]?.ToString();
                    string sMsg = dtResult.Rows[0]["MSG"]?.ToString();

                    if (sResult.Equals("1"))
                    {
                        /*
                         * 현재 문서의 마지막 결재자와 현재 결재자가 같을시 결재 진행 상태를 완료로함
                         * 현재 문서의 마지막 결재자 와 현재 결재자가 같지 않을시 결재 진행 상태를 진행중으로
                         */
                        if (!string.IsNullOrEmpty(sState) && sState != "4")
                        {
                            dicParams.Clear();
                            dicParams.Add("CMD", "OK_CHANGE");
                            dicParams.Add("SLINO", sSLINO);
                            dicParams.Add("SUSER", sUser);

                            DataTable dtResult2 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                            if (dtResult2.Rows.Count > 0)
                            {
                                string sEVcheck = dtResult2.Rows[0]["OUT_FINAL"]?.ToString();
                                if (sEVcheck.Equals("2"))// 마지막 결재 후 완료 상태 변경 후에 필요한 이벤트는  여기에다가 처리하면됨
                                {

                                }
                            }
                        }

                        XtraMessageBox.Show(sMsg);
                        _SLINO = sSLINO;
                        BtnRetr.PerformClick();
                    }
                    else
                    {
                        XtraMessageBox.Show(sMsg);
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
        }
        #endregion

        #region 승인취소
        private void BtnApplCan_Click(object sender, EventArgs e)
        {
            try
            {
                string sSlino = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();

                if (string.IsNullOrEmpty(sSlino))
                {
                    XtraMessageBox.Show("승인취소할 문서를 선택해주세요.");
                    return;
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "SAVECANCEL");
                dicParams.Add("SLINO", sSlino);
                dicParams.Add("SUSER", FmMainToolBar2.drUser["USRCD"]?.ToString());

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null && dt.Rows.Count > 0)
                {
                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());
                }

                _SLINO = sSlino;
                BtnRetr.PerformClick();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 반려
        public string _RESON;
        private void BtnReturn_Click(object sender, EventArgs e)
        {
            try
            {
                int idx = GridViewRetr.GetFocusedDataSourceRowIndex();
                string sDoctp = GridViewRetr.GetFocusedRowCellValue("DOCTP")?.ToString();
                string sSLINO = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();
                string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();
                string sPLNCD = GridViewRetr.GetFocusedRowCellValue("PLNCD")?.ToString();
                string sSTATE = string.Empty;

                if (string.IsNullOrEmpty(sSLINO))
                {
                    XtraMessageBox.Show("반려 처리할 문서를 선택해 주세요.");
                    return;
                }

                string sXtrMsg = string.Empty;
                //본인이 작성한 문서 반려시 최종반려상태로 변경
                if (sPLNCD.Equals(sUser))
                {
                    sSTATE = "5";//최종반려
                    sXtrMsg = "문서번호 : " + sSLINO +
                           " \r\n본인이 작성한 결재문서는 최종반려 처리됩니다." +
                           " \r\n선택된 항목을 최종반려하시겠습니까?";
                }
                else
                {
                    sSTATE = "3";
                    sXtrMsg = "문서번호 : " + sSLINO +
                           " \r\n선택된 항목을 반려하시겠습니까?";
                }

                if (XtraMessageBox.Show(sXtrMsg, "결재문서 반려여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                RptApplSystemP2 frm = new RptApplSystemP2();
                frm.Owner = this;

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Clear();
                    dicParams.Add("CMD", "SAVENO");
                    dicParams.Add("SLINO", sSLINO);
                    dicParams.Add("RESON", _RESON);
                    dicParams.Add("STATE", sSTATE);
                    dicParams.Add("SUSER", sUser);

                    DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                    if (dtResult != null)
                    {
                        string sResult = dtResult.Rows[0]["RESULT"]?.ToString();
                        string sMsg = dtResult.Rows[0]["MSG"]?.ToString();

                        if (sResult.Equals("1"))
                        {
                            XtraMessageBox.Show(sMsg);
                            _SLINO = sSLINO;
                            BtnRetr.PerformClick();
                        }
                        else
                        {
                            XtraMessageBox.Show(sMsg);
                        }
                    }
                }
                else
                {

                }
                
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
        }
        #endregion

        #region 양식다운로드
        private void BtnDnSample_Click(object sender, EventArgs e)
        {
            PopSelDoc frm = new PopSelDoc();

            frm.Owner = this;
            frm.ShowDialog();
        }
        #endregion

        #region 파일다운로드 -> 임시저장형태로 변경 
        private void BtnDownload_Click(object sender, EventArgs e)
        {
            string sSLINO = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();

            if (string.IsNullOrEmpty(sSLINO))
            {
                XtraMessageBox.Show("선택한 문서가 없습니다.");
                return;
            }

            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                Cursor = Cursors.WaitCursor;

                string downPath = attFolder + sSLINO;

                if (!Directory.Exists(downPath))
                {
                    Directory.CreateDirectory(downPath);
                }

                spread.Document.ExportToPdf(downPath+ "/"+ sSLINO + ".pdf");

                Process.Start(downPath + "/" + sSLINO + ".pdf");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
                return;
            }
            finally {
                Cursor = Cursors.Default;
                SplashScreenManager.CloseForm();
            }

        }

        //엑셀다운로드 버튼클릭 => 더블 클릭시 엑셀 다운로드로 변경
        private void GridViewRetr3_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2) 
            {
                DownExcelFile();
            }
            else if(e.Clicks == 1)
            {
                string sSlino = GridViewRetr3.GetFocusedRowCellValue("SLINO")?.ToString();

                GetDtlData(sSlino);
            }
        }

        private void DownExcelFile()
        {
            string sSLINO = GridViewRetr3.GetFocusedRowCellValue("SLINO")?.ToString();

            if (string.IsNullOrEmpty(sSLINO))
            {
                XtraMessageBox.Show("선택한 문서가 없습니다.");
                return;
            }

            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                Cursor = Cursors.WaitCursor;

                string downPath = attFolder + sSLINO;

                if (!Directory.Exists(downPath))
                {
                    Directory.CreateDirectory(downPath);
                }

                spread.SaveDocument(downPath + "/" + sSLINO + ".xlsx");

                Process.Start(downPath + "/" + sSLINO + ".xlsx");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
                SplashScreenManager.CloseForm();
            }
        }
        #endregion

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }
        private void GridViewRetrcheck_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                string sChk = GridViewRetr3.GetRowCellValue(e.RowHandle, gridColumn24)?.ToString();
                if (sChk == null)
                {
                    sChk = "";
                }
                if (sChk=='1'.ToString())
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else
                {
                    e.Appearance.BackColor = Color.White;
                }

            }
        }
        private void CboFindObj_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                ComboBoxEdit cbo = (ComboBoxEdit)sender;
                SettingTextEditAutoComplete(cbo.EditValue?.ToString());
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
            }
        }

        //텍스트박스 lookup
        private void SettingTextEditAutoComplete(string sGb)
        {
            AutoCompleteStringCollection collection = new AutoCompleteStringCollection();

            DataTable dt = new DataTable();

            if (sGb.Equals("문서번호"))
            {
                dt = GetLookupData("1");
            }
            if (sGb.Equals("제목"))
            {
                dt = GetLookupData("2");
            }
            if (sGb.Equals("작성자"))
            {
                dt = GetLookupData("3");
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                collection.Add(dt.Rows[i]["CD"]?.ToString());
            }

            TxtFindWord.MaskBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            TxtFindWord.MaskBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            TxtFindWord.MaskBox.AutoCompleteCustomSource = collection;
        }

        private void RdiState_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }
        #endregion

        #region 첨부파일 기능 제거 2022-07-19
        private void GridViewRetr2_RowClick(object sender, RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                string sSlino = GridViewRetr.GetFocusedRowCellValue("SLINO")?.ToString();
                string filenm = GridViewRetr2.GetFocusedRowCellValue("FILNM")?.ToString();

                string downPath = attFolder + sSlino;

                if (!Directory.Exists(downPath))
                {
                    Directory.CreateDirectory(downPath);
                }

                if (ComnEtcFunc.FTPFileDownload(sInitDir+ sSlino+"/" + filenm, downPath +"/"+ filenm, user, pw))
                {
                    //XtraMessageBox.Show("다운로드를 완료하였습니다.");
                    Process.Start(downPath +"/"+ filenm);
                }
            }
        }
        #endregion

        private void RptApplSystem_FormClosed(object sender, FormClosedEventArgs e)
        {
            //ExcelCloseAndDelete();
        }

        #region 엑셀 닫은 후 파일삭제 [미사용 2022-12-08]
        private void ExcelCloseAndDelete()
        {
            try
            {
                Process[] prcList = Process.GetProcessesByName("Excel");
                if (prcList.Length != 0)
                {
                    for (int i = 0; i < prcList.Length; i++)
                    {
                        Process prcessInfo = prcList[i];

                        string sProcessFileName = prcessInfo.MainWindowTitle;

                        if (sProcessFileName.Equals("전자결재 문서 작성.xlsx - Excel"))
                        {
                            prcessInfo.Kill();
                        }
                    }
                }

                Thread.Sleep(1000);

                FileInfo fileInfo = new FileInfo(tempDoctPath + @"/전자결재 문서 작성.xlsx");
                if (fileInfo.Exists)
                {
                    //파일삭제
                    fileInfo.Delete();
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

        }
        #endregion

        #region GridRow 마우스 우클릭 메뉴
        private void GridViewRetr3_PopupMenuShowing(object sender, PopupMenuShowingEventArgs e)
        {
            GridView view = sender as GridView;
            if (e.MenuType == DevExpress.XtraGrid.Views.Grid.GridMenuType.Row)
            {
                int rowHandle = e.HitInfo.RowHandle;
                e.Menu.Items.Clear();

                e.Menu.Items.Add(CreateSubMenuRows(view, rowHandle, "0"));
            }
        }

        class RowInfo
        {
            public RowInfo(GridView view, int rowHandle)
            {
                this.RowHandle = rowHandle;
                this.View = view;
            }
            public GridView View;
            public int RowHandle;
        }

        private DXMenuItem CreateSubMenuRows(GridView view, int rowHandle, string str)
        {
            string rowCommandCaption1 = "문서 복사";
            DXMenuItem item = null;
            if (str.Equals("0"))
            {
                item = new DXMenuItem(rowCommandCaption1, new EventHandler(DocCopy));
                item.Tag = new RowInfo(view, rowHandle);
                item.Enabled = view.IsDataRow(rowHandle);
                //subMenu.Items.Add(menuItemKjRow);
            }

            return item;
        }

        private void DocCopy(object sender, EventArgs e)
        {
            string SLINO = GridViewRetr3.GetFocusedRowCellValue("SLINO")?.ToString();
            string DOCTP = GridViewRetr3.GetFocusedRowCellValue("DOCTP")?.ToString();

            DXMenuItem menuItem = sender as DXMenuItem;
            RowInfo ri = menuItem.Tag as RowInfo;
            if (ri != null)
            {
                RptApplSystemP1 fm = new RptApplSystemP1();

                //fm.Owner = this;
                fm.AddModifyGb = "COPY";
                fm._SLINO = SLINO;
                fm._DOCTP = DOCTP;
                fm.DataRowSendEvent += new RptApplSystemP1.SendDataHandler(GetAddData);
                fm.Show();
                //if (fm.ShowDialog() == DialogResult.OK)
                //{
                //    BtnRetr.PerformClick();
                //}
            }
        }
        #endregion

        private void RptApplSystem_TextChanged(object sender, EventArgs e)
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

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void TxtFindWord_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            GridView view = new GridView();
            int[] selectedRows = null;

            selectedRows = GridViewRetr3.GetSelectedRows();
            view = GridViewRetr3;

            if (selectedRows.Length == 0)
            {
                XtraMessageBox.Show("저장하려는 행을 체크하세요.");
                return;
            }
            try
            {
                for (int i = 0; i < selectedRows.Length; i++)
                {
                    int iRowidx = selectedRows[i];
                    
                    string sSLINO = view.GetRowCellValue(iRowidx, "SLINO")?.ToString();
                    
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Clear();
                    dicParams.Add("CMD", "SAVEKU");
                    dicParams.Add("SLINO", sSLINO);

                    DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                    if (dtResult != null)
                    {
                        string sResult = dtResult.Rows[0]["RESULT"]?.ToString();
                        string sMsg = dtResult.Rows[0]["MSG"]?.ToString();

                        if (sResult.Equals("1"))
                        {
                            //XtraMessageBox.Show(sMsg);
                            //BtnRetr.PerformClick();
                        }
                        else
                        {
                            //XtraMessageBox.Show(sMsg);
                        }
                    }
                }
                GetGridRetr3();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
        }

        private void BtnCan_Click(object sender, EventArgs e)
        {
            GridView view = new GridView();
            int[] selectedRows = null;

            selectedRows = GridViewRetr3.GetSelectedRows();
            view = GridViewRetr3;

            if (selectedRows.Length == 0)
            {
                XtraMessageBox.Show("취소하려는 행을 체크하세요.");
                return;
            }
            try
            {
                for (int i = 0; i < selectedRows.Length; i++)
                {
                    int iRowidx = selectedRows[i];

                    string sSLINO = view.GetRowCellValue(iRowidx, "SLINO")?.ToString();

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Clear();
                    dicParams.Add("CMD", "CENKU");
                    dicParams.Add("SLINO", sSLINO);

                    DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                    if (dtResult != null)
                    {
                        string sResult = dtResult.Rows[0]["RESULT"]?.ToString();
                        string sMsg = dtResult.Rows[0]["MSG"]?.ToString();

                        if (sResult.Equals("1"))
                        {
                            //XtraMessageBox.Show(sMsg);
                            //BtnRetr.PerformClick();
                        }
                        else
                        {
                           // XtraMessageBox.Show(sMsg);
                        }
                    }
                }
                GetGridRetr3();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
        }
    }
}