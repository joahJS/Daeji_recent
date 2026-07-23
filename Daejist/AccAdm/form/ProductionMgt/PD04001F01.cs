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
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid;
using System.Net;
using System.IO;
using DevExpress.XtraGrid;
using DevExpress.XtraBars;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Data.SqlClient;

/* 
 * 작성일자 : 2021-03-24
 * 작성자   : 고혜성
 * 비고     : (현업요청)
 *             해당 화면은 기존 생산일보에서 검수내역을 등록하던 것에서
 *            별도 화면을 구성하여 검수전용으로 화면을 추가
 * ---------------------------------HISTORY---------------------------------
 * 
 * 수정일자 : 2021-04-06
 * 수정자   : 고혜성
 * Reference Key : #0001, #0002
 * 수정내용 : (현업요청)
 *            1. 입고만 조회되어야 함
 *            2. 등급코드의 대구분이 표기되어야함
 *            3. 검수내역조회 화면 기능이 등록에도 나타나야 함
 *            
 * 수정일자 : 2021-04-14
 * 수정자   : 고혜성
 * Reference Key : #0003
 * 수정내용 : (현업요청)
 *            1. 별도 엑셀양식파일을 로컬에서 참조하기 위하여 FTP 관련 세팅
 *            2. 해당 양식에 주간/월간 보고서 출력을 위하여 리포트 관련 로직 추가
 *            
 * 수정일자 : 2021-04-14
 * 수정자   : 고혜성
 * Reference Key : #0005
 * 수정내용 : (현업요청)
 *            1. 이미지 삭제로직 추가
 * 
 * 수정일자 : 2021-04-16
 * 수정자   : 고혜성
 * Reference Key : #0006
 * 수정내용 : (현업요청)
 *            1. 리포트 출력 시 총평은 검색기간 내 마지막 총평이 나오도록 쿼리수정
 *            
 * 수정일자 : 2021-04-22
 * 수정자   : 고혜성
 * Reference Key : #0007
 * 수정내용 : (현업요청)
 *            1. 검색기간 내 총평 MAX 값을 가져오도록 쿼리수정
 *          
 * 요청일자 : 2021-04-23(단톡방)
 * 수정일자 : 2021-04-23
 * 수정자   : 고혜성
 * Reference Key : #0008
 * 수정내용 : (현업요청)
 *          1. 거래처별 합계 탭 쿼리 中 거래처별 담당자 및 감량조정(검수감량 - 현재감량) 컬럼 추가
 *          2. 탭 전체 Filter 전체 True
 *     
 * 수정일자 : 2021-05-02
 * 수정자   : 고혜성
 * Reference Key : #0009
 * 수정내용 : (현업요청)
 *            1. 이미지 삭제는 해당 건 등록자와 접속자가 일치하는 경우에만 삭제하도록 수정
 *            
 * 수정일자 : 2021-05-10
 * 수정자   : 정은영
 * Reference Key : #0010
 * 수정내용 : (현업요청)
 *            1. 월간검수보고서 레포트 수정(중량 큰순으로 정렬, 합계 추가)
 * 
 */
namespace AccAdm
{
    public partial class PD04001F01 : DevExpress.XtraEditors.XtraForm
    {
        public PD04001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private List<string> _ListWordByGumsu = new List<string>();
        private List<string> _ListWordBySummary = new List<string>();
        private void PD04001F01_Load(object sender, EventArgs e)
        {
            _ListWordByGumsu.Add("업체명");
            _ListWordByGumsu.Add("등급");
            _ListWordByGumsu.Add("검수자");

            _ListWordBySummary.Add("업체명");

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            arrGrdView = new GridView[] { GridViewRetr, GridViewRetr2 };
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

            InitControls();
            UpdateDropDownButton(BtnWeek);
            InitControlsTileView();

            ComnEtcFunc.SetBoundGridLookUp(RepoChrgid, "HR_EMP_BASIS", "EMP_ID", "EMP_NM");

            BtnRetr.PerformClick();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                string sFindIdx = CboFindIdx.SelectedIndex < 0 ? "0" : CboFindIdx.SelectedIndex.ToString();
                string sFindWord = TxtFindWord.EditValue?.ToString().Trim();
                string sSelfYn = ChkSelf.EditValue?.ToString() ?? string.Empty;
                string sItnlYn = RdgbItnlYn.EditValue?.ToString() ?? string.Empty;

                if (string.IsNullOrEmpty(sYmdFrom))
                {
                    XtraMessageBox.Show("계근일자를 올바르게 입력하세요.");
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                    return;
                }
                else if (string.IsNullOrEmpty(sYmdFrom))
                {
                    XtraMessageBox.Show("계근일자를 올바르게 입력하세요.");
                    DateEditTo.SelectAll();
                    DateEditTo.Focus();
                    return;
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);
                dicParams.Add("FIND_IDX", sFindIdx);
                dicParams.Add("FIND_WORD", sFindWord);
                dicParams.Add("SELF_YN", sSelfYn);
                dicParams.Add("ITNL_YN", sItnlYn);

                /* 
                 * 
                 * #0002
                 * 
                 */
                //GridRetr.DataSource = null;
                //GridRetr.DataSource = GetInspectionInfo(dicParams);
                //if (GridViewRetr.RowCount > 0)
                //{
                //    GridViewRetr.Focus();
                //}

                GridControl grdList = new GridControl();
                GridView grdView = new GridView();
                DataTable dt = new DataTable();
                if (TabControl.SelectedTabPage == TabPageGumsu)
                {
                    grdList = GridRetr;
                    grdView = GridViewRetr;
                    dt = GetInspectionInfo(dicParams);
                }
                else
                {
                    grdList = GridRetr2;
                    grdView = GridViewRetr2;
                    dt = GetSummary(dicParams);
                }

                grdList.DataSource = null;
                grdList.DataSource = dt;
                if (grdView.RowCount > 0)
                {
                    grdView.Focus();
                }
                else if (grdView.RowCount == 0)
                {
                    if (string.IsNullOrEmpty(TxtFindWord.EditValue?.ToString()))
                    {
                        TxtFindWord.SelectAll();
                        TxtFindWord.Focus();
                    }
                    else
                    {
                        DateEditFrom.SelectAll();
                        DateEditFrom.Focus();
                    }
                }

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        //검수내역 마스터 조회
        private DataTable GetInspectionInfo(Dictionary<string, string> dicParams)
        {
            /*
             * #0001
             */
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendFormat("\r\n SELECT *  ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT A1.JUNPYOID ");
            //strSql.AppendFormat("\r\n               , C1.ISPT_NO ");
            //strSql.AppendFormat("\r\n               , DATE_FORMAT(C1.ENT_DT,'%Y-%m-%d') AS ENT_DT ");
            //strSql.AppendFormat("\r\n               , C1.ENT_ID AS CUSER ");
            //strSql.AppendFormat("\r\n               , C2.EMP_NM AS CUSER_NM ");
            //strSql.AppendFormat("\r\n               , A1.J_DATE ");
            //strSql.AppendFormat("\r\n               , A1.KERATYPE ");
            //strSql.AppendFormat("\r\n               , D1.DAEGUBUN ");
            //strSql.AppendFormat("\r\n               , A1.GUMSU_SERIAL ");
            //strSql.AppendFormat("\r\n               , B1.EMP_NM AS ISPT_USER ");
            //strSql.AppendFormat("\r\n               , A1.SUN ");
            //strSql.AppendFormat("\r\n               , A1.J_BNUM ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHERID ELSE A1.J_ASSIGNID END AS DEALER_CD ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHER ELSE A1.J_COMPANY END AS DEALER_NM ");
            //strSql.AppendFormat("\r\n               , A1.J_SERIAL ");
            //strSql.AppendFormat("\r\n               , A1.GUBUN1 ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.SECONDTIME ELSE A1.FIRSTTIME END AS FIRSTTIME ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.SECONDWEIGHT ELSE A1.FIRSTWEIGHT END AS FIRSTWEIGHT ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.FIRSTTIME ELSE A1.SECONDTIME END AS SECONDTIME ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.FIRSTWEIGHT ELSE A1.SECONDWEIGHT END AS SECONDWEIGHT ");
            //strSql.AppendFormat("\r\n               , IFNULL(C1.CHAGAM, IFNULL(A2.CHAGAM, CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END)) AS FIRST_CHAGAM ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END AS CHAGAM_ADMT ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT ");
            //strSql.AppendFormat("\r\n               , A1.J_STATE ");
            //strSql.AppendFormat("\r\n               , C1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , A1.GUMSUBIGO ");
            //strSql.AppendFormat("\r\n               , C1.ISPT_OPN ");
            //strSql.AppendFormat("\r\n               , C1.IMG_CNT ");
            //strSql.AppendFormat("\r\n            FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN MESURING_SEQ A2 ");
            //strSql.AppendFormat("\r\n              ON A1.JUNPYOID = ( SELECT X1.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE A2.JUNPYOID = X1.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                    AND X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  ORDER BY X1.JUNPYOID, X1.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n            LEFT JOIN HR_EMP_BASIS B1 ");
            //strSql.AppendFormat("\r\n              ON A1.GUMSU_SERIAL = B1.EMP_ID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN MESURE_ISPT_INFO C1 ");
            //strSql.AppendFormat("\r\n              ON A1.JUNPYOID = C1.JUNPYOID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.USRCD ");
            //strSql.AppendFormat("\r\n                             , X2.EMP_NM ");
            //strSql.AppendFormat("\r\n                          FROM ZUSRLST X1 ");
            //strSql.AppendFormat("\r\n                          LEFT JOIN HR_EMP_BASIS X2 ");
            //strSql.AppendFormat("\r\n                            ON X1.INSANO = X2.EMP_ID ) C2 ");
            //strSql.AppendFormat("\r\n              ON C1.ENT_ID = C2.USRCD ");
            //strSql.AppendFormat("\r\n            LEFT JOIN JAJAE D1 ");
            //strSql.AppendFormat("\r\n              ON A1.J_SERIAL = D1.J_SERIAL ");
            //strSql.AppendFormat("\r\n           WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n             AND D1.DAEGUBUN IN ('고철A', '고철B', '슈레더')");
            //strSql.AppendFormat("\r\n             AND A1.KERATYPE = '입고'  ");
            //strSql.AppendFormat("\r\n             AND A1.J_SERIAL > 0 ");
            //strSql.AppendFormat("\r\n        ) Y1 ");
            //strSql.AppendFormat("\r\n  WHERE Y1.FIRST_CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n    AND (('{0}' = '' AND 1 = 1 ) ", dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = '0' AND Y1.DEALER_NM LIKE '%{1}%')", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = '1' AND Y1.GUBUN1 LIKE '%{1}%')", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = '2' AND Y1.ISPT_USER LIKE '%{1}%') )", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);

            //if (!FmMainToolBar2.EmpID.Equals("****")){
            //    strSql.AppendFormat("\r\n    AND (('{0}' <> 'Y' AND 1 = 1) ", dicParams["SELF_YN"]);
            //    strSql.AppendFormat("\r\n         OR");
            //    strSql.AppendFormat("\r\n         ('{0}' = 'Y' AND Y1.GUMSU_SERIAL = {1}) )", dicParams["SELF_YN"], FmMainToolBar2.EmpID);
            //    strSql.AppendFormat("\r\n  ORDER BY Y1.J_DATE, Y1.SUN ");
            //}
            #endregion

            strSql.AppendLine("SELECT *");
            strSql.AppendLine(" FROM(");
            strSql.AppendLine("        SELECT A1.JUNPYOID");
            strSql.AppendLine("             , C1.ISPT_NO");
            strSql.AppendLine("             , CONVERT(VARCHAR(20), C1.ENT_DT, 23) AS ENT_DT");
            strSql.AppendLine("             , C1.ENT_ID AS CUSER");
            strSql.AppendLine("             , C2.EMP_NM AS CUSER_NM");
            strSql.AppendLine("             , A1.J_DATE");
            strSql.AppendLine("             , A1.KERATYPE");
            strSql.AppendLine("             , D1.DAEGUBUN");
            strSql.AppendLine("             , A1.GUMSU_SERIAL");
            strSql.AppendLine("             , B1.EMP_NM AS ISPT_USER");
            strSql.AppendLine("             , A1.SUN");
            strSql.AppendLine("             , A1.J_BNUM");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHERID ELSE A1.J_ASSIGNID END AS DEALER_CD");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHER ELSE A1.J_COMPANY END AS DEALER_NM");
            strSql.AppendLine("             , A1.J_SERIAL");
            strSql.AppendLine("             , A1.GUBUN1");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN A1.SECONDTIME ELSE A1.FIRSTTIME END AS FIRSTTIME");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN A1.SECONDWEIGHT ELSE A1.FIRSTWEIGHT END AS FIRSTWEIGHT");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN A1.FIRSTTIME ELSE A1.SECONDTIME END AS SECONDTIME");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN A1.FIRSTWEIGHT ELSE A1.SECONDWEIGHT END AS SECONDWEIGHT");
            strSql.AppendLine("             , ISNULL(C1.CHAGAM, ISNULL(A2.CHAGAM, CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END)) AS FIRST_CHAGAM");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END AS CHAGAM_ADMT");
            strSql.AppendLine("             , CASE WHEN A1.KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT");
            strSql.AppendLine("             , A1.J_STATE");
            strSql.AppendLine("             , C1.ITNL_YN");
            strSql.AppendLine("             , A1.GUMSUBIGO");
            strSql.AppendLine("             , C1.ISPT_OPN");
            strSql.AppendLine("             , C1.IMG_CNT");
            strSql.AppendLine("          FROM MESURING A1");
            strSql.AppendLine("         LEFT JOIN MESURING_SEQ A2");
            strSql.AppendLine("            ON A1.JUNPYOID = (SELECT TOP 1 X1.JUNPYOID");
            strSql.AppendLine("                                 FROM MESURING_SEQ X1");
            strSql.AppendLine("                                WHERE A2.JUNPYOID = X1.JUNPYOID");
            strSql.AppendLine("                                  AND X1.CHAGAM > 0");
            strSql.AppendLine("                                ORDER BY X1.JUNPYOID, X1.CHG_SEQ");
            strSql.AppendLine("                                )");
            strSql.AppendLine("          LEFT JOIN HR_EMP_BASIS B1");
            strSql.AppendLine("            ON A1.GUMSU_SERIAL = B1.EMP_ID");
            strSql.AppendLine("          LEFT JOIN MESURE_ISPT_INFO C1");
            strSql.AppendLine("            ON A1.JUNPYOID = C1.JUNPYOID");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.USRCD");
            strSql.AppendLine("                           , X2.EMP_NM");
            strSql.AppendLine("                        FROM ZUSRLST X1");
            strSql.AppendLine("                        LEFT JOIN HR_EMP_BASIS X2");
            strSql.AppendLine("                          ON X1.INSANO = X2.EMP_ID) C2");
            strSql.AppendLine("           ON C1.ENT_ID = C2.USRCD");
            strSql.AppendLine("          LEFT JOIN JAJAE D1");
            strSql.AppendLine("            ON A1.J_SERIAL = D1.J_SERIAL");
            strSql.AppendFormat("\r\n           WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n             AND D1.DAEGUBUN IN ('고철A', '고철B', '슈레더')");
            strSql.AppendFormat("\r\n             AND A1.KERATYPE = '입고'  ");
            strSql.AppendFormat("\r\n             AND A1.J_SERIAL > 0 ");
            strSql.AppendFormat("\r\n        ) Y1 ");
            strSql.AppendFormat("\r\n  WHERE Y1.FIRST_CHAGAM > 0 ");
            strSql.AppendFormat("\r\n    AND (('{0}' = '' AND 1 = 1 ) ", dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR");
            strSql.AppendFormat("\r\n         ('{0}' = '0' AND Y1.DEALER_NM LIKE '%{1}%')", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR");
            strSql.AppendFormat("\r\n         ('{0}' = '1' AND Y1.GUBUN1 LIKE '%{1}%')", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR");
            strSql.AppendFormat("\r\n         ('{0}' = '2' AND Y1.ISPT_USER LIKE '%{1}%') )", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);

            if (!FmMainToolBar2.EmpID.Equals("****"))
            {
                strSql.AppendFormat("\r\n    AND (('{0}' <> 'Y' AND 1 = 1) ", dicParams["SELF_YN"]);
                strSql.AppendFormat("\r\n         OR");
                strSql.AppendFormat("\r\n         ('{0}' = 'Y' AND Y1.GUMSU_SERIAL = {1}) )", dicParams["SELF_YN"], FmMainToolBar2.EmpID);
                strSql.AppendFormat("\r\n  ORDER BY Y1.J_DATE, Y1.SUN ");
            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            #region[2021-04-06 이전 쿼리]

            //strSql.AppendFormat("\r\n SELECT *  ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT A1.JUNPYOID ");
            //strSql.AppendFormat("\r\n               , C1.ISPT_NO ");
            //strSql.AppendFormat("\r\n               , DATE_FORMAT(C1.ENT_DT,'%Y-%m-%d') AS ENT_DT ");
            //strSql.AppendFormat("\r\n               , C1.ENT_ID AS CUSER ");
            //strSql.AppendFormat("\r\n               , C2.EMP_NM AS CUSER_NM ");
            //strSql.AppendFormat("\r\n               , A1.J_DATE ");
            //strSql.AppendFormat("\r\n               , A1.KERATYPE ");
            //strSql.AppendFormat("\r\n               , A1.GUMSU_SERIAL ");
            //strSql.AppendFormat("\r\n               , B1.EMP_NM AS ISPT_USER ");
            //strSql.AppendFormat("\r\n               , A1.SUN ");
            //strSql.AppendFormat("\r\n               , A1.J_BNUM ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHERID ELSE A1.J_ASSIGNID END AS DEALER_CD ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHER ELSE A1.J_COMPANY END AS DEALER_NM ");
            //strSql.AppendFormat("\r\n               , A1.J_SERIAL ");
            //strSql.AppendFormat("\r\n               , A1.GUBUN1 ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.SECONDTIME ELSE A1.FIRSTTIME END AS FIRSTTIME ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.SECONDWEIGHT ELSE A1.FIRSTWEIGHT END AS FIRSTWEIGHT ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.FIRSTTIME ELSE A1.SECONDTIME END AS SECONDTIME ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN A1.FIRSTWEIGHT ELSE A1.SECONDWEIGHT END AS SECONDWEIGHT ");
            ////strSql.AppendFormat("\r\n               , A1.FIRSTTIME ");
            ////strSql.AppendFormat("\r\n               , A1.FIRSTWEIGHT ");
            ////strSql.AppendFormat("\r\n               , A1.SECONDTIME ");
            ////strSql.AppendFormat("\r\n               , A1.SECONDWEIGHT ");
            //strSql.AppendFormat("\r\n               , IFNULL(C1.CHAGAM, IFNULL(A2.CHAGAM, CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END)) AS FIRST_CHAGAM ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END AS CHAGAM_ADMT ");
            //strSql.AppendFormat("\r\n               , CASE WHEN A1.KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT ");
            //strSql.AppendFormat("\r\n               , A1.J_STATE ");
            //strSql.AppendFormat("\r\n               , C1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , A1.GUMSUBIGO ");
            //strSql.AppendFormat("\r\n               , C1.ISPT_OPN ");
            //strSql.AppendFormat("\r\n               , C1.IMG_CNT ");
            //strSql.AppendFormat("\r\n            FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.JUNPYOID ");
            //strSql.AppendFormat("\r\n                             , X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                          FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                         LIMIT 1 ) A2 ");
            //strSql.AppendFormat("\r\n              ON A1.JUNPYOID = A2.JUNPYOID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN HR_EMP_BASIS B1 ");
            //strSql.AppendFormat("\r\n              ON A1.GUMSU_SERIAL = B1.EMP_ID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN MESURE_ISPT_INFO C1 ");
            //strSql.AppendFormat("\r\n              ON A1.JUNPYOID = C1.JUNPYOID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.USRCD ");
            //strSql.AppendFormat("\r\n                             , X2.EMP_NM ");
            //strSql.AppendFormat("\r\n                          FROM ZUSRLST X1 ");
            //strSql.AppendFormat("\r\n                          LEFT JOIN HR_EMP_BASIS X2 ");
            //strSql.AppendFormat("\r\n                            ON X1.INSANO = X2.EMP_ID ) C2 ");
            //strSql.AppendFormat("\r\n              ON C1.ENT_ID = C2.USRCD ");
            //strSql.AppendFormat("\r\n            LEFT JOIN JAJAE D1 ");
            //strSql.AppendFormat("\r\n              ON A1.J_SERIAL = D1.J_SERIAL ");
            //strSql.AppendFormat("\r\n           WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n             AND D1.DAEGUBUN IN ('고철A', '고철B', '슈레더')");
            //strSql.AppendFormat("\r\n             AND A1.KERATYPE <> '직송'  ");
            //strSql.AppendFormat("\r\n             AND A1.J_SERIAL > 0 ");
            //strSql.AppendFormat("\r\n        ) Y1 ");
            //strSql.AppendFormat("\r\n  WHERE Y1.FIRST_CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n    AND (('{0}' = '' AND 1 = 1 ) ", dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = '0' AND Y1.DEALER_NM LIKE '%{1}%')", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = '1' AND Y1.GUBUN1 LIKE '%{1}%')", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = '2' AND Y1.ISPT_USER LIKE '%{1}%') )", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n    AND (('{0}' <> 'Y' AND 1 = 1) ", dicParams["SELF_YN"]);
            //strSql.AppendFormat("\r\n         OR");
            //strSql.AppendFormat("\r\n         ('{0}' = 'Y' AND Y1.GUMSU_SERIAL = {1}) )", dicParams["SELF_YN"], FmMainToolBar2.EmpID);
            //strSql.AppendFormat("\r\n  ORDER BY Y1.J_DATE, Y1.SUN ");

            //return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            #endregion[2021-04-06 이전 쿼리]
        }

        /*
         * #0002
         * 
         * #0006
         * 
         * #0007
         * 
         * #0008
         */
        private DataTable GetSummary(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            #region[이전쿼리 ~23.06.28]
            strSql.Clear();
            strSql.AppendLine("DECLARE @FDATE VARCHAR(10) = '" + dicParams["DATE_F"] + "' , @TDATE VARCHAR(10) = '" + dicParams["DATE_T"] + "';");
            strSql.AppendLine("WITH TEMP1 AS(                                                                                               ");
            strSql.AppendLine("    SELECT A1.IPCHULGO_MAIPID                                                                                ");
            strSql.AppendLine("         , CASE WHEN A22.JUNPYOID IS NULL THEN A23.EMP_NM ELSE A22.CHRG_NM END AS CHRG_NM                 ");
            strSql.AppendLine("      FROM MESURING A1                                                                                     ");
            strSql.AppendLine("      LEFT JOIN ACC_DEALER_CD A2                                                                           ");
            strSql.AppendLine("        ON A1.MAIPCHERID = A2.DEALER_CD                                                                      ");
            strSql.AppendLine("      LEFT JOIN INLIST A22                                                                                 ");
            strSql.AppendLine("        ON A1.IPCHULGO_MAIPID = A22.J_ID                                                                     ");
            strSql.AppendLine("      LEFT JOIN HR_EMP_BASIS A23                                                                           ");
            strSql.AppendLine("        ON A2.CHRG_ID = A23.EMP_ID                                                                           ");
            strSql.AppendLine("      LEFT JOIN JAJAE B1                                                                                   ");
            strSql.AppendLine("        ON A1.J_SERIAL = B1.J_SERIAL                                                                         ");
            strSql.AppendLine("     WHERE A1.J_DATE BETWEEN @FDATE AND @TDATE                                                    ");
            strSql.AppendLine("       AND A1.KERATYPE = '입고'                                                                             ");
            strSql.AppendLine("       AND B1.DAEGUBUN IN('고철A', '고철B')                                                                 ");
            strSql.AppendLine("),TEMP2 AS(                                                                                                  ");
            strSql.AppendLine("    SELECT '1' AS OPN_GB                                                                                     ");
            strSql.AppendLine("         , '스크랩' AS GB                                                                                    ");
            strSql.AppendLine("         , A1.MAIPCHERID                                                                                     ");
            strSql.AppendLine("         , A2.DEALER_NM                                                                                      ");
            strSql.AppendLine("         , A6.CHRG_NM                                                                                        ");
            strSql.AppendLine("         , SUM(A1.IWEIGHT) AS WGT                                                                            ");
            strSql.AppendLine("         , ISNULL(A3.CHAGAM, ISNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM                                           ");
            strSql.AppendLine("         , COUNT(*) AS IN_CNT                                                                                ");
            strSql.AppendLine("         , ISNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT                                                            ");
            strSql.AppendLine("         , SUM(A1.ICHAGAM) AS CHAGAM_WGT                                                                     ");
            strSql.AppendLine("         , CASE WHEN ISNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN                           ");
            strSql.AppendLine("      FROM MESURING A1                                                                                       ");
            strSql.AppendLine("      LEFT JOIN ACC_DEALER_CD A2                                                                             ");
            strSql.AppendLine("        ON A1.MAIPCHERID = A2.DEALER_CD                                                                      ");
            strSql.AppendLine("      LEFT JOIN(SELECT X1.DEALER_CD                                                                          ");
            strSql.AppendLine("                     , SUM(X1.CHAGAM) AS CHAGAM                                                              ");
            strSql.AppendLine("                     , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT                              ");
            strSql.AppendLine("                  FROM MESURE_ISPT_INFO X1                                                                   ");
            strSql.AppendLine("                  LEFT JOIN MESURING X2 ON X1.JUNPYOID = X2.JUNPYOID                                      ");
            strSql.AppendLine("                  WHERE X2.J_DATE BETWEEN @FDATE AND @TDATE                                     ");
            strSql.AppendLine("                 GROUP BY X1.DEALER_CD ) A3                                                                  ");
            strSql.AppendLine("       ON A1.MAIPCHERID = A3.DEALER_CD                                                                       ");
            strSql.AppendLine("      AND A1.J_DATE BETWEEN @FDATE AND @TDATE                                                                ");
            strSql.AppendLine("     LEFT JOIN(SELECT X1.MAIPCHERID                                                                          ");
            strSql.AppendLine("                    , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1                                                   ");
            strSql.AppendLine("                                 WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT                               ");
            strSql.AppendLine("                 FROM MESURING X1                                                                            ");
            strSql.AppendLine("                 LEFT JOIN MESURING_SEQ X2                                                                   ");
            strSql.AppendLine("                   ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                                 ");
            strSql.AppendLine("                                       FROM MESURING_SEQ W                                                   ");
            strSql.AppendLine("                                      WHERE X2.JUNPYOID = W.JUNPYOID                                         ");
            strSql.AppendLine("                                        AND W.CHAGAM > 0                                                     ");
            strSql.AppendLine("                                      ORDER BY W.JUNPYOID, W.CHG_SEQ )                                       ");
            strSql.AppendLine("                  LEFT JOIN JAJAE X3                                                                         ");
            strSql.AppendLine("                    ON X1.J_SERIAL = X3.J_SERIAL                                                             ");
            strSql.AppendLine("                 WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                                   ");
            strSql.AppendLine("                   AND X3.DAEGUBUN IN('고철A', '고철B')                                                      ");
            strSql.AppendLine("                 GROUP BY X1.MAIPCHERID ) A4                                                                 ");
            strSql.AppendLine("       ON A1.MAIPCHERID = A4.MAIPCHERID                                                                      ");
            strSql.AppendLine("     LEFT JOIN(SELECT TOP 1 X1.JUNPYOID, X1.CHAGAM                                                           ");
            strSql.AppendLine("                 FROM MESURING_SEQ X1                                                                        ");
            strSql.AppendLine("                WHERE X1.CHAGAM > 0 ) A5                                                                     ");
            strSql.AppendLine("       ON A1.JUNPYOID = A5.JUNPYOID                                                                          ");
            strSql.AppendLine("     LEFT JOIN TEMP1 A6                                                                                      ");
            strSql.AppendLine("       ON A1.IPCHULGO_MAIPID = A6.IPCHULGO_MAIPID                                                            ");
            strSql.AppendLine("    WHERE A1.ipchulgo_maipid = A6.IPCHULGO_MAIPID                                                            ");
            strSql.AppendLine("    GROUP BY A1.MAIPCHERID, A2.DEALER_NM, A3.CHAGAM, A4.CHAGAM_CNT, A3.ITNL_CNT, A6.CHRG_NM                  ");
            strSql.AppendLine("    ), TEMP3 AS(                                                                                             ");
            strSql.AppendLine("        SELECT Y1.OPN_GB                                                                                     ");
            strSql.AppendLine("             , Y1.GB                                                                                         ");
            strSql.AppendLine("             , Y1.MAIPCHERID                                                                                 ");
            strSql.AppendLine("             , Y1.DEALER_NM                                                                                  ");
            strSql.AppendLine("             , Y1.CHRG_NM                                                                                    ");
            strSql.AppendLine("             , Y1.WGT                                                                                        ");
            strSql.AppendLine("             , Y2.SUM_CHAGAM                                                                                 ");
            strSql.AppendLine("             , Y1.IN_CNT                                                                                     ");
            strSql.AppendLine("             , Y1.CHAGAM_CNT                                                                                 ");
            strSql.AppendLine("             , Y1.CHAGAM_WGT                                                                                 ");
            strSql.AppendLine("             , ISNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT                                                        ");
            strSql.AppendLine("             , ISNULL(Y1.CHAGAM_WGT, 0)  -ISNULL(Y2.SUM_CHAGAM, 0) AS ADMT_WGT                                ");
            strSql.AppendLine("             , Y1.ITNL_YN                                                                                    ");
            strSql.AppendLine("             , Y3.OPN_DATE                                                                                   ");
            strSql.AppendLine("             , Y3.OPN_RMK                                                                                    ");
            strSql.AppendLine("             , @TDATE AS DATE_T                                                                              ");
            strSql.AppendLine("          FROM TEMP2 Y1                                                                                      ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.MAIPCHERID                                                                     ");
            strSql.AppendLine("                         , SUM(ISNULL(X3.CHAGAM, ISNULL(X2.CHAGAM, ISNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM    ");
            strSql.AppendLine("                      FROM MESURING X1                                                                       ");
            strSql.AppendLine("                      LEFT JOIN MESURING_SEQ X2                                                              ");
            strSql.AppendLine("                        ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                            ");
            strSql.AppendLine("                                            FROM MESURING_SEQ W                                              ");
            strSql.AppendLine("                                            WHERE X2.JUNPYOID = W.JUNPYOID                                   ");
            strSql.AppendLine("                                              AND W.CHAGAM > 0                                               ");
            strSql.AppendLine("                                            ORDER BY W.JUNPYOID, W.CHG_SEQ)                                  ");
            strSql.AppendLine("                        LEFT JOIN MESURE_ISPT_INFO X3                                                        ");
            strSql.AppendLine("                          ON X1.JUNPYOID = X3.JUNPYOID                                                       ");
            strSql.AppendLine("                        LEFT JOIN JAJAE X4                                                                   ");
            strSql.AppendLine("                          ON X1.J_SERIAL = X4.J_SERIAL                                                       ");
            strSql.AppendLine("                       WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                 ");
            strSql.AppendLine("                         AND X1.KERATYPE = '입고'                                                            ");
            strSql.AppendLine("                         AND X4.DAEGUBUN IN('고철A', '고철B')                                                ");
            strSql.AppendLine("                       GROUP BY X1.MAIPCHERID ) Y2                                                           ");
            strSql.AppendLine("            ON Y1.MAIPCHERID = Y2.MAIPCHERID                                                                 ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.DEALER_CD                                                                      ");
            strSql.AppendLine("                         , X1.OPN_GB                                                                         ");
            strSql.AppendLine("                         , X1.OPN_DATE AS OPN_DATE                                                           ");
            strSql.AppendLine("                         , X1.OPN_RMK AS OPN_RMK                                                             ");
            strSql.AppendLine("                      FROM MESURE_OPN_HISTORY X1                                                             ");
            strSql.AppendLine("                     WHERE X1.OPN_DATE = (SELECT MAX(W.OPN_DATE) AS OPN_DATE                                 ");
            strSql.AppendLine("                                            FROM MESURE_OPN_HISTORY W                                        ");
            strSql.AppendLine("                                           WHERE REPLACE(W.OPN_RMK, ' ', '') <> ''                           ");
            strSql.AppendLine("                                             AND W.OPN_DATE BETWEEN @FDATE AND @TDATE            ");
            strSql.AppendLine("                                             AND X1.DEALER_CD = W.DEALER_CD                                  ");
            strSql.AppendLine("                                             AND X1.OPN_GB = W.OPN_GB                                        ");
            strSql.AppendLine("                                           GROUP BY W.DEALER_CD, W.OPN_GB )                                  ");
            strSql.AppendLine("                        GROUP BY X1.DEALER_CD, X1.OPN_GB, X1.OPN_DATE, X1.OPN_RMK                            ");
            strSql.AppendLine("                  ) Y3                                                                                       ");
            strSql.AppendLine("            ON Y1.MAIPCHERID = Y3.DEALER_CD                                                                  ");
            strSql.AppendLine("           AND Y1.OPN_GB = Y3.OPN_GB                                                                         ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.MAIPCHERID                                                                     ");
            strSql.AppendLine("                         , COUNT(*) AS RETURN_CNT                                                            ");
            strSql.AppendLine("                      FROM MESURING X1                                                                       ");
            strSql.AppendLine("                     WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                  ");
            strSql.AppendLine("                       AND X1.KERATYPE = '입고'                                                              ");
            strSql.AppendLine("                       AND X1.J_SERIAL = 4049042--스크랩반품                                                 ");
            strSql.AppendLine("                     GROUP BY X1.MAIPCHERID ) Y4                                                             ");
            strSql.AppendLine("           ON Y1.MAIPCHERID = Y4.MAIPCHERID                                                                  ");
            strSql.AppendLine("        WHERE Y1.CHAGAM_CNT > 0                                                                              ");
            strSql.AppendLine("    ), TEMP4 AS(                                                                                             ");
            strSql.AppendLine("        SELECT A1.IPCHULGO_MAIPID                                                                            ");
            strSql.AppendLine("             , CASE WHEN A22.JUNPYOID IS NULL THEN A23.EMP_NM ELSE A22.CHRG_NM END AS CHRG_NM                ");
            strSql.AppendLine("          FROM MESURING A1                                                                                   ");
            strSql.AppendLine("          LEFT JOIN ACC_DEALER_CD A2                                                                         ");
            strSql.AppendLine("            ON A1.MAIPCHERID = A2.DEALER_CD                                                                  ");
            strSql.AppendLine("          LEFT JOIN INLIST A22                                                                               ");
            strSql.AppendLine("            ON A1.IPCHULGO_MAIPID = A22.J_ID                                                                 ");
            strSql.AppendLine("          LEFT JOIN HR_EMP_BASIS A23                                                                         ");
            strSql.AppendLine("            ON A2.CHRG_ID = A23.EMP_ID                                                                       ");
            strSql.AppendLine("          LEFT JOIN JAJAE B1                                                                                 ");
            strSql.AppendLine("            ON A1.J_SERIAL = B1.J_SERIAL                                                                     ");
            strSql.AppendLine("         WHERE A1.J_DATE BETWEEN @FDATE AND @TDATE                                               ");
            strSql.AppendLine("           AND A1.KERATYPE = '입고'                                                                          ");
            strSql.AppendLine("           AND B1.DAEGUBUN IN('슈레더')                                                                      ");
            strSql.AppendLine("    ), TEMP5 AS(                                                                                             ");
            strSql.AppendLine("        SELECT '2' AS OPN_GB                                                                                 ");
            strSql.AppendLine("             , '슈레더' AS GB                                                                                ");
            strSql.AppendLine("             , A1.MAIPCHERID                                                                                 ");
            strSql.AppendLine("             , A2.DEALER_NM                                                                                  ");
            strSql.AppendLine("             , SUM(A1.IWEIGHT) AS WGT                                                                        ");
            strSql.AppendLine("             , ISNULL(A3.CHAGAM, ISNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM                                       ");
            strSql.AppendLine("             , COUNT(*) AS IN_CNT                                                                            ");
            strSql.AppendLine("             , ISNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT                                                        ");
            strSql.AppendLine("             , SUM(A1.ICHAGAM) AS CHAGAM_WGT                                                                 ");
            strSql.AppendLine("             , A6.CHRG_NM                                                                                    ");
            strSql.AppendLine("             , CASE WHEN ISNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN                       ");
            strSql.AppendLine("          FROM MESURING A1                                                                                   ");
            strSql.AppendLine("          LEFT JOIN ACC_DEALER_CD A2                                                                         ");
            strSql.AppendLine("            ON A1.MAIPCHERID = A2.DEALER_CD                                                                  ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.DEALER_CD                                                                      ");
            strSql.AppendLine("                         , SUM(X1.CHAGAM) AS CHAGAM                                                          ");
            strSql.AppendLine("                         , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT                          ");
            strSql.AppendLine("                      FROM MESURE_ISPT_INFO X1                                                               ");
            strSql.AppendLine("                  LEFT JOIN MESURING X2 ON X1.JUNPYOID = X2.JUNPYOID                                      ");
            strSql.AppendLine("                  WHERE X2.J_DATE BETWEEN @FDATE AND @TDATE                                     ");
            strSql.AppendLine("                     GROUP BY X1.DEALER_CD ) A3                                                              ");
            strSql.AppendLine("            ON A1.MAIPCHERID = A3.DEALER_CD                                                                  ");
            strSql.AppendLine("           AND A1.J_DATE BETWEEN @FDATE AND @TDATE                                               ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.MAIPCHERID                                                                     ");
            strSql.AppendLine("                         , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1                                              ");
            strSql.AppendLine("                                      WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT                          ");
            strSql.AppendLine("                      FROM MESURING X1                                                                       ");
            strSql.AppendLine("                      LEFT JOIN MESURING_SEQ X2                                                              ");
            strSql.AppendLine("                        ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                            ");
            strSql.AppendLine("                                            FROM MESURING_SEQ W                                              ");
            strSql.AppendLine("                                           WHERE X2.JUNPYOID = W.JUNPYOID                                    ");
            strSql.AppendLine("                                             AND W.CHAGAM > 0                                                ");
            strSql.AppendLine("                                           ORDER BY W.JUNPYOID, W.CHG_SEQ )                                  ");
            strSql.AppendLine("                      LEFT JOIN JAJAE X3                                                                     ");
            strSql.AppendLine("                        ON X1.J_SERIAL = X3.J_SERIAL                                                         ");
            strSql.AppendLine("                     WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                   ");
            strSql.AppendLine("                       AND X3.DAEGUBUN = '슈레더'                                                            ");
            strSql.AppendLine("                     GROUP BY X1.MAIPCHERID ) A4                                                             ");
            strSql.AppendLine("            ON A1.MAIPCHERID = A4.MAIPCHERID                                                                 ");
            strSql.AppendLine("          LEFT JOIN(SELECT TOP 1 X1.JUNPYOID, X1.CHAGAM                                                      ");
            strSql.AppendLine("                      FROM MESURING_SEQ X1                                                                   ");
            strSql.AppendLine("                     WHERE X1.CHAGAM > 0 ) A5                                                                ");
            strSql.AppendLine("           ON A1.JUNPYOID = A5.JUNPYOID                                                                      ");
            strSql.AppendLine("         LEFT JOIN TEMP4 A6                                                                                  ");
            strSql.AppendLine("           ON A1.IPCHULGO_MAIPID = A6.IPCHULGO_MAIPID                                                        ");
            strSql.AppendLine("        WHERE A1.IPCHULGO_MAIPID = A6.IPCHULGO_MAIPID                                                        ");
            strSql.AppendLine("        GROUP BY A1.MAIPCHERID, A2.DEALER_NM, A3.CHAGAM, A4.CHAGAM_CNT, A6.CHRG_NM, A3.ITNL_CNT              ");
            strSql.AppendLine("    ), TEMP6 AS(                                                                                             ");
            strSql.AppendLine("    SELECT Y1.OPN_GB                                                                                         ");
            strSql.AppendLine("         , Y1.GB                                                                                             ");
            strSql.AppendLine("         , Y1.MAIPCHERID                                                                                     ");
            strSql.AppendLine("         , Y1.DEALER_NM                                                                                      ");
            strSql.AppendLine("         , Y1.CHRG_NM                                                                                        ");
            strSql.AppendLine("         , Y1.WGT                                                                                            ");
            strSql.AppendLine("         , Y2.SUM_CHAGAM                                                                                     ");
            strSql.AppendLine("         , Y1.IN_CNT                                                                                         ");
            strSql.AppendLine("         , Y1.CHAGAM_CNT                                                                                     ");
            strSql.AppendLine("         , Y1.CHAGAM_WGT                                                                                     ");
            strSql.AppendLine("         , ISNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT                                                            ");
            strSql.AppendLine("         , ISNULL(Y1.CHAGAM_WGT, 0)  -ISNULL(Y2.SUM_CHAGAM, 0) AS ADMT_WGT                                     ");
            strSql.AppendLine("         , Y1.ITNL_YN                                                                                        ");
            strSql.AppendLine("         , Y3.OPN_DATE                                                                                       ");
            strSql.AppendLine("         , Y3.OPN_RMK                                                                                        ");
            strSql.AppendLine("         , @TDATE AS DATE_T                                                                            ");
            strSql.AppendLine("      FROM TEMP5 Y1                                                                                          ");
            strSql.AppendLine("      LEFT JOIN(SELECT X1.MAIPCHERID                                                                         ");
            strSql.AppendLine("                     , SUM(ISNULL(X3.CHAGAM, ISNULL(X2.CHAGAM, ISNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM        ");
            strSql.AppendLine("                  FROM MESURING X1                                                                           ");
            strSql.AppendLine("                  LEFT JOIN MESURING_SEQ X2                                                                  ");
            strSql.AppendLine("                    ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                                ");
            strSql.AppendLine("                                        FROM MESURING_SEQ W                                                  ");
            strSql.AppendLine("                                       WHERE X2.JUNPYOID = W.JUNPYOID                                        ");
            strSql.AppendLine("                                         AND W.CHAGAM > 0                                                    ");
            strSql.AppendLine("                                       ORDER BY W.JUNPYOID, W.CHG_SEQ)                                       ");
            strSql.AppendLine("                  LEFT JOIN MESURE_ISPT_INFO X3                                                              ");
            strSql.AppendLine("                    ON X1.JUNPYOID = X3.JUNPYOID                                                             ");
            strSql.AppendLine("                  LEFT JOIN JAJAE X4                                                                         ");
            strSql.AppendLine("                    ON X1.J_SERIAL = X4.J_SERIAL                                                             ");
            strSql.AppendLine("                 WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                       ");
            strSql.AppendLine("                   AND X1.KERATYPE = '입고'                                                                  ");
            strSql.AppendLine("                   AND X4.DAEGUBUN IN('슈레더')                                                              ");
            strSql.AppendLine("                 GROUP BY X1.MAIPCHERID ) Y2                                                                 ");
            strSql.AppendLine("        ON Y1.MAIPCHERID = Y2.MAIPCHERID                                                                     ");
            strSql.AppendLine("      LEFT JOIN(SELECT X1.DEALER_CD                                                                          ");
            strSql.AppendLine("                     , X1.OPN_GB                                                                             ");
            strSql.AppendLine("                     , X1.OPN_DATE AS OPN_DATE                                                               ");
            strSql.AppendLine("                     , X1.OPN_RMK AS OPN_RMK                                                                 ");
            strSql.AppendLine("                  FROM MESURE_OPN_HISTORY X1                                                                 ");
            strSql.AppendLine("                 WHERE X1.OPN_DATE = (SELECT MAX(W.OPN_DATE) AS OPN_DATE                                     ");
            strSql.AppendLine("                                        FROM MESURE_OPN_HISTORY W                                            ");
            strSql.AppendLine("                                       WHERE REPLACE(W.OPN_RMK, ' ', '') <> ''                               ");
            strSql.AppendLine("                                         AND W.OPN_DATE BETWEEN @FDATE AND @TDATE                ");
            strSql.AppendLine("                                         AND X1.DEALER_CD = W.DEALER_CD                                      ");
            strSql.AppendLine("                                         AND X1.OPN_GB = W.OPN_GB                                            ");
            strSql.AppendLine("                                       GROUP BY W.DEALER_CD, W.OPN_GB )                                      ");
            strSql.AppendLine("                 GROUP BY X1.DEALER_CD, X1.OPN_GB, X1.OPN_DATE, X1.OPN_RMK ) Y3                              ");
            strSql.AppendLine("        ON Y1.MAIPCHERID = Y3.DEALER_CD                                                                      ");
            strSql.AppendLine("       AND Y1.OPN_GB = Y3.OPN_GB                                                                             ");
            strSql.AppendLine("      LEFT JOIN(SELECT X1.MAIPCHERID                                                                         ");
            strSql.AppendLine("                     , COUNT(*) AS RETURN_CNT                                                                ");
            strSql.AppendLine("                  FROM MESURING X1                                                                           ");
            strSql.AppendLine("                 WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                       ");
            strSql.AppendLine("                   AND X1.KERATYPE = '입고'                                                                  ");
            strSql.AppendLine("                   AND X1.J_SERIAL = 5050042--슈레더반품                                                     ");
            strSql.AppendLine("                 GROUP BY X1.MAIPCHERID ) Y4                                                                 ");
            strSql.AppendLine("        ON Y1.MAIPCHERID = Y4.MAIPCHERID                                                                     ");
            strSql.AppendLine("      WHERE Y1.CHAGAM_CNT > 0                                                                                ");
            strSql.AppendLine("    )                                                                                                        ");
            strSql.AppendLine("                                                                                                             ");
            strSql.AppendLine("SELECT *                                                                                                     ");
            strSql.AppendLine("  FROM (                                                                                                     ");
            strSql.AppendLine("        SELECT * FROM TEMP3                                                                                  ");
            strSql.AppendLine("         UNION ALL                                                                                           ");
            strSql.AppendLine("        SELECT * FROM TEMP6 ) Z1                                                                             ");
            strSql.AppendLine("       WHERE (( '" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 )");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ( '" + dicParams["FIND_IDX"] + "' = '0' AND Z1.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))");
            strSql.AppendLine("         AND (( '" + dicParams["ITNL_YN"] + "' = 'ALL' AND 1 = 1 ) ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ( '" + dicParams["ITNL_YN"] + "' = 'Y' AND Z1.ITNL_YN = '유' )");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ( '" + dicParams["ITNL_YN"] + "' = 'N' AND Z1.ITNL_YN = '무' ))");
            strSql.AppendLine("       ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");
            #endregion

            strSql.Clear();
            strSql.AppendLine("DECLARE @FDATE VARCHAR(10) = '" + dicParams["DATE_F"] + "' , @TDATE VARCHAR(10) = '" + dicParams["DATE_T"] + "';");
            strSql.AppendLine("       WITH SCRAP AS(                                                                                                                           ");
            strSql.AppendLine("  SELECT *                                                                                                                                             ");
            strSql.AppendLine("  FROM(                                                                                                                                                 ");
            strSql.AppendLine("          SELECT Y1.OPN_GB                                                                                                                   ");
            strSql.AppendLine("              , Y1.GB                                                                                                                                     ");
            strSql.AppendLine("              , Y1.MAIPCHERID                                                                                                                     ");
            strSql.AppendLine("              , Y1.DEALER_NM                                                                                                                      ");
            strSql.AppendLine("              , Y1.CHRG_ID                                                                                                                            ");
            strSql.AppendLine("              , Y1.WGT                                                                                                                                   ");
            strSql.AppendLine("              , Y2.SUM_CHAGAM                                                                                                                  ");
            strSql.AppendLine("              , Y1.IN_CNT                                                                                                                              ");
            strSql.AppendLine("              , Y1.CHAGAM_CNT                                                                                                                   ");
            strSql.AppendLine("              , Y1.CHAGAM_WGT                                                                                                                  ");
            strSql.AppendLine("              , (Y2.SUM_CHAGAM - Y1.CHAGAM_WGT)AS ADMT_WGT                                                   ");
            strSql.AppendLine("              , ISNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT                                                                   ");
            strSql.AppendLine("              , Y1.ITNL_YN                                                                                                                             ");
            strSql.AppendLine("              , Y3.OPN_DATE                                                                                                                         ");
            strSql.AppendLine("              , Y3.OPN_RMK                                                                                                                           ");
            strSql.AppendLine("              , @FDATE AS DATE_T                                                                                                               ");
            strSql.AppendLine("              , @TDATE AS DATE_F                                                                                                               ");
            strSql.AppendLine("          FROM(                                                                                                                                            ");
            strSql.AppendLine("                  SELECT '1' AS OPN_GB                                                                                                          ");
            strSql.AppendLine("                      , '스크랩' AS GB                                                                                                                    ");
            strSql.AppendLine("                      , A1.MAIPCHERID                                                                                                                ");
            strSql.AppendLine("                      , A2.DEALER_NM                                                                                                                ");
            strSql.AppendLine("                      , A2.CHRG_ID                                                                                                                      ");
            strSql.AppendLine("                      , SUM(A1.IWEIGHT) AS WGT                                                                                              ");
            strSql.AppendLine("                      , ISNULL(A3.CHAGAM, ISNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM                             ");
            strSql.AppendLine("                      , COUNT(*) AS IN_CNT                                                                                                      ");
            strSql.AppendLine("                      , ISNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT                                                           ");
            strSql.AppendLine("                     , SUM(A1.ICHAGAM) AS CHAGAM_WGT                                                                             ");
            strSql.AppendLine("                     , CASE WHEN ISNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN          ");
            strSql.AppendLine("                     FROM MESURING A1                                                                                                            ");
            strSql.AppendLine("                     LEFT JOIN ACC_DEALER_CD A2                                                                                         ");
            strSql.AppendLine("                     ON A1.MAIPCHERID = A2.DEALER_CD                                                                                 ");
            strSql.AppendLine("                     LEFT JOIN(SELECT X1.DEALER_CD                                                                                     ");
            strSql.AppendLine("                                     , SUM(X1.CHAGAM) AS CHAGAM                                                                         ");
            strSql.AppendLine("                                     , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT                ");
            strSql.AppendLine("                                 FROM MESURE_ISPT_INFO X1                                                                                ");
            strSql.AppendLine("                 LEFT JOIN MESURING X2 ON X1.JUNPYOID = X2.JUNPYOID                                               ");
            strSql.AppendLine("                 WHERE X2.J_DATE BETWEEN @FDATE AND @TDATE                                                          ");
            strSql.AppendLine("                                 GROUP BY X1.DEALER_CD) A3                                                                                ");
            strSql.AppendLine("                     ON A1.MAIPCHERID = A3.DEALER_CD                                                                                 ");
            strSql.AppendLine("                     AND A1.J_DATE BETWEEN @FDATE AND @TDATE                                                           ");
            strSql.AppendLine("                     LEFT JOIN(SELECT X1.MAIPCHERID                                                                                    ");
            strSql.AppendLine("                                     , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1                                                   ");
            strSql.AppendLine("                                                 WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT                        ");
            strSql.AppendLine("                                 FROM MESURING X1                                                                                                  ");
            strSql.AppendLine("                                 LEFT JOIN MESURING_SEQ X2                                                                                 ");
            strSql.AppendLine("                                     ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                                ");
            strSql.AppendLine("                                                         FROM MESURING_SEQ W                                                                   ");
            strSql.AppendLine("                                                         WHERE X2.JUNPYOID = W.JUNPYOID                                               ");
            strSql.AppendLine("                                                         AND W.CHAGAM > 0                                                                            ");
            strSql.AppendLine("                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ)                                              ");
            strSql.AppendLine("                                 LEFT JOIN JAJAE X3                                                                                                  ");
            strSql.AppendLine("                                     ON X1.J_SERIAL = X3.J_SERIAL                                                                            ");
            strSql.AppendLine("                                 WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                            ");
            strSql.AppendLine("                                 AND X3.DAEGUBUN IN('고철A', '고철B')                                                                       ");
            strSql.AppendLine("                                 GROUP BY X1.MAIPCHERID) A4                                                                                 ");
            strSql.AppendLine("                     ON A1.MAIPCHERID = A4.MAIPCHERID                                                                                  ");
            strSql.AppendLine("                     LEFT JOIN(SELECT TOP 1 X1.JUNPYOID, X1.CHAGAM                                                         ");
            strSql.AppendLine("                                 FROM MESURING_SEQ X1                                                                                           ");
            strSql.AppendLine("                                 WHERE X1.CHAGAM > 0) A5                                                                                        ");
            strSql.AppendLine("                     ON A1.JUNPYOID = A5.JUNPYOID                                                                                          ");
            strSql.AppendLine("                     LEFT JOIN JAJAE B1                                                                    ");
            strSql.AppendLine("                     ON A1.J_SERIAL = B1.J_SERIAL                                                          ");
            strSql.AppendLine("                 WHERE A1.J_DATE BETWEEN @FDATE AND @TDATE                                                 ");
            strSql.AppendLine("                     AND A1.KERATYPE = '입고'                                                              ");
            strSql.AppendLine("                     AND B1.DAEGUBUN IN('고철A', '고철B')                                                  ");
            strSql.AppendLine("                 GROUP BY A1.MAIPCHERID, A2.DEALER_NM, A3.CHAGAM, A4.CHAGAM_CNT, A3.ITNL_CNT, A2.CHRG_ID           ");
            strSql.AppendLine("                 ) Y1                                                                                      ");
            strSql.AppendLine("         LEFT JOIN(SELECT X1.MAIPCHERID                                                                    ");
            strSql.AppendLine("                         , SUM(ISNULL(X3.CHAGAM, ISNULL(X2.CHAGAM, ISNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM  ");
            strSql.AppendLine("                         FROM MESURING X1                                                                  ");
            strSql.AppendLine("                         LEFT JOIN MESURING_SEQ X2                                                         ");
            strSql.AppendLine("                         ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                         ");
            strSql.AppendLine("                                             FROM MESURING_SEQ W                                           ");
            strSql.AppendLine("                                             WHERE X2.JUNPYOID = W.JUNPYOID                                ");
            strSql.AppendLine("                                                 AND W.CHAGAM > 0                                          ");
            strSql.AppendLine("                                             ORDER BY W.JUNPYOID, W.CHG_SEQ)                               ");
            strSql.AppendLine("                          LEFT JOIN MESURE_ISPT_INFO X3                                                     ");
            strSql.AppendLine("                          ON X1.JUNPYOID = X3.JUNPYOID                                                      ");
            strSql.AppendLine("                          LEFT JOIN JAJAE X4                                                                ");
            strSql.AppendLine("                          ON X1.J_SERIAL = X4.J_SERIAL                                                      ");
            strSql.AppendLine("                      WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                             ");
            strSql.AppendLine("                          AND X1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                          AND X4.DAEGUBUN IN('고철A', '고철B')                                              ");
            strSql.AppendLine("                      GROUP BY X1.MAIPCHERID) Y2                                                           ");
            strSql.AppendLine("                     ON Y1.MAIPCHERID = Y2.MAIPCHERID                                                              ");
            strSql.AppendLine("                 LEFT JOIN(SELECT X1.DEALER_CD                                                                     ");
            strSql.AppendLine("                                     , X1.OPN_GB                                                                   ");
            strSql.AppendLine("                                     , MAX(X1.OPN_DATE) AS OPN_DATE                                                ");
            strSql.AppendLine("                                     , MAX(X1.OPN_RMK) AS OPN_RMK                                                  ");
            strSql.AppendLine("                                 FROM MESURE_OPN_HISTORY X1                                                        ");
            strSql.AppendLine("                                 WHERE REPLACE(X1.OPN_RMK, ' ', '') <> ''                                          ");
            strSql.AppendLine("                                 GROUP BY X1.DEALER_CD, X1.OPN_GB ) Y3                                             ");
            strSql.AppendLine("                     ON Y1.MAIPCHERID = Y3.DEALER_CD                                                               ");
            strSql.AppendLine("                     AND Y1.OPN_GB = Y3.OPN_GB                                                                     ");
            strSql.AppendLine("                     AND Y3.OPN_DATE BETWEEN @FDATE AND @TDATE                                                     ");
            strSql.AppendLine("                 LEFT JOIN(SELECT X1.MAIPCHERID                                                                    ");
            strSql.AppendLine("                                 , COUNT(*) AS RETURN_CNT                                                          ");
            strSql.AppendLine("                                 FROM MESURING X1                                                                  ");
            strSql.AppendLine("                             WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                             ");
            strSql.AppendLine("                                 AND X1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                                 AND X1.J_SERIAL = 4049042--스크랩반품                                             ");
            strSql.AppendLine("                             GROUP BY X1.MAIPCHERID) Y4                                                            ");
            strSql.AppendLine("                     ON Y1.MAIPCHERID = Y4.MAIPCHERID                                                              ");
            strSql.AppendLine("                 WHERE Y1.CHAGAM_CNT > 0                                                                           ");
            strSql.AppendLine("             ) Z1                                                                            ");
            strSql.AppendLine("         ), SHREDER AS(                                                                       ");
            strSql.AppendLine("         SELECT *                                                                       ");
            strSql.AppendLine("         FROM(                                                                                                         ");
            strSql.AppendLine("                SELECT Y1.OPN_GB                                                                               ");
            strSql.AppendLine("                     , Y1.GB                                                                                   ");
            strSql.AppendLine("                     , Y1.MAIPCHERID                                                                           ");
            strSql.AppendLine("                     , Y1.DEALER_NM                                                                            ");
            strSql.AppendLine("                     , Y1.CHRG_ID                                                                            ");
            strSql.AppendLine("                     , Y1.WGT                                                                                  ");
            strSql.AppendLine("                     , Y2.SUM_CHAGAM                                                                           ");
            strSql.AppendLine("                     , Y1.IN_CNT                                                                               ");
            strSql.AppendLine("                     , Y1.CHAGAM_CNT                                                                           ");
            strSql.AppendLine("                     , Y1.CHAGAM_WGT                                                                           ");
            strSql.AppendLine("                     , (Y2.SUM_CHAGAM - Y1.CHAGAM_WGT)AS ADMT_WGT                                                   ");
            strSql.AppendLine("                     , ISNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT                                                  ");
            strSql.AppendLine("                     , Y1.ITNL_YN                                                                              ");
            strSql.AppendLine("                     , Y3.OPN_DATE                                                                              ");
            strSql.AppendLine("                     , Y3.OPN_RMK                                                                               ");
            strSql.AppendLine("                     , @FDATE AS DATE_T                                                                         ");
            strSql.AppendLine("                     , @TDATE AS DATE_F                                                                         ");
            strSql.AppendLine("                  FROM(                                                                                                       ");
            strSql.AppendLine("                         SELECT '2' AS OPN_GB                                                                   ");
            strSql.AppendLine("                              , '슈레더' AS GB                                                                  ");
            strSql.AppendLine("                              , A1.MAIPCHERID                                                                   ");
            strSql.AppendLine("                              , A2.DEALER_NM                                                                    ");
            strSql.AppendLine("                              , A2.CHRG_ID                                                                  ");
            strSql.AppendLine("                              , SUM(A1.IWEIGHT) AS WGT                                                          ");
            strSql.AppendLine("                              , ISNULL(A3.CHAGAM, ISNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM                         ");
            strSql.AppendLine("                              , COUNT(*) AS IN_CNT                                                              ");
            strSql.AppendLine("                              , ISNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT                                          ");
            strSql.AppendLine("                              , SUM(A1.ICHAGAM) AS CHAGAM_WGT                                                   ");
            strSql.AppendLine("                              , CASE WHEN ISNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN         ");
            strSql.AppendLine("                           FROM MESURING A1                                                                     ");
            strSql.AppendLine("                           LEFT JOIN ACC_DEALER_CD A2                                                           ");
            strSql.AppendLine("                             ON A1.MAIPCHERID = A2.DEALER_CD                                                    ");
            strSql.AppendLine("                           LEFT JOIN(SELECT X1.DEALER_CD                                                        ");
            strSql.AppendLine("                                            , SUM(X1.CHAGAM) AS CHAGAM                                         ");
            strSql.AppendLine("                                            , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT         ");
            strSql.AppendLine("                                         FROM MESURE_ISPT_INFO X1                                              ");
            strSql.AppendLine("                         LEFT JOIN MESURING X2 ON X1.JUNPYOID = X2.JUNPYOID                             ");
            strSql.AppendLine("                         WHERE X2.J_DATE BETWEEN @FDATE AND @TDATE                                    ");
            strSql.AppendLine("                                        GROUP BY X1.DEALER_CD) A3                                              ");
            strSql.AppendLine("                             ON A1.MAIPCHERID = A3.DEALER_CD                                                   ");
            strSql.AppendLine("                            AND A1.J_DATE BETWEEN @FDATE AND @TDATE                                            ");
            strSql.AppendLine("                           LEFT JOIN(SELECT X1.MAIPCHERID                                                      ");
            strSql.AppendLine("                                            , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1                             ");
            strSql.AppendLine("                                                         WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT         ");
            strSql.AppendLine("                                         FROM MESURING X1                                                      ");
            strSql.AppendLine("                                         LEFT JOIN MESURING_SEQ X2                                             ");
            strSql.AppendLine("                                           ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                           ");
            strSql.AppendLine("                                                                FROM MESURING_SEQ W                            ");
            strSql.AppendLine("                                                               WHERE X2.JUNPYOID = W.JUNPYOID                  ");
            strSql.AppendLine("                                                                 AND W.CHAGAM > 0                              ");
            strSql.AppendLine("                                                               ORDER BY W.JUNPYOID, W.CHG_SEQ)                 ");
            strSql.AppendLine("                                         LEFT JOIN JAJAE X3                                                    ");
            strSql.AppendLine("                                           ON X1.J_SERIAL = X3.J_SERIAL                                        ");
            strSql.AppendLine("                                        WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                              ");
            strSql.AppendLine("                                          AND X3.DAEGUBUN = '슈레더'                                           ");
            strSql.AppendLine("                                        GROUP BY X1.MAIPCHERID) A4                                            ");
            strSql.AppendLine("                             ON A1.MAIPCHERID = A4.MAIPCHERID                                                 ");
            strSql.AppendLine("                           LEFT JOIN(SELECT TOP 1 X1.JUNPYOID, X1.CHAGAM                                      ");
            strSql.AppendLine("                                         FROM MESURING_SEQ X1                                                 ");
            strSql.AppendLine("                                        WHERE X1.CHAGAM > 0) A5                                               ");
            strSql.AppendLine("                             ON A1.JUNPYOID = A5.JUNPYOID                                                     ");
            strSql.AppendLine("                           LEFT JOIN JAJAE B1                                                                 ");
            strSql.AppendLine("                             ON A1.J_SERIAL = B1.J_SERIAL                                                     ");
            strSql.AppendLine("                          WHERE A1.J_DATE BETWEEN @FDATE AND @TDATE                                           ");
            strSql.AppendLine("                            AND A1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                            AND B1.DAEGUBUN IN('슈레더')                                                      ");
            strSql.AppendLine("                          GROUP BY A1.MAIPCHERID, A2.DEALER_NM, A3.CHAGAM, A4.CHAGAM_CNT, A3.ITNL_CNT ,A2.CHRG_ID        ");
            strSql.AppendLine("                       ) Y1                                                                                   ");
            strSql.AppendLine("                  LEFT JOIN(SELECT X1.MAIPCHERID                                                                   ");
            strSql.AppendLine("                                  , SUM(ISNULL(X3.CHAGAM, ISNULL(X2.CHAGAM, ISNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            strSql.AppendLine("                               FROM MESURING X1                                                                    ");
            strSql.AppendLine("                               LEFT JOIN MESURING_SEQ X2                                                           ");
            strSql.AppendLine("                                 ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                         ");
            strSql.AppendLine("                                                      FROM MESURING_SEQ W                                          ");
            strSql.AppendLine("                                                     WHERE X2.JUNPYOID = W.JUNPYOID                                ");
            strSql.AppendLine("                                                       AND W.CHAGAM > 0                                            ");
            strSql.AppendLine("                                                     ORDER BY W.JUNPYOID, W.CHG_SEQ)                               ");
            strSql.AppendLine("                               LEFT JOIN MESURE_ISPT_INFO X3                                                       ");
            strSql.AppendLine("                                 ON X1.JUNPYOID = X3.JUNPYOID                                                 ");
            strSql.AppendLine("                               LEFT JOIN JAJAE X4                                                             ");
            strSql.AppendLine("                                 ON X1.J_SERIAL = X4.J_SERIAL                                                 ");
            strSql.AppendLine("                              WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                       ");
            strSql.AppendLine("                                AND X1.KERATYPE = '입고'                                                      ");
            strSql.AppendLine("                                AND X4.DAEGUBUN IN('슈레더')                                                  ");
            strSql.AppendLine("                              GROUP BY X1.MAIPCHERID ) Y2                                                     ");
            strSql.AppendLine("                    ON Y1.MAIPCHERID = Y2.MAIPCHERID                                                          ");
            strSql.AppendLine("                  LEFT JOIN(SELECT X1.DEALER_CD                                                               ");
            strSql.AppendLine("                                     , X1.OPN_GB                                                              ");
            strSql.AppendLine("                                     , MAX(X1.OPN_DATE) AS OPN_DATE                                           ");
            strSql.AppendLine("                                     , MAX(X1.OPN_RMK) AS OPN_RMK                                             ");
            strSql.AppendLine("                                  FROM MESURE_OPN_HISTORY X1                                                  ");
            strSql.AppendLine("                                 WHERE REPLACE(X1.OPN_RMK, ' ', '') <> ''                                     ");
            strSql.AppendLine("                                 GROUP BY X1.DEALER_CD, X1.OPN_GB ) Y3                                        ");
            strSql.AppendLine("                    ON Y1.MAIPCHERID = Y3.DEALER_CD                                                                                                    ");
            strSql.AppendLine("                   AND Y1.OPN_GB = Y3.OPN_GB                                                                                                                ");
            strSql.AppendLine("                   AND Y3.OPN_DATE BETWEEN @FDATE AND @TDATE                                                                                ");
            strSql.AppendLine("                  LEFT JOIN(SELECT X1.MAIPCHERID                                                                                                          ");
            strSql.AppendLine("                                  , COUNT(*) AS RETURN_CNT                                                                                                     ");
            strSql.AppendLine("                               FROM MESURING X1                                                                                                                        ");
            strSql.AppendLine("                              WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                                                   ");
            strSql.AppendLine("                                AND X1.KERATYPE = '입고'                                                                                                              ");
            strSql.AppendLine("                                AND X1.J_SERIAL = 5050042--슈레더반품                                                                                     ");
            strSql.AppendLine("                              GROUP BY X1.MAIPCHERID) Y4                                                                                                        ");
            strSql.AppendLine("                   ON Y1.MAIPCHERID = Y4.MAIPCHERID                                                                                                       ");
            strSql.AppendLine("                 WHERE Y1.CHAGAM_CNT > 0                                                                                                                        ");
            strSql.AppendLine("              ) Z1                                                                                                                                                                     ");
            strSql.AppendLine("          )                                                                                                                                                                              ");
            strSql.AppendLine("          SELECT T.*FROM                                                                                                                                                  ");
            strSql.AppendLine("         (SELECT * FROM SCRAP                                                                                                                                        ");
            strSql.AppendLine("          UNION ALL                                                                                                                                                            ");
            strSql.AppendLine("          SELECT * FROM SHREDER                                                                                                                                    ");
            strSql.AppendLine("         )T                                                                                                                                                                             ");
            strSql.AppendLine("       WHERE (( '" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 )                                                                                  ");
            strSql.AppendLine("              OR                                                                                                                                                                        ");
            strSql.AppendLine("              ( '" + dicParams["FIND_IDX"] + "' = '0' AND T.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))   ");
            strSql.AppendLine("         AND (( '" + dicParams["ITNL_YN"] + "' = 'ALL' AND 1 = 1 )                                                                                    ");
            strSql.AppendLine("              OR                                                                                                                                                                         ");
            strSql.AppendLine("              ( '" + dicParams["ITNL_YN"] + "' = 'Y' AND T.ITNL_YN = '유' )                                                                         ");
            strSql.AppendLine("              OR                                                                                                                                                                         ");
            strSql.AppendLine("              ( '" + dicParams["ITNL_YN"] + "' = 'N' AND T.ITNL_YN = '무' ))                                                                         ");
            strSql.AppendLine("         ORDER BY                                                                                                                                                                  ");
            strSql.AppendLine("         T.OPN_GB ASC, T.WGT DESC, REPLACE(T.DEALER_NM, '(주)', '') ASC                                                                    ");


            #region mariaDB
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.CHRG_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y1.CHAGAM_WGT, 0) - IFNULL(Y2.SUM_CHAGAM, 0) AS ADMT_WGT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '1' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '스크랩' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN A22.JUNPYOID IS NULL THEN A23.EMP_NM ELSE A22.CHRG_NM END AS CHRG_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      LEFT JOIN INLIST A22 ");
            //strSql.AppendFormat("\r\n                        ON A1.IPCHULGO_MAIPID = A22.J_ID ");
            //strSql.AppendFormat("\r\n                      LEFT JOIN HR_EMP_BASIS A23 ");
            //strSql.AppendFormat("\r\n                        ON A2.CHRG_ID = A23.EMP_ID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD   ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB   ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE AS OPN_DATE  ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK AS OPN_RMK   ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1   ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = ( SELECT MAX(W.OPN_DATE) AS OPN_DATE  ");
            //strSql.AppendFormat("\r\n                                                 FROM MESURE_OPN_HISTORY W ");
            //strSql.AppendFormat("\r\n                                                WHERE REPLACE(W.OPN_RMK, ' ', '') <> ''  ");
            //strSql.AppendFormat("\r\n                                                  AND W.OPN_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                                  AND X1.DEALER_CD = W.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                                  AND X1.OPN_GB = W.OPN_GB ");
            //strSql.AppendFormat("\r\n                                                GROUP BY W.DEALER_CD, W.OPN_GB ) ");
            //strSql.AppendFormat("\r\n                         GROUP BY X1.DEALER_CD, X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                         ORDER BY X1.OPN_DATE DESC ) Y3 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 4049042 #스크랩반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n           UNION ALL ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.CHRG_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y1.CHAGAM_WGT, 0) - IFNULL(Y2.SUM_CHAGAM, 0) AS ADMT_WGT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '2' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '슈레더' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN A22.JUNPYOID IS NULL THEN A23.EMP_NM ELSE A22.CHRG_NM END AS CHRG_NM ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      LEFT JOIN INLIST A22 ");
            //strSql.AppendFormat("\r\n                        ON A1.IPCHULGO_MAIPID = A22.J_ID ");
            //strSql.AppendFormat("\r\n                      LEFT JOIN HR_EMP_BASIS A23 ");
            //strSql.AppendFormat("\r\n                        ON A2.CHRG_ID = A23.EMP_ID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN = '슈레더' ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD   ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB   ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE AS OPN_DATE  ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK AS OPN_RMK   ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1   ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = ( SELECT MAX(W.OPN_DATE) AS OPN_DATE  ");
            //strSql.AppendFormat("\r\n                                                 FROM MESURE_OPN_HISTORY W ");
            //strSql.AppendFormat("\r\n                                                WHERE REPLACE(W.OPN_RMK, ' ', '') <> ''  ");
            //strSql.AppendFormat("\r\n                                                  AND W.OPN_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                                  AND X1.DEALER_CD = W.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                                  AND X1.OPN_GB = W.OPN_GB ");
            //strSql.AppendFormat("\r\n                                                GROUP BY W.DEALER_CD, W.OPN_GB ) ");
            //strSql.AppendFormat("\r\n                         GROUP BY X1.DEALER_CD, X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                         ORDER BY X1.OPN_DATE DESC ) Y3    ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD    ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB   ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 5050042 #슈레더반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            //strSql.AppendFormat("\r\n  WHERE (( '{0}' = '' AND 1 = 1 )", dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR ");
            //strSql.AppendFormat("\r\n         ( '{0}' = '0' AND Z1.DEALER_NM LIKE '%{1}%' ))", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n    AND (( '{0}' = 'ALL' AND 1 = 1 ) ", dicParams["ITNL_YN"]);
            //strSql.AppendFormat("\r\n         OR ");
            //strSql.AppendFormat("\r\n         ( '{0}' = 'Y' AND Z1.ITNL_YN = '유' )", dicParams["ITNL_YN"]);
            //strSql.AppendFormat("\r\n         OR ");
            //strSql.AppendFormat("\r\n         ( '{0}' = 'N' AND Z1.ITNL_YN = '무' ))", dicParams["ITNL_YN"]);
            //strSql.AppendFormat("\r\n  ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");
            #endregion

            #region[2021-04-22 이전쿼리]

            //strSql.Clear();
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '1' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '스크랩' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 4049042 #스크랩반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n           UNION ALL ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '2' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '슈레더' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN = '슈레더' ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 5050042 #슈레더반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            //strSql.AppendFormat("\r\n  WHERE (( '{0}' = '' AND 1 = 1 )", dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n         OR ");
            //strSql.AppendFormat("\r\n         ( '{0}' = '0' AND Z1.DEALER_NM LIKE '%{1}%' ))", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            //strSql.AppendFormat("\r\n    AND (( '{0}' = 'ALL' AND 1 = 1 ) ", dicParams["ITNL_YN"]);
            //strSql.AppendFormat("\r\n         OR ");
            //strSql.AppendFormat("\r\n         ( '{0}' = 'Y' AND Z1.ITNL_YN = '유' )", dicParams["ITNL_YN"]);
            //strSql.AppendFormat("\r\n         OR ");
            //strSql.AppendFormat("\r\n         ( '{0}' = 'N' AND Z1.ITNL_YN = '무' ))", dicParams["ITNL_YN"]);
            //strSql.AppendFormat("\r\n  ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");

            #endregion[2021-04-22 이전쿼리]

            #region[2021-04-06 이전쿼리]

            //strSql.Clear();
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '1' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '스크랩' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                                      , COUNT(*) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X1.ICHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN ( SELECT Y1.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                          , Y1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                       FROM MESURING_SEQ Y1 ");
            //strSql.AppendFormat("\r\n                                      WHERE Y1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                      LIMIT 1 ) X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X2.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 4049042 #스크랩반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           UNION ALL ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '2' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '슈레더' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                                      , COUNT(*) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X1.ICHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN ( SELECT Y1.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                          , Y1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                       FROM MESURING_SEQ Y1 ");
            //strSql.AppendFormat("\r\n                                      WHERE Y1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                      LIMIT 1 ) X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X2.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 5050042 #슈레더반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            //strSql.AppendFormat("\r\n  ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");

            #endregion[2021-04-06 이전쿼리]

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetSummaryScrap(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("DECLARE @FDATE VARCHAR(10) = '"+ dicParams["DATE_F"] + "' , @TDATE VARCHAR(10) = '"+ dicParams["DATE_T"] + "';");
            strSql.AppendLine("SELECT *                                                                                                    ");
            strSql.AppendLine("  FROM (                                                                                                    ");
            strSql.AppendLine("          SELECT Y1.OPN_GB                                                                                  ");
            strSql.AppendLine("              , Y1.GB                                                                                       ");
            strSql.AppendLine("              , Y1.MAIPCHERID                                                                               ");
            strSql.AppendLine("              , Y1.DEALER_NM                                                                                ");
            strSql.AppendLine("              , Y1.WGT                                                                                      ");
            strSql.AppendLine("              , Y2.SUM_CHAGAM                                                                               ");
            strSql.AppendLine("              , Y1.IN_CNT                                                                                   ");
            strSql.AppendLine("              , Y1.CHAGAM_CNT                                                                               ");
            strSql.AppendLine("              , Y1.CHAGAM_WGT                                                                               ");
            strSql.AppendLine("              , ISNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT                                                      ");
            strSql.AppendLine("              , Y1.ITNL_YN                                                                                  ");
            strSql.AppendLine("              , Y3.OPN_DATE                                                                                 ");
            strSql.AppendLine("              , Y3.OPN_RMK                                                                                  ");
            strSql.AppendLine("              , @TDATE AS DATE_T                                                                            ");
            strSql.AppendLine("              , @FDATE AS DATE_F                                                                            ");
            strSql.AppendLine("          FROM(                                                                                             ");
            strSql.AppendLine("                  SELECT '1' AS OPN_GB                                                                      ");
            strSql.AppendLine("                      , '스크랩' AS GB                                                                      ");
            strSql.AppendLine("                      , A1.MAIPCHERID                                                                       ");
            strSql.AppendLine("                      , A2.DEALER_NM                                                                        ");
            strSql.AppendLine("                      , SUM(A1.IWEIGHT) AS WGT                                                              ");
            strSql.AppendLine("                      , ISNULL(A3.CHAGAM, ISNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM                             ");
            strSql.AppendLine("                      , COUNT(*) AS IN_CNT                                                                  ");
            strSql.AppendLine("                      , ISNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT                                              ");
            strSql.AppendLine("                      , SUM(A1.ICHAGAM) AS CHAGAM_WGT                                                       ");
            strSql.AppendLine("                      , CASE WHEN ISNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN             ");
            strSql.AppendLine("                      FROM MESURING A1                                                                      ");
            strSql.AppendLine("                      LEFT JOIN ACC_DEALER_CD A2                                                            ");
            strSql.AppendLine("                      ON A1.MAIPCHERID = A2.DEALER_CD                                                       ");
            strSql.AppendLine("                      LEFT JOIN(SELECT X1.DEALER_CD                                                         ");
            strSql.AppendLine("                                      , SUM(X1.CHAGAM) AS CHAGAM                                            ");
            strSql.AppendLine("                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT            ");
            strSql.AppendLine("                                  FROM MESURE_ISPT_INFO X1                                                  ");
            strSql.AppendLine("                  LEFT JOIN MESURING X2 ON X1.JUNPYOID = X2.JUNPYOID                                      ");
            strSql.AppendLine("                  WHERE X2.J_DATE BETWEEN @FDATE AND @TDATE                                     ");
            strSql.AppendLine("                                  GROUP BY X1.DEALER_CD) A3                                                 ");
            strSql.AppendLine("                      ON A1.MAIPCHERID = A3.DEALER_CD                                                       ");
            strSql.AppendLine("                      AND A1.J_DATE BETWEEN @FDATE AND @TDATE                                               ");
            strSql.AppendLine("                      LEFT JOIN(SELECT X1.MAIPCHERID                                                        ");
            strSql.AppendLine("                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1                                ");
            strSql.AppendLine("                                                  WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT             ");
            strSql.AppendLine("                                  FROM MESURING X1                                                          ");
            strSql.AppendLine("                                  LEFT JOIN MESURING_SEQ X2                                                 ");
            strSql.AppendLine("                                      ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                             ");
            strSql.AppendLine("                                                          FROM MESURING_SEQ W                               ");
            strSql.AppendLine("                                                          WHERE X2.JUNPYOID = W.JUNPYOID                    ");
            strSql.AppendLine("                                                          AND W.CHAGAM > 0                                  ");
            strSql.AppendLine("                                                          ORDER BY W.JUNPYOID, W.CHG_SEQ)                   ");
            strSql.AppendLine("                                  LEFT JOIN JAJAE X3                                                        ");
            strSql.AppendLine("                                      ON X1.J_SERIAL = X3.J_SERIAL                                          ");
            strSql.AppendLine("                                  WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                 ");
            strSql.AppendLine("                                  AND X3.DAEGUBUN IN('고철A', '고철B')                                      ");
            strSql.AppendLine("                                  GROUP BY X1.MAIPCHERID) A4                                                ");
            strSql.AppendLine("                      ON A1.MAIPCHERID = A4.MAIPCHERID                                                      ");
            strSql.AppendLine("                      LEFT JOIN(SELECT TOP 1 X1.JUNPYOID, X1.CHAGAM                                         ");
            strSql.AppendLine("                                  FROM MESURING_SEQ X1                                                      ");
            strSql.AppendLine("                                  WHERE X1.CHAGAM > 0) A5                                                   ");
            strSql.AppendLine("                      ON A1.JUNPYOID = A5.JUNPYOID                                                          ");
            strSql.AppendLine("                      LEFT JOIN JAJAE B1                                                                    ");
            strSql.AppendLine("                      ON A1.J_SERIAL = B1.J_SERIAL                                                          ");
            strSql.AppendLine("                  WHERE A1.J_DATE BETWEEN @FDATE AND @TDATE                                                 ");
            strSql.AppendLine("                      AND A1.KERATYPE = '입고'                                                              ");
            strSql.AppendLine("                      AND B1.DAEGUBUN IN('고철A', '고철B')                                                  ");
            strSql.AppendLine("                  GROUP BY A1.MAIPCHERID, A2.DEALER_NM, A3.CHAGAM, A4.CHAGAM_CNT, A3.ITNL_CNT               ");
            strSql.AppendLine("                  ) Y1                                                                                      ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.MAIPCHERID                                                                    ");
            strSql.AppendLine("                          , SUM(ISNULL(X3.CHAGAM, ISNULL(X2.CHAGAM, ISNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM  ");
            strSql.AppendLine("                          FROM MESURING X1                                                                  ");
            strSql.AppendLine("                          LEFT JOIN MESURING_SEQ X2                                                         ");
            strSql.AppendLine("                          ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                         ");
            strSql.AppendLine("                                              FROM MESURING_SEQ W                                           ");
            strSql.AppendLine("                                              WHERE X2.JUNPYOID = W.JUNPYOID                                ");
            strSql.AppendLine("                                                  AND W.CHAGAM > 0                                          ");
            strSql.AppendLine("                                              ORDER BY W.JUNPYOID, W.CHG_SEQ)                               ");
            strSql.AppendLine("                          LEFT JOIN MESURE_ISPT_INFO X3                                                     ");
            strSql.AppendLine("                          ON X1.JUNPYOID = X3.JUNPYOID                                                      ");
            strSql.AppendLine("                          LEFT JOIN JAJAE X4                                                                ");
            strSql.AppendLine("                          ON X1.J_SERIAL = X4.J_SERIAL                                                      ");
            strSql.AppendLine("                      WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                             ");
            strSql.AppendLine("                          AND X1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                          AND X4.DAEGUBUN IN('고철A', '고철B')                                              ");
            strSql.AppendLine("                      GROUP BY X1.MAIPCHERID ) Y2                                                           ");
            strSql.AppendLine("              ON Y1.MAIPCHERID = Y2.MAIPCHERID                                                              ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.DEALER_CD                                                                     ");
            strSql.AppendLine("                              , X1.OPN_GB                                                                   ");
            strSql.AppendLine("                              , MAX(X1.OPN_DATE) AS OPN_DATE                                                ");
            strSql.AppendLine("                              , MAX(X1.OPN_RMK) AS OPN_RMK                                                  ");
            strSql.AppendLine("                          FROM MESURE_OPN_HISTORY X1                                                        ");
            strSql.AppendLine("                          WHERE REPLACE(X1.OPN_RMK, ' ', '') <> ''                                          ");
            strSql.AppendLine("                          GROUP BY X1.DEALER_CD, X1.OPN_GB ) Y3                                             ");
            strSql.AppendLine("              ON Y1.MAIPCHERID = Y3.DEALER_CD                                                               ");
            strSql.AppendLine("              AND Y1.OPN_GB = Y3.OPN_GB                                                                     ");
            strSql.AppendLine("              AND Y3.OPN_DATE BETWEEN @FDATE AND @TDATE                                                     ");
            strSql.AppendLine("          LEFT JOIN(SELECT X1.MAIPCHERID                                                                    ");
            strSql.AppendLine("                          , COUNT(*) AS RETURN_CNT                                                          ");
            strSql.AppendLine("                          FROM MESURING X1                                                                  ");
            strSql.AppendLine("                      WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                             ");
            strSql.AppendLine("                          AND X1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                          AND X1.J_SERIAL = 4049042--스크랩반품                                             ");
            strSql.AppendLine("                      GROUP BY X1.MAIPCHERID) Y4                                                            ");
            strSql.AppendLine("              ON Y1.MAIPCHERID = Y4.MAIPCHERID                                                              ");
            strSql.AppendLine("          WHERE Y1.CHAGAM_CNT > 0                                                                           ");
            strSql.AppendLine("      ) Z1                                                                                                  ");
            strSql.AppendLine("  ORDER BY Z1.WGT DESC, Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC                                ");

            #region MariaDB
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_F ", dicParams["DATE_F"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '1' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '스크랩' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD  ");
            //strSql.AppendFormat("\r\n                               , X1.OPN_GB  ");
            //strSql.AppendFormat("\r\n                               , MAX(X1.OPN_DATE) AS OPN_DATE ");
            //strSql.AppendFormat("\r\n                               , MAX(X1.OPN_RMK) AS OPN_RMK  ");
            //strSql.AppendFormat("\r\n                            FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                           WHERE REPLACE(X1.OPN_RMK, ' ', '') <> '' ");
            //strSql.AppendFormat("\r\n                           GROUP BY X1.DEALER_CD, X1.OPN_GB ) Y3 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD   ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB  ");
            //strSql.AppendFormat("\r\n             AND Y3.OPN_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 4049042 #스크랩반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            ////#0010
            //strSql.AppendFormat("\r\n  ORDER BY Z1.WGT DESC, Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");
            #endregion

            #region[2021-04-16 이전 쿼리]

            //strSql.Clear();
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_F ", dicParams["DATE_F"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '1' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '스크랩' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('고철A', '고철B') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 4049042 #스크랩반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            //strSql.AppendFormat("\r\n  ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");

            #endregion[2021-04-16 이전 쿼리]

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetSummaryShreder(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("DECLARE @FDATE VARCHAR(10) = '" + dicParams["DATE_F"] + "' , @TDATE VARCHAR(10) = '" + dicParams["DATE_T"] + "';");
            strSql.AppendLine("SELECT *                                                                                                   ");
            strSql.AppendLine("  FROM (                                                                                                   ");
            strSql.AppendLine("         SELECT Y1.OPN_GB                                                                                  ");
            strSql.AppendLine("              , Y1.GB                                                                                      ");
            strSql.AppendLine("              , Y1.MAIPCHERID                                                                              ");
            strSql.AppendLine("              , Y1.DEALER_NM                                                                               ");
            strSql.AppendLine("              , Y1.WGT                                                                                     ");
            strSql.AppendLine("              , Y2.SUM_CHAGAM                                                                              ");
            strSql.AppendLine("              , Y1.IN_CNT                                                                                  ");
            strSql.AppendLine("              , Y1.CHAGAM_CNT                                                                              ");
            strSql.AppendLine("              , Y1.CHAGAM_WGT                                                                              ");
            strSql.AppendLine("              , ISNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT                                                     ");
            strSql.AppendLine("              , Y1.ITNL_YN                                                                                 ");
            strSql.AppendLine("              , Y3.OPN_DATE                                                                                ");
            strSql.AppendLine("              , Y3.OPN_RMK                                                                                 ");
            strSql.AppendLine("              , @TDATE AS DATE_T                                                                           ");
            strSql.AppendLine("              , @FDATE AS DATE_F                                                                           ");
            strSql.AppendLine("           FROM(                                                                                           ");
            strSql.AppendLine("                  SELECT '2' AS OPN_GB                                                                     ");
            strSql.AppendLine("                       , '슈레더' AS GB                                                                    ");
            strSql.AppendLine("                       , A1.MAIPCHERID                                                                     ");
            strSql.AppendLine("                       , A2.DEALER_NM                                                                      ");
            strSql.AppendLine("                       , SUM(A1.IWEIGHT) AS WGT                                                            ");
            strSql.AppendLine("                       , ISNULL(A3.CHAGAM, ISNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM                           ");
            strSql.AppendLine("                       , COUNT(*) AS IN_CNT                                                                ");
            strSql.AppendLine("                       , ISNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT                                            ");
            strSql.AppendLine("                       , SUM(A1.ICHAGAM) AS CHAGAM_WGT                                                     ");
            strSql.AppendLine("                       , CASE WHEN ISNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN           ");
            strSql.AppendLine("                    FROM MESURING A1                                                                       ");
            strSql.AppendLine("                    LEFT JOIN ACC_DEALER_CD A2                                                             ");
            strSql.AppendLine("                      ON A1.MAIPCHERID = A2.DEALER_CD                                                      ");
            strSql.AppendLine("                    LEFT JOIN(SELECT X1.DEALER_CD                                                          ");
            strSql.AppendLine("                                     , SUM(X1.CHAGAM) AS CHAGAM                                            ");
            strSql.AppendLine("                                     , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT            ");
            strSql.AppendLine("                                  FROM MESURE_ISPT_INFO X1                                                 ");
            strSql.AppendLine("                  LEFT JOIN MESURING X2 ON X1.JUNPYOID = X2.JUNPYOID                                      ");
            strSql.AppendLine("                  WHERE X2.J_DATE BETWEEN @FDATE AND @TDATE                                     ");
            strSql.AppendLine("                                 GROUP BY X1.DEALER_CD) A3                                                 ");
            strSql.AppendLine("                      ON A1.MAIPCHERID = A3.DEALER_CD                                                      ");
            strSql.AppendLine("                     AND A1.J_DATE BETWEEN @FDATE AND @TDATE                                               ");
            strSql.AppendLine("                    LEFT JOIN(SELECT X1.MAIPCHERID                                                         ");
            strSql.AppendLine("                                     , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1                                ");
            strSql.AppendLine("                                                  WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT            ");
            strSql.AppendLine("                                  FROM MESURING X1                                                         ");
            strSql.AppendLine("                                  LEFT JOIN MESURING_SEQ X2                                                ");
            strSql.AppendLine("                                    ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                              ");
            strSql.AppendLine("                                                         FROM MESURING_SEQ W                               ");
            strSql.AppendLine("                                                        WHERE X2.JUNPYOID = W.JUNPYOID                     ");
            strSql.AppendLine("                                                          AND W.CHAGAM > 0                                 ");
            strSql.AppendLine("                                                        ORDER BY W.JUNPYOID, W.CHG_SEQ)                    ");
            strSql.AppendLine("                                  LEFT JOIN JAJAE X3                                                       ");
            strSql.AppendLine("                                    ON X1.J_SERIAL = X3.J_SERIAL                                           ");
            strSql.AppendLine("                                 WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                 ");
            strSql.AppendLine("                                   AND X3.DAEGUBUN = '슈레더'                                              ");
            strSql.AppendLine("                                 GROUP BY X1.MAIPCHERID) A4                                                ");
            strSql.AppendLine("                      ON A1.MAIPCHERID = A4.MAIPCHERID                                                     ");
            strSql.AppendLine("                    LEFT JOIN(SELECT TOP 1 X1.JUNPYOID, X1.CHAGAM                                          ");
            strSql.AppendLine("                                  FROM MESURING_SEQ X1                                                     ");
            strSql.AppendLine("                                 WHERE X1.CHAGAM > 0) A5                                                   ");
            strSql.AppendLine("                      ON A1.JUNPYOID = A5.JUNPYOID                                                         ");
            strSql.AppendLine("                    LEFT JOIN JAJAE B1                                                                     ");
            strSql.AppendLine("                      ON A1.J_SERIAL = B1.J_SERIAL                                                         ");
            strSql.AppendLine("                   WHERE A1.J_DATE BETWEEN @FDATE AND @TDATE                                               ");
            strSql.AppendLine("                     AND A1.KERATYPE = '입고'                                                              ");
            strSql.AppendLine("                     AND B1.DAEGUBUN IN('슈레더')                                                          ");
            strSql.AppendLine("                   GROUP BY A1.MAIPCHERID, A2.DEALER_NM, A3.CHAGAM, A4.CHAGAM_CNT, A3.ITNL_CNT             ");
            strSql.AppendLine("                ) Y1                                                                                       ");
            strSql.AppendLine("           LEFT JOIN(SELECT X1.MAIPCHERID                                                                  ");
            strSql.AppendLine("                           , SUM(ISNULL(X3.CHAGAM, ISNULL(X2.CHAGAM, ISNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM");
            strSql.AppendLine("                        FROM MESURING X1                                                                   ");
            strSql.AppendLine("                        LEFT JOIN MESURING_SEQ X2                                                          ");
            strSql.AppendLine("                          ON X1.JUNPYOID = (SELECT TOP 1 W.JUNPYOID                                        ");
            strSql.AppendLine("                                               FROM MESURING_SEQ W                                         ");
            strSql.AppendLine("                                              WHERE X2.JUNPYOID = W.JUNPYOID                               ");
            strSql.AppendLine("                                                AND W.CHAGAM > 0                                           ");
            strSql.AppendLine("                                              ORDER BY W.JUNPYOID, W.CHG_SEQ)                              ");
            strSql.AppendLine("                        LEFT JOIN MESURE_ISPT_INFO X3                                                      ");
            strSql.AppendLine("                          ON X1.JUNPYOID = X3.JUNPYOID                                                     ");
            strSql.AppendLine("                        LEFT JOIN JAJAE X4                                                                 ");
            strSql.AppendLine("                          ON X1.J_SERIAL = X4.J_SERIAL                                                     ");
            strSql.AppendLine("                       WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                           ");
            strSql.AppendLine("                         AND X1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                         AND X4.DAEGUBUN IN('슈레더')                                                      ");
            strSql.AppendLine("                       GROUP BY X1.MAIPCHERID ) Y2                                                         ");
            strSql.AppendLine("             ON Y1.MAIPCHERID = Y2.MAIPCHERID                                                              ");
            strSql.AppendLine("           LEFT JOIN(SELECT X1.DEALER_CD                                                                   ");
            strSql.AppendLine("                              , X1.OPN_GB                                                                  ");
            strSql.AppendLine("                              , MAX(X1.OPN_DATE) AS OPN_DATE                                               ");
            strSql.AppendLine("                              , MAX(X1.OPN_RMK) AS OPN_RMK                                                 ");
            strSql.AppendLine("                           FROM MESURE_OPN_HISTORY X1                                                      ");
            strSql.AppendLine("                          WHERE REPLACE(X1.OPN_RMK, ' ', '') <> ''                                         ");
            strSql.AppendLine("                          GROUP BY X1.DEALER_CD, X1.OPN_GB ) Y3                                            ");
            strSql.AppendLine("             ON Y1.MAIPCHERID = Y3.DEALER_CD                                                               ");
            strSql.AppendLine("            AND Y1.OPN_GB = Y3.OPN_GB                                                                      ");
            strSql.AppendLine("            AND Y3.OPN_DATE BETWEEN @FDATE AND @TDATE                                                      ");
            strSql.AppendLine("           LEFT JOIN(SELECT X1.MAIPCHERID                                                                  ");
            strSql.AppendLine("                           , COUNT(*) AS RETURN_CNT                                                        ");
            strSql.AppendLine("                        FROM MESURING X1                                                                   ");
            strSql.AppendLine("                       WHERE X1.J_DATE BETWEEN @FDATE AND @TDATE                                           ");
            strSql.AppendLine("                         AND X1.KERATYPE = '입고'                                                          ");
            strSql.AppendLine("                         AND X1.J_SERIAL = 5050042--슈레더반품                                             ");
            strSql.AppendLine("                       GROUP BY X1.MAIPCHERID) Y4                                                          ");
            strSql.AppendLine("            ON Y1.MAIPCHERID = Y4.MAIPCHERID                                                               ");
            strSql.AppendLine("          WHERE Y1.CHAGAM_CNT > 0                                                                          ");
            strSql.AppendLine("       ) Z1                                                                                                ");
            strSql.AppendLine("   ORDER BY Z1.WGT DESC, Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC                              ");

            #region MariaDB
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_F ", dicParams["DATE_F"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '2' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '슈레더' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN = '슈레더' ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD  ");
            //strSql.AppendFormat("\r\n                               , X1.OPN_GB  ");
            //strSql.AppendFormat("\r\n                               , MAX(X1.OPN_DATE) AS OPN_DATE ");
            //strSql.AppendFormat("\r\n                               , MAX(X1.OPN_RMK) AS OPN_RMK  ");
            //strSql.AppendFormat("\r\n                            FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                           WHERE REPLACE(X1.OPN_RMK, ' ', '') <> '' ");
            //strSql.AppendFormat("\r\n                           GROUP BY X1.DEALER_CD, X1.OPN_GB ) Y3 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD   ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB  ");
            //strSql.AppendFormat("\r\n             AND Y3.OPN_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 5050042 #슈레더반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            ////#0010
            //strSql.AppendFormat("\r\n  ORDER BY Z1.WGT DESC, Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");
            #endregion

            #region[2021-04-16 이전 쿼리]

            //strSql.Clear();
            //strSql.AppendFormat("\r\n ");
            //strSql.AppendFormat("\r\n SELECT * ");
            //strSql.AppendFormat("\r\n   FROM ( ");
            //strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            //strSql.AppendFormat("\r\n               , Y1.GB ");
            //strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            //strSql.AppendFormat("\r\n               , Y1.WGT ");
            //strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            //strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n               , '{0}' AS DATE_F ", dicParams["DATE_F"]);
            //strSql.AppendFormat("\r\n            FROM ( ");
            //strSql.AppendFormat("\r\n                   SELECT '2' AS OPN_GB ");
            //strSql.AppendFormat("\r\n                        , '슈레더' AS GB ");
            //strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            //strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            //strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            //strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            //strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            //strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            //strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X2.CHAGAM > 0 THEN 1 ");
            //strSql.AppendFormat("\r\n                                                   WHEN X1.ICHAGAM > 0 THEN 1 END) AS CHAGAM_CNT  ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING X1  ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                                     ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                          FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                                         WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                           AND W.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                                         ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                                         LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                                   LEFT JOIN JAJAE X3 ");
            //strSql.AppendFormat("\r\n                                     ON X1.J_SERIAL = X3.J_SERIAL ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                                    AND X3.DAEGUBUN = '슈레더' ");
            //strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4  ");
            //strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID  ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            //strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            //strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            //strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            //strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            //strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            //strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            //strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                 ) Y1 ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURING_SEQ X2 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = ( SELECT W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                FROM MESURING_SEQ W ");
            //strSql.AppendFormat("\r\n                                               WHERE X2.JUNPYOID = W.JUNPYOID ");
            //strSql.AppendFormat("\r\n                                                 AND W.CHAGAM > 0");
            //strSql.AppendFormat("\r\n                                               ORDER BY W.JUNPYOID, W.CHG_SEQ ");
            //strSql.AppendFormat("\r\n                                               LIMIT 1 ) ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            //strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            //strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            //strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('슈레더') ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            //strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            //strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            //strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            //strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            //strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            //strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            //strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            //strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            //strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            //strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            //strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 5050042 #슈레더반품 ");
            //strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            //strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            //strSql.AppendFormat("\r\n           WHERE Y1.CHAGAM_CNT > 0 ");
            //strSql.AppendFormat("\r\n        ) Z1 ");
            //strSql.AppendFormat("\r\n  ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");

            #endregion[2021-04-16 이전 쿼리]

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            /*
             * #0002
             */

            if (TabControl.SelectedTabPage == TabPageGumsu)
            {
                SaveInspectionInfo();
            }
            else if (TabControl.SelectedTabPage == TabPageSummary)
            {
                SaveSummaryInfo();
            }
        }

        private void SaveInspectionInfo()
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                if (GridViewRetr.RowCount == 0)
                {
                    XtraMessageBox.Show("리스트에 계근자료가 존재하지 않습니다.");
                    return;
                }

                DataTable dtPrv = (DataTable)GridRetr.DataSource;
                DataTable dt = dtPrv.GetChanges(DataRowState.Modified);

                if (dt == null)
                {
                    XtraMessageBox.Show("리스트에 내용을 입력하세요.");
                    return;
                }

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sFIND_JUNPYOID = dt.Rows[0]["JUNPYOID"]?.ToString();

                foreach (DataRow row in dt.Rows)
                {
                    string sJUNPYOID = row["JUNPYOID"].ToString();
                    string sISPT_NO = row["ISPT_NO"].ToString();
                    string sISPT_DT = DateTime.Today.ToString("yyyy-MM-dd");
                    string sJ_BNUM = row["J_BNUM"].ToString();
                    string sDEALER_CD = row["DEALER_CD"].ToString();
                    string sDEALER_NM = row["DEALER_NM"].ToString();
                    string sJ_SERIAL = row["J_SERIAL"].ToString();
                    string sGUBUN1 = row["GUBUN1"].ToString();

                    double dWEIGHT = 0;
                    double.TryParse(row["WEIGHT"].ToString(), out dWEIGHT);
                    string sWEIGHT = dWEIGHT.ToString();

                    string sJ_STATE = row["J_STATE"].ToString();

                    double dCHAGAM = 0;
                    double.TryParse(row["FIRST_CHAGAM"].ToString(), out dCHAGAM);
                    string sCHAGAM = dCHAGAM.ToString();

                    string sITNL_YN = string.IsNullOrEmpty(row["ITNL_YN"]?.ToString()) ? "N" : row["ITNL_YN"]?.ToString();
                    string sISPT_OPN = row["ISPT_OPN"].ToString();

                    double dIMG_CNT = 0;
                    double.TryParse(row["IMG_CNT"].ToString(), out dIMG_CNT);
                    string sIMG_CNT = dIMG_CNT.ToString();

                    //검사번호 채번
                    if (string.IsNullOrEmpty(sISPT_NO))
                    {
                        strSql.Clear();
                        strSql.AppendFormat(" ");
                        #region MariaDB
                        //strSql.AppendFormat(" SELECT LPAD(CAST(IFNULL(MAX(A.ISPT_NO), 0) + 1 AS INT), '7', '0') AS MAX_VAL ");
                        //strSql.AppendFormat("   FROM MESURE_ISPT_INFO A ");
                        //strSql.AppendFormat("  WHERE A.JUNPYOID = {0} ", sJUNPYOID);
                        #endregion

                        strSql.AppendLine("SELECT REPLICATE(0,7 - LEN(ISNULL(MAX(A.ISPT_NO), 0) + 1)) +CONVERT(VARCHAR(20), ISNULL(MAX(A.ISPT_NO), 0) + 1) AS MAX_VAL");
                        strSql.AppendLine("  FROM MESURE_ISPT_INFO A");
                        strSql.AppendLine(" WHERE A.JUNPYOID = "+ sJUNPYOID);
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        sISPT_NO = cmd.ExecuteScalar()?.ToString();
                    }

                    strSql.Clear();
                    #region mariaDB
                    //strSql.AppendFormat(" ");
                    //strSql.AppendFormat("\r\n INSERT INTO MESURE_ISPT_INFO ");
                    //strSql.AppendFormat("\r\n           ( JUNPYOID ");
                    //strSql.AppendFormat("\r\n           , ISPT_NO ");
                    //strSql.AppendFormat("\r\n           , ISPT_DT ");
                    //strSql.AppendFormat("\r\n           , J_BNUM ");
                    //strSql.AppendFormat("\r\n           , DEALER_CD ");
                    //strSql.AppendFormat("\r\n           , DEALER_NM ");
                    //strSql.AppendFormat("\r\n           , J_SERIAL ");
                    //strSql.AppendFormat("\r\n           , GUBUN1 ");
                    //strSql.AppendFormat("\r\n           , WEIGHT ");
                    //strSql.AppendFormat("\r\n           , CHAGAM ");
                    //strSql.AppendFormat("\r\n           , J_STATE ");
                    //strSql.AppendFormat("\r\n           , ITNL_YN ");
                    //strSql.AppendFormat("\r\n           , ISPT_OPN ");
                    //strSql.AppendFormat("\r\n           , ENT_ID ");
                    //strSql.AppendFormat("\r\n           , ENT_DT  ) ");
                    //strSql.AppendFormat("\r\n     VALUES( {0} ", sJUNPYOID); //JUNPYOID
                    //strSql.AppendFormat("\r\n           , '{0}' ", sISPT_NO);  //ISPT_NO
                    //strSql.AppendFormat("\r\n           , '{0}' ", sISPT_DT);  //ISPT_DT
                    //strSql.AppendFormat("\r\n           , '{0}' ", sJ_BNUM);   //J_BNUM
                    //strSql.AppendFormat("\r\n           , {0} ", sDEALER_CD);//DEALER_CD
                    //strSql.AppendFormat("\r\n           , '{0}' ", sDEALER_NM);//DEALER_NM
                    //strSql.AppendFormat("\r\n           , {0} ", sJ_SERIAL); //J_SERIAL
                    //strSql.AppendFormat("\r\n           , '{0}' ", sGUBUN1);   //GUBUN1
                    //strSql.AppendFormat("\r\n           , {0} ", sWEIGHT);   //WEIGHT
                    //strSql.AppendFormat("\r\n           , {0} ", sCHAGAM);   //CHAGAM
                    //strSql.AppendFormat("\r\n           , '{0}' ", sJ_STATE);  //J_STATE
                    //strSql.AppendFormat("\r\n           , '{0}' ", sITNL_YN);  //ITNL_YN
                    //strSql.AppendFormat("\r\n           , '{0}' ", sISPT_OPN); //ISPT_OPN
                    //strSql.AppendFormat("\r\n           , {0} ", FmMainToolBar2.UserID);   //ENT_ID
                    //strSql.AppendFormat("\r\n           , NOW() ) "); //ENT_DT
                    //strSql.AppendFormat("\r\n     ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendFormat("\r\n             ISPT_DT   = '{0}' ", sISPT_DT);
                    //strSql.AppendFormat("\r\n           , J_BNUM    = '{0}' ", sJ_BNUM);
                    //strSql.AppendFormat("\r\n           , DEALER_CD = {0} ", sDEALER_CD);
                    //strSql.AppendFormat("\r\n           , DEALER_NM = '{0}' ", sDEALER_NM);
                    //strSql.AppendFormat("\r\n           , J_SERIAL  = {0} ", sJ_SERIAL);
                    //strSql.AppendFormat("\r\n           , GUBUN1    = '{0}' ", sGUBUN1);
                    //strSql.AppendFormat("\r\n           , WEIGHT    = {0} ", sWEIGHT);
                    //strSql.AppendFormat("\r\n           , CHAGAM    = {0} ", sCHAGAM);
                    //strSql.AppendFormat("\r\n           , J_STATE   = '{0}' ", sJ_STATE);
                    //strSql.AppendFormat("\r\n           , ITNL_YN   = '{0}' ", sITNL_YN);
                    //strSql.AppendFormat("\r\n           , ISPT_OPN  = '{0}' ", sISPT_OPN);
                    //strSql.AppendFormat("\r\n           , MFY_ID    = {0} ", FmMainToolBar2.UserID);
                    //strSql.AppendFormat("\r\n           , MFY_DT    = NOW() ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM MESURE_ISPT_INFO WHERE JUNPYOID = "+ sJUNPYOID + " AND ISPT_NO = '"+ sISPT_NO + "')");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         UPDATE MESURE_ISPT_INFO");
                    strSql.AppendFormat("\r\n         SET ISPT_DT   = '{0}' ", sISPT_DT);
                    strSql.AppendFormat("\r\n           , J_BNUM    = '{0}' ", sJ_BNUM);
                    strSql.AppendFormat("\r\n           , DEALER_CD = {0} ", sDEALER_CD);
                    strSql.AppendFormat("\r\n           , DEALER_NM = '{0}' ", sDEALER_NM);
                    strSql.AppendFormat("\r\n           , J_SERIAL  = {0} ", sJ_SERIAL);
                    strSql.AppendFormat("\r\n           , GUBUN1    = '{0}' ", sGUBUN1);
                    strSql.AppendFormat("\r\n           , WEIGHT    = {0} ", sWEIGHT);
                    strSql.AppendFormat("\r\n           , CHAGAM    = {0} ", sCHAGAM);
                    strSql.AppendFormat("\r\n           , J_STATE   = '{0}' ", sJ_STATE);
                    strSql.AppendFormat("\r\n           , ITNL_YN   = '{0}' ", sITNL_YN);
                    strSql.AppendFormat("\r\n           , ISPT_OPN  = '{0}' ", sISPT_OPN);
                    strSql.AppendFormat("\r\n           , MFY_ID    = {0} ", FmMainToolBar2.UserID);
                    strSql.AppendFormat("\r\n           , MFY_DT    = CONVERT(VARCHAR(20),GETDATE(),20) ");
                    strSql.AppendLine("             WHERE JUNPYOID = " + sJUNPYOID + " AND ISPT_NO = '" + sISPT_NO + "'");
                    strSql.AppendLine("   END");
                    strSql.AppendLine("ELSE");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendFormat("\r\n INSERT INTO MESURE_ISPT_INFO ");
                    strSql.AppendFormat("\r\n           ( JUNPYOID ");
                    strSql.AppendFormat("\r\n           , ISPT_NO ");
                    strSql.AppendFormat("\r\n           , ISPT_DT ");
                    strSql.AppendFormat("\r\n           , J_BNUM ");
                    strSql.AppendFormat("\r\n           , DEALER_CD ");
                    strSql.AppendFormat("\r\n           , DEALER_NM ");
                    strSql.AppendFormat("\r\n           , J_SERIAL ");
                    strSql.AppendFormat("\r\n           , GUBUN1 ");
                    strSql.AppendFormat("\r\n           , WEIGHT ");
                    strSql.AppendFormat("\r\n           , CHAGAM ");
                    strSql.AppendFormat("\r\n           , J_STATE ");
                    strSql.AppendFormat("\r\n           , ITNL_YN ");
                    strSql.AppendFormat("\r\n           , ISPT_OPN ");
                    strSql.AppendFormat("\r\n           , ENT_ID ");
                    strSql.AppendFormat("\r\n           , ENT_DT  ) ");
                    strSql.AppendFormat("\r\n     VALUES( {0} ", sJUNPYOID); //JUNPYOID
                    strSql.AppendFormat("\r\n           , '{0}' ", sISPT_NO);  //ISPT_NO
                    strSql.AppendFormat("\r\n           , '{0}' ", sISPT_DT);  //ISPT_DT
                    strSql.AppendFormat("\r\n           , '{0}' ", sJ_BNUM);   //J_BNUM
                    strSql.AppendFormat("\r\n           , {0} ", sDEALER_CD);//DEALER_CD
                    strSql.AppendFormat("\r\n           , '{0}' ", sDEALER_NM);//DEALER_NM
                    strSql.AppendFormat("\r\n           , {0} ", sJ_SERIAL); //J_SERIAL
                    strSql.AppendFormat("\r\n           , '{0}' ", sGUBUN1);   //GUBUN1
                    strSql.AppendFormat("\r\n           , {0} ", sWEIGHT);   //WEIGHT
                    strSql.AppendFormat("\r\n           , {0} ", sCHAGAM);   //CHAGAM
                    strSql.AppendFormat("\r\n           , '{0}' ", sJ_STATE);  //J_STATE
                    strSql.AppendFormat("\r\n           , '{0}' ", sITNL_YN);  //ITNL_YN
                    strSql.AppendFormat("\r\n           , '{0}' ", sISPT_OPN); //ISPT_OPN
                    strSql.AppendFormat("\r\n           , {0} ", FmMainToolBar2.UserID);   //ENT_ID
                    strSql.AppendFormat("\r\n           , CONVERT(VARCHAR(20),GETDATE(),20) ) "); //ENT_DT
                    strSql.AppendLine("END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridCol1JunpyoID, sFIND_JUNPYOID);

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SaveSummaryInfo()
        {
            if (TabControl.SelectedTabPage != TabPageSummary)
                return;

            try
            {
                Cursor = Cursors.WaitCursor;

                DataTable dtPrv = (DataTable)GridRetr2.DataSource;
                DataTable dt = dtPrv.GetChanges(DataRowState.Modified);

                if (dt == null)
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("리스트에 총평을 입력하세요.");
                    return;
                }

                string sDealerCd = string.Empty;
                string sOpnGb = string.Empty;
                //bool bYn = false;
                //foreach(DataRow row in dt.Rows)
                //{
                //    sDealerCd = row["MAIPCHERID"].ToString();
                //    sOpnGb = row["OPN_GB"].ToString();
                //    string sRMK = row["OPN_RMK"].ToString();

                //    if (string.IsNullOrEmpty(sRMK))
                //    {
                //        bYn = true;
                //        Cursor = Cursors.Default;
                //        XtraMessageBox.Show("총평을 입력하세요.");
                //    }
                //}

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                int iCnt = 0;
                foreach (DataRow row in dt.Rows)
                {
                    string sDEALER_CD = row["MAIPCHERID"].ToString();
                    string sOPN_GB = row["OPN_GB"].ToString();
                    string sOPN_DATE = row["OPN_DATE"].ToString();
                    string sDATE_T = row["DATE_T"].ToString();
                    string sOPN_RMK = row["OPN_RMK"].ToString();

                    //sOPN_DATE = string.IsNullOrEmpty(sOPN_DATE) ? sDATE_T : sOPN_DATE;
                    sOPN_DATE = sDATE_T;

                    if (iCnt++ == 0)
                    {
                        sDealerCd = sDEALER_CD;
                        sOpnGb = sOPN_GB;
                    }

                    strSql.Clear();
                    #region MariaDB
                    //strSql.AppendFormat(" ");
                    //strSql.AppendFormat(" INSERT INTO MESURE_OPN_HISTORY ");
                    //strSql.AppendFormat("           ( DEALER_CD ");
                    //strSql.AppendFormat("           , OPN_GB ");
                    //strSql.AppendFormat("           , OPN_DATE ");
                    //strSql.AppendFormat("           , OPN_RMK ");
                    //strSql.AppendFormat("           , ENT_ID ");
                    //strSql.AppendFormat("           , ENT_DT ) ");
                    //strSql.AppendFormat("     VALUES( {0} ", sDEALER_CD); //DEALER_CD
                    //strSql.AppendFormat("           , '{0}' ", sOPN_GB); //OPN_GB
                    //strSql.AppendFormat("           , '{0}' ", sOPN_DATE); //OPN_DATE
                    //strSql.AppendFormat("           , '{0}' ", sOPN_RMK); //OPN_RMK
                    //strSql.AppendFormat("           , {0} ", FmMainToolBar2.UserID); //ENT_ID
                    //strSql.AppendFormat("           , NOW() ) "); //ENT_DT
                    //strSql.AppendFormat("        ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendFormat("           OPN_RMK = '{0}' ", sOPN_RMK); //OPN_RMK
                    //strSql.AppendFormat("         , MFY_ID = {0} ", FmMainToolBar2.UserID); //MFY_ID
                    //strSql.AppendFormat("         , MFY_DT = NOW() "); //MFY_DT
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM MESURE_OPN_HISTORY WHERE DEALER_CD = "+ sDEALER_CD + " AND OPN_GB = '"+ sOPN_GB + "' AND OPN_DATE = '"+ sOPN_DATE + "')");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         UPDATE MESURE_OPN_HISTORY");
                    strSql.AppendLine("            SET OPN_RMK = '"+ sOPN_RMK + "'");
                    strSql.AppendLine("              , MFY_ID = "+ FmMainToolBar2.UserID);
                    strSql.AppendLine("              , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20)");
                    strSql.AppendLine("          WHERE DEALER_CD = " + sDEALER_CD + " AND OPN_GB = '" + sOPN_GB + "' AND OPN_DATE = '" + sOPN_DATE + "'");
                    strSql.AppendLine("   END");
                    strSql.AppendLine("ELSE");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("       INSERT INTO MESURE_OPN_HISTORY");
                    strSql.AppendLine("            ( DEALER_CD");
                    strSql.AppendLine("            , OPN_GB");
                    strSql.AppendLine("            , OPN_DATE");
                    strSql.AppendLine("            , OPN_RMK");
                    strSql.AppendLine("            , ENT_ID");
                    strSql.AppendLine("            , ENT_DT )");
                    strSql.AppendLine("      VALUES( "+ sDEALER_CD);
                    strSql.AppendLine("            , '"+ sOPN_GB + "'");
                    strSql.AppendLine("            , '"+ sOPN_DATE + "'");
                    strSql.AppendLine("            , '"+ sOPN_RMK + "'");
                    strSql.AppendLine("            , "+ FmMainToolBar2.UserID);
                    strSql.AppendLine("            , CONVERT(VARCHAR(20), GETDATE(), 20))");
                    strSql.AppendLine("   END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewRetr2.FocusedRowHandle = GridViewRetr2.LocateByDisplayText(0, GridCol2OpnGb, sOpnGb);
                GridViewRetr2.FocusedRowHandle = GridViewRetr2.LocateByDisplayText(GridViewRetr2.FocusedRowHandle, GridCol2MaipcherID, sDealerCd);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
            {
                if (GridImg.DataSource != null)
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


            string sENT_DT = GridViewRetr.GetFocusedRowCellValue(GridColEntDt)?.ToString().Replace("-", "");
            string sJUNPYOID = GridViewRetr.GetFocusedRowCellValue(GridCol1JunpyoID)?.ToString();

            GetImagesFromFTP(sENT_DT, sJUNPYOID);

        }

        private void GetImagesFromFTP(string sEntDt, string sJunpyoID)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                //string sInitDir = string.Format(@"ftp://192.168.0.202/Gumsu_Images/{0}/{1}/{2}/{3}", sEntDt.Substring(0, 4), sEntDt.Substring(4, 2), sEntDt.Substring(6, 2), sJunpyoID);
                string sInitDir = string.Format(@"ftp://{0}/Gumsu_Images/{1}/{2}/{3}/{4}", ComnEtcFunc.FTP_IP, sEntDt.Substring(0, 4), sEntDt.Substring(4, 2), sEntDt.Substring(6, 2), sJunpyoID);

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

        private void PD04001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            //else if (e.KeyCode == Keys.F8)
            //    BtnExcel.PerformClick();
            else if (e.KeyCode == Keys.F8)
                DropBtnExcel.PerformClick();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewSummary_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            string sCUser = GridViewRetr.GetRowCellValue(e.RowHandle, GridCol1CUser)?.ToString();
            if (!string.IsNullOrEmpty(sCUser))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else
            {
                e.Appearance.BackColor = Color.Yellow;
            }

            //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewSummary_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RepoTxt1FirstChagam_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridCol1FirstChagam, txt.EditValue);
        }

        private void RepoChk1ItmlYn_EditValueChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridCol1ItnlYn, chk.EditValue);
        }

        private void RepoTxt1IsptOpn_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridCol1IsptOpn, txt.EditValue);
        }

        private void ChkSelf_CheckedChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewRetr_ShowingEditor(object sender, CancelEventArgs e)
        {
            string sUser = GridViewRetr.GetFocusedRowCellValue(GridCol1CUser)?.ToString() ?? string.Empty;
            if (!string.IsNullOrEmpty(sUser) && !sUser.Equals(FmMainToolBar2.UserID))
            {
                e.Cancel = true;
            }
        }

        private void PD04001F01_TextChanged(object sender, EventArgs e)
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

        private void RepoBtnEdit1ImgCnt_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            string sJunpyoId = GridViewRetr.GetFocusedRowCellValue(GridCol1JunpyoID)?.ToString();
            string sIsptNo = GridViewRetr.GetFocusedRowCellValue(GridCol1IsptNo)?.ToString();
            if (string.IsNullOrEmpty(sIsptNo))
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("저장버튼을 클릭 한 후 이미지를 등록하세요.");
                return;
            }

            PD04001F02 frm = new PD04001F02();

            frm._JUNPYOID = sJunpyoId;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridCol1JunpyoID, sJunpyoId);
            }
        }

        /*
         * #0002
         */

        private void TabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (TabControl.SelectedTabPage == TabPageGumsu)
            {
                LayoutSelf.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutItnl.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                SetComboboxItem(_ListWordByGumsu);
            }
            else
            {
                LayoutSelf.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutItnl.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                SetComboboxItem(_ListWordBySummary);
            }

            BtnRetr.PerformClick();
        }

        private void SetComboboxItem(List<string> list)
        {
            CboFindIdx.Properties.Items.Clear();
            TxtFindWord.EditValue = string.Empty;
            foreach (string item in list)
            {
                CboFindIdx.Properties.Items.Add(item);
            }

            if (CboFindIdx.Properties.Items.Count > 0)
                CboFindIdx.SelectedIndex = 0;
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (TabControl.SelectedTabPage == TabPageGumsu)
            {
                ComnEtcFunc.ExportExcelFile("검수내역리스트", GridRetr);
            }
            else
            {
                ComnEtcFunc.ExportExcelFile("검수내역합계", GridRetr2);
            }
        }

        private void GridViewRetr2_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RdgbItnlYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void DropBtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "주간보고서")
            {
                GetWeeklyReport();
            }
            else if (tag == "월간보고서")
            {
                GetMonthlyReport();
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnWeek;
        BarButtonItem BtnMonth;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnWeek = new BarButtonItem(barManager1, "주간보고서");
            BtnMonth = new BarButtonItem(barManager1, "월간보고서");
            popupMenu1.AddItem(BtnWeek);
            popupMenu1.AddItem(BtnMonth);

            DropBtnExcel.DropDownControl = popupMenu1;

            BtnWeek.Tag = "주간보고서";
            BtnWeek.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnWeek_ItemClick);

            BtnMonth.Tag = "월간보고서";
            BtnMonth.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnMonth_ItemClick);
        }

        private void BtnWeek_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            GetWeeklyReport();
        }

        private void BtnMonth_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            GetMonthlyReport();
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnExcel.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnExcel.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnExcel.Tag = submenuItem.Tag;
            //DropBtnExcel.Text = string.Format("{0}", submenuItem.Tag);
        }

        /*
         * #0003
         */
        private void GetWeeklyReport()
        {
            FileInfo_1 fileInfo = new FileInfo_1("1");

            Cursor = Cursors.WaitCursor;
            string[] sPath = fileInfo.CheckFileInfo();
            Cursor = Cursors.Default;

            if (sPath != null)
            {
                SetWeeklyReportForm(sPath[0], sPath[1]);
                Process.Start(sPath[1]);
            }
        }

        private void GetMonthlyReport()
        {
            FileInfo_1 fileInfo = new FileInfo_1("2");

            Cursor = Cursors.WaitCursor;
            string[] sPath = fileInfo.CheckFileInfo();
            Cursor = Cursors.Default;

            if (sPath != null)
            {
                SetMonthlyReportForm(sPath[0], sPath[1]);
                Process.Start(sPath[1]);
            }
        }

        Excel.Application ExcelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        Excel.Workbook wbk = null;
        Excel.Worksheet ws1k = null;
        Excel.Worksheet ws2k = null;
        private void SetWeeklyReportForm(string StandardPath, string SavePath)
        {
            try
            {
                if (!File.Exists(StandardPath))
                {
                    XtraMessageBox.Show("엑셀파일 양식이 존재하지 않습니다.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                ExcelApp = new Excel.Application();
                wb = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                ws = wb.Worksheets["주간집계"] as Excel.Worksheet;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);

                DataTable dtScrap = GetSummaryScrap(dicParams);
                DataTable dtShreder = GetSummaryShreder(dicParams);
                if (dtScrap.Rows.Count == 0 && dtShreder.Rows.Count == 0)
                {
                    string sMSG = string.Format("{0}~{1}까지의 스크랩/슈레더의 결과가 존재하지 않습니다.", sYmdFrom, sYmdTo);
                    return;
                }

                //일자부분 세팅
                ws.Range["F6"].Value = sYmdFrom;
                ws.Range["L6"].Value = sYmdTo;

                DateTime dtStartTime = DateTime.Parse(sYmdFrom);
                DateTime calculationDate = new DateTime(dtStartTime.Year, dtStartTime.Month, dtStartTime.Day);   //주차를 구할 일자
                DateTime calculationDate1 = new DateTime(dtStartTime.Year, dtStartTime.Month, 1); //기준일
                Calendar calenderCalc = CultureInfo.CurrentCulture.Calendar;

                //DayOfWeek.Sunday 인수는 기준 요일
                int usWeekNumber = calenderCalc.GetWeekOfYear(calculationDate, CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - calenderCalc.GetWeekOfYear(calculationDate1, CalendarWeekRule.FirstDay, DayOfWeek.Sunday) + 1;

                string sWeek = string.Format("{0}월 {1}주차", dtStartTime.Month.ToString().PadLeft(2, '0'), usWeekNumber);
                ws.Range["Q6"].Value = sWeek;

                //스크랩 21건이 최대
                int iRowCount = dtScrap.Rows.Count > 21 ? 21 : dtScrap.Rows.Count;
                int iStartRow = 7;
                for (int i = 0; i < iRowCount; i++)
                {
                    string sDEALER_NM = dtScrap.Rows[i]["DEALER_NM"]?.ToString(); //업체명
                    int iWGT = 0; //입고중량
                    int.TryParse(dtScrap.Rows[i]["WGT"]?.ToString(), out iWGT);
                    int iSUM_CHAGAM = 0; //감량중량
                    int.TryParse(dtScrap.Rows[i]["SUM_CHAGAM"]?.ToString(), out iSUM_CHAGAM);
                    int iIN_CNT = 0; //입고횟수
                    int.TryParse(dtScrap.Rows[i]["IN_CNT"]?.ToString(), out iIN_CNT);
                    int iRETURN_CNT = 0; //반품횟수
                    int.TryParse(dtScrap.Rows[i]["RETURN_CNT"]?.ToString(), out iRETURN_CNT);
                    int iCHAGAM_CNT = 0; //감량횟수
                    int.TryParse(dtScrap.Rows[i]["CHAGAM_CNT"]?.ToString(), out iCHAGAM_CNT);
                    int iCHAGAM_WGT = 0; //감량조정
                    int.TryParse(dtScrap.Rows[i]["CHAGAM_WGT"]?.ToString(), out iCHAGAM_WGT);
                    string sITNL_YN = dtScrap.Rows[i]["ITNL_YN"]?.ToString(); //고의유무
                    string sOPN_RMK = dtScrap.Rows[i]["OPN_RMK"]?.ToString(); //검수소견

                    int iApplyRowIdx = iStartRow + (i + 1);


                    ws.Range["U" + iApplyRowIdx].Font.Size = 7;
                    ws.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                    ws.Range["J" + iApplyRowIdx].Value = iWGT;
                    ws.Range["M" + iApplyRowIdx].Value = iSUM_CHAGAM;
                    ws.Range["O" + iApplyRowIdx].Value = iIN_CNT;
                    ws.Range["Q" + iApplyRowIdx].Value = iRETURN_CNT;
                    ws.Range["S" + iApplyRowIdx].Value = iCHAGAM_CNT;
                    ws.Range["U" + iApplyRowIdx].Value = (iSUM_CHAGAM - iCHAGAM_WGT);
                    SetRangeFont(ws.Range["U" + iApplyRowIdx], (iSUM_CHAGAM - iCHAGAM_WGT).ToString().Length);
                    ws.Range["W" + iApplyRowIdx].Value = iSUM_CHAGAM - (iSUM_CHAGAM - iCHAGAM_WGT);
                    ws.Range["X" + iApplyRowIdx].Value = sITNL_YN;
                    ws.Range["Z" + iApplyRowIdx].Value = sOPN_RMK;
                }


                //슈레더 15건이 최대
                iRowCount = dtShreder.Rows.Count > 21 ? 21 : dtShreder.Rows.Count;
                iStartRow = 32;
                for (int i = 0; i < iRowCount; i++)
                {
                    string sDEALER_NM = dtShreder.Rows[i]["DEALER_NM"]?.ToString(); //업체명
                    int iWGT = 0; //입고중량
                    int.TryParse(dtShreder.Rows[i]["WGT"]?.ToString(), out iWGT);
                    int iSUM_CHAGAM = 0; //감량중량
                    int.TryParse(dtShreder.Rows[i]["SUM_CHAGAM"]?.ToString(), out iSUM_CHAGAM);
                    int iIN_CNT = 0; //입고횟수
                    int.TryParse(dtShreder.Rows[i]["IN_CNT"]?.ToString(), out iIN_CNT);
                    int iRETURN_CNT = 0; //반품횟수
                    int.TryParse(dtShreder.Rows[i]["RETURN_CNT"]?.ToString(), out iRETURN_CNT);
                    int iCHAGAM_CNT = 0; //감량횟수
                    int.TryParse(dtShreder.Rows[i]["CHAGAM_CNT"]?.ToString(), out iCHAGAM_CNT);
                    int iCHAGAM_WGT = 0; //감량조정
                    int.TryParse(dtShreder.Rows[i]["CHAGAM_WGT"]?.ToString(), out iCHAGAM_WGT);
                    string sITNL_YN = dtShreder.Rows[i]["ITNL_YN"]?.ToString(); //고의유무
                    string sOPN_RMK = dtShreder.Rows[i]["OPN_RMK"]?.ToString(); //검수소견

                    int iApplyRowIdx = iStartRow + (i + 1);


                    ws.Range["U" + iApplyRowIdx].Font.Size = 7;
                    ws.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                    ws.Range["J" + iApplyRowIdx].Value = iWGT;
                    ws.Range["M" + iApplyRowIdx].Value = iSUM_CHAGAM;
                    ws.Range["O" + iApplyRowIdx].Value = iIN_CNT;
                    ws.Range["Q" + iApplyRowIdx].Value = iRETURN_CNT;
                    ws.Range["S" + iApplyRowIdx].Value = iCHAGAM_CNT;
                    ws.Range["U" + iApplyRowIdx].Value = (iSUM_CHAGAM - iCHAGAM_WGT);
                    SetRangeFont(ws.Range["U" + iApplyRowIdx], (iSUM_CHAGAM - iCHAGAM_WGT).ToString().Length);
                    ws.Range["W" + iApplyRowIdx].Value = iSUM_CHAGAM - (iSUM_CHAGAM - iCHAGAM_WGT);
                    ws.Range["X" + iApplyRowIdx].Value = sITNL_YN;
                    ws.Range["Z" + iApplyRowIdx].Value = sOPN_RMK;
                }


                if (File.Exists(SavePath))
                    File.Delete(SavePath);

                Cursor = Cursors.Default;
                //wb.SaveAs(SavePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //파일 닫기... 
                wb.SaveAs(SavePath);
                wb.Close(false, Type.Missing, Type.Missing);
                wb = null;
                ExcelApp.Quit();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
            finally
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
        }

        private void SetMonthlyReportForm(string StandardPath, string SavePath)
        {
            try
            {
                if (!File.Exists(StandardPath))
                {
                    XtraMessageBox.Show("엑셀파일 양식이 존재하지 않습니다.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                ExcelApp = new Excel.Application();
                wbk = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                ws1k = wbk.Worksheets["월간(스크랩)"] as Excel.Worksheet;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);

                DataTable dtScrap = GetSummaryScrap(dicParams);
                DataTable dtShreder = GetSummaryShreder(dicParams);
                if (dtScrap.Rows.Count == 0 && dtShreder.Rows.Count == 0)
                {
                    string sMSG = string.Format("{0}~{1}까지의 스크랩/슈레더의 결과가 존재하지 않습니다.", sYmdFrom, sYmdTo);
                    return;
                }

                //일자부분 세팅
                ws1k.Range["F6"].Value = sYmdFrom;
                ws1k.Range["L6"].Value = sYmdTo;

                DateTime dtStartTime = DateTime.Parse(sYmdFrom);
                DateTime calculationDate = new DateTime(dtStartTime.Year, dtStartTime.Month, dtStartTime.Day);   //주차를 구할 일자
                DateTime calculationDate1 = new DateTime(dtStartTime.Year, dtStartTime.Month, 1); //기준일
                Calendar calenderCalc = CultureInfo.CurrentCulture.Calendar;

                //DayOfWeek.Sunday 인수는 기준 요일
                int usWeekNumber = calenderCalc.GetWeekOfYear(calculationDate, CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - calenderCalc.GetWeekOfYear(calculationDate1, CalendarWeekRule.FirstDay, DayOfWeek.Sunday) + 1;

                //스크랩 21건이 최대
                int iRowCount = dtScrap.Rows.Count > 100 ? 100 : dtScrap.Rows.Count;
                int iStartRow = 7;
                for (int i = 0; i < iRowCount; i++)
                {
                    string sDEALER_NM = dtScrap.Rows[i]["DEALER_NM"]?.ToString(); //업체명
                    int iWGT = 0; //입고중량
                    int.TryParse(dtScrap.Rows[i]["WGT"]?.ToString(), out iWGT);
                    int iSUM_CHAGAM = 0; //감량중량
                    int.TryParse(dtScrap.Rows[i]["SUM_CHAGAM"]?.ToString(), out iSUM_CHAGAM);
                    int iIN_CNT = 0; //입고횟수
                    int.TryParse(dtScrap.Rows[i]["IN_CNT"]?.ToString(), out iIN_CNT);
                    int iRETURN_CNT = 0; //반품횟수
                    int.TryParse(dtScrap.Rows[i]["RETURN_CNT"]?.ToString(), out iRETURN_CNT);
                    int iCHAGAM_CNT = 0; //감량횟수
                    int.TryParse(dtScrap.Rows[i]["CHAGAM_CNT"]?.ToString(), out iCHAGAM_CNT);
                    int iCHAGAM_WGT = 0; //현재감량
                    int.TryParse(dtScrap.Rows[i]["CHAGAM_WGT"]?.ToString(), out iCHAGAM_WGT);
                    string sITNL_YN = dtScrap.Rows[i]["ITNL_YN"]?.ToString(); //고의유무
                    string sOPN_RMK = dtScrap.Rows[i]["OPN_RMK"]?.ToString(); //검수소견

                    int iApplyRowIdx = iStartRow + (i + 1);


                    ws1k.Range["U" + iApplyRowIdx].Font.Size = 7;
                    ws1k.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                    ws1k.Range["J" + iApplyRowIdx].Value = iWGT;
                    ws1k.Range["M" + iApplyRowIdx].Value = iSUM_CHAGAM;
                    ws1k.Range["O" + iApplyRowIdx].Value = iIN_CNT;
                    ws1k.Range["Q" + iApplyRowIdx].Value = iRETURN_CNT;
                    ws1k.Range["S" + iApplyRowIdx].Value = iCHAGAM_CNT;
                    ws1k.Range["U" + iApplyRowIdx].Value = (iSUM_CHAGAM - iCHAGAM_WGT);
                    SetRangeFont(ws1k.Range["U" + iApplyRowIdx], (iSUM_CHAGAM - iCHAGAM_WGT).ToString().Length);
                    ws1k.Range["W" + iApplyRowIdx].Value = iCHAGAM_WGT;
                    ws1k.Range["X" + iApplyRowIdx].Value = sITNL_YN;
                    ws1k.Range["Z" + iApplyRowIdx].Value = sOPN_RMK;
                }

                ws2k = wbk.Worksheets["월간(슈레더)"] as Excel.Worksheet;
                //일자부분 세팅
                ws2k.Range["F6"].Value = sYmdFrom;
                ws2k.Range["L6"].Value = sYmdTo;

                //슈레더 15건이 최대
                iRowCount = dtShreder.Rows.Count;
                iStartRow = 7;
                for (int i = 0; i < iRowCount; i++)
                {
                    string sDEALER_NM = dtShreder.Rows[i]["DEALER_NM"]?.ToString(); //업체명
                    int iWGT = 0; //입고중량
                    int.TryParse(dtShreder.Rows[i]["WGT"]?.ToString(), out iWGT);
                    int iSUM_CHAGAM = 0; //감량중량
                    int.TryParse(dtShreder.Rows[i]["SUM_CHAGAM"]?.ToString(), out iSUM_CHAGAM);
                    int iIN_CNT = 0; //입고횟수
                    int.TryParse(dtShreder.Rows[i]["IN_CNT"]?.ToString(), out iIN_CNT);
                    int iRETURN_CNT = 0; //반품횟수
                    int.TryParse(dtShreder.Rows[i]["RETURN_CNT"]?.ToString(), out iRETURN_CNT);
                    int iCHAGAM_CNT = 0; //감량횟수
                    int.TryParse(dtShreder.Rows[i]["CHAGAM_CNT"]?.ToString(), out iCHAGAM_CNT);
                    int iCHAGAM_WGT = 0; //감량조정
                    int.TryParse(dtShreder.Rows[i]["CHAGAM_WGT"]?.ToString(), out iCHAGAM_WGT);
                    string sITNL_YN = dtShreder.Rows[i]["ITNL_YN"]?.ToString(); //고의유무
                    string sOPN_RMK = dtShreder.Rows[i]["OPN_RMK"]?.ToString(); //검수소견

                    int iApplyRowIdx = iStartRow + (i + 1);


                    ws2k.Range["U" + iApplyRowIdx].Font.Size = 7;
                    ws2k.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                    ws2k.Range["J" + iApplyRowIdx].Value = iWGT;
                    ws2k.Range["M" + iApplyRowIdx].Value = iSUM_CHAGAM;
                    ws2k.Range["O" + iApplyRowIdx].Value = iIN_CNT;
                    ws2k.Range["Q" + iApplyRowIdx].Value = iRETURN_CNT;
                    ws2k.Range["S" + iApplyRowIdx].Value = iCHAGAM_CNT;
                    ws2k.Range["U" + iApplyRowIdx].Value = (iSUM_CHAGAM - iCHAGAM_WGT); ;
                    SetRangeFont(ws2k.Range["U" + iApplyRowIdx], (iSUM_CHAGAM - iCHAGAM_WGT).ToString().Length);
                    ws2k.Range["W" + iApplyRowIdx].Value = iSUM_CHAGAM - (iSUM_CHAGAM - iCHAGAM_WGT);
                    ws2k.Range["X" + iApplyRowIdx].Value = sITNL_YN;
                    ws2k.Range["Z" + iApplyRowIdx].Value = sOPN_RMK;
                }



                if (File.Exists(SavePath))
                    File.Delete(SavePath);

                Cursor = Cursors.Default;
                //wb.SaveAs(SavePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //파일 닫기... 
                wbk.SaveAs(SavePath);
                wbk.Close(false, Type.Missing, Type.Missing);
                wbk = null;
                ExcelApp.Quit();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(wbk);
                ReleaseExcelObject(ws1k);
                ReleaseExcelObject(ws2k);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
            finally
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(wbk);
                ReleaseExcelObject(ws1k);
                ReleaseExcelObject(ws2k);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
        }

        private void SetRangeFont(Excel.Range rg, int Length)
        {
            if (Length <= 3)
            {
                rg.Font.Size = 8;
            }
            else if (Length == 4)
            {
                rg.Font.Size = 7;
            }
            else if (Length > 4)
            {
                rg.Font.Size = 6;
            }
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        BarManager barManager2;
        PopupMenu popupMenu2;
        BarButtonItem BtnDelImg;
        private void InitControlsTileView()
        {
            barManager2 = new BarManager();
            barManager2.Form = this;

            popupMenu2 = new PopupMenu(barManager2);
            BtnDelImg = new BarButtonItem(barManager2, "이미지삭제");
            popupMenu2.AddItem(BtnDelImg);

            BtnDelImg.Tag = "이미지삭제";
            //BtnDelImg.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnWeek_ItemClick);

            barManager2.ItemClick += barManager2_ItemClick;
        }

        private void barManager2_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (e.Item.Tag.Equals("이미지삭제"))
            {
                DeleteImageFromFTP();
            }
        }

        private string _FILE_NAME = string.Empty;
        private DataRow _MASTER_ROW;
        private void TileViewImg_ItemRightClick(object sender, DevExpress.XtraGrid.Views.Tile.TileViewItemClickEventArgs e)
        {
            _FILE_NAME = TileViewImg.GetRowCellValue(e.Item.RowHandle, TileColName)?.ToString();
            _MASTER_ROW = GridViewRetr.GetFocusedDataRow();

            popupMenu2.ShowPopup(Control.MousePosition);
        }

        /*
         * #0005
         */
        private void DeleteImageFromFTP()
        {
            /*
             * #0009
             */
            string sENT_ID = GridViewRetr.GetFocusedRowCellValue(GridCol1CUser)?.ToString() ?? string.Empty;
            string sENT_NM = GridViewRetr.GetFocusedRowCellValue(GridCol1CUserNm)?.ToString() ?? string.Empty;
            string sEXS_ID = FmMainToolBar2.UserID ?? string.Empty;
            string sEXS_NM = FmMainToolBar2.drUser["USRNM"]?.ToString() ?? string.Empty;

            if (!sEXS_ID.Equals(sENT_ID))
            {
                string sMSG = string.Format("현재사용자 : {0}" +
                    "\r\n등록자 : {1}" +
                    "\r\n해당 사용자는 등록자가 아니므로 이미지를 삭제할 수 없습니다."
                    , sEXS_NM
                    , sENT_NM);
                XtraMessageBox.Show(sMSG);
                return;
            }

            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            try
            {
                if (_MASTER_ROW == null || string.IsNullOrEmpty(_FILE_NAME))
                {
                    XtraMessageBox.Show("올바른 데이터를 클릭하세요.");
                    return;
                }

                if (XtraMessageBox.Show("해당 이미지를 정말 삭제하시겠습니까?", "검수이미지 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    return;

                Cursor = Cursors.WaitCursor;

                DateTime dtJ_DATE = DateTime.Parse(_MASTER_ROW["ENT_DT"].ToString());
                string sYear = dtJ_DATE.Year.ToString();
                string sMonth = dtJ_DATE.Month.ToString().PadLeft(2, '0');
                string sDay = dtJ_DATE.Day.ToString().PadLeft(2, '0');

                //string ftpPath = @"ftp://192.168.0.202/Gumsu_Images/" + sYear + "/" + sMonth + "/" + sDay + "/" + _MASTER_ROW["JUNPYOID"] + "/" + _FILE_NAME;
                string ftpPath = @"ftp://"+ ComnEtcFunc.FTP_IP + "/Gumsu_Images/" + sYear + "/" + sMonth + "/" + sDay + "/" + _MASTER_ROW["JUNPYOID"] + "/" + _FILE_NAME;
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;

                bool bResult = ComnEtcFunc.DeleteFTPFile(ftpPath, user, pw);
                if (bResult)
                {
                    DecreateImageCount(_MASTER_ROW["JUNPYOID"].ToString(), _MASTER_ROW["ISPT_NO"].ToString());
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("삭제가 완료되었습니다.");

                    string sENT_DT = GridViewRetr.GetFocusedRowCellValue(GridColEntDt)?.ToString().Replace("-", "");
                    string sJUNPYOID = GridViewRetr.GetFocusedRowCellValue(GridCol1JunpyoID)?.ToString();

                    GetImagesFromFTP(sENT_DT, sJUNPYOID);
                }
                else
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("삭제에 실패하였습니다.");
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("[DeleteImageFromFTP]" + ex.Message);
            }
        }

        private void DecreateImageCount(string JunpyoID, string IsptNo)
        {
            try
            {
                int iImgCnt = 0;
                int.TryParse(GridViewRetr.GetFocusedRowCellValue(GridCol1ImgCnt).ToString(), out iImgCnt);

                GridViewRetr.SetFocusedRowCellValue(GridCol1ImgCnt, --iImgCnt);
                
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendFormat("\r\n ");
                strSql.AppendFormat("\r\n UPDATE MESURE_ISPT_INFO ");
                strSql.AppendFormat("\r\n    SET IMG_CNT = {0} ", iImgCnt);
                strSql.AppendFormat("\r\n  WHERE JUNPYOID = {0} ", JunpyoID);
                strSql.AppendFormat("\r\n    AND ISPT_NO = {0} ", IsptNo);
                
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void TileViewImg_ItemDoubleClick(object sender, DevExpress.XtraGrid.Views.Tile.TileViewItemClickEventArgs e)
        {
            SendImage();
        }

        private void SendImage()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                byte[] byArr2 = (byte[])TileViewImg.GetFocusedRowCellValue(TileColImg);

                byte[] byArr = ObjectToByteArray(TileViewImg.GetFocusedRowCellValue(TileColImg));

                Image img = null;

                ImageConverter imgCvt = new ImageConverter();
                img = (Image)imgCvt.ConvertFrom(byArr2);

                //Image x = (Bitmap)((new ImageConverter()).ConvertFrom(byArr));

                //BinaryFormatter bf = new BinaryFormatter();
                //using (MemoryStream ms = new MemoryStream())
                //{
                //    bf.Serialize(ms, TileViewImg.GetFocusedRowCellValue(TileColImg));
                //    img = Image.FromStream(ms);
                //}

                //using (var ms = new MemoryStream(byArr))
                //{
                //    img = Image.FromStream(ms);
                //}

                Cursor = Cursors.Default;
                PD04001F03 frm = new PD04001F03();

                frm._IMAGE = img;
                frm.Show();
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("[SendImage]" + ex.Message);
            }
            
        }

        private byte[] ObjectToByteArray(object obj)
        {
            if (obj == null)
                return null;
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream())
            {
                bf.Serialize(ms, obj);
                return ms.ToArray();
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}