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
using DevExpress.XtraBars;
using System.IO;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-04
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            입출금 출력 시 Focusing된 Row 건별로 출력하도록 수정
 *            메소드 Print 참조
 *            
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2021-04-20
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 잔액 안맞는 현상 쿼리수정
 *            
 * 수정일자 : 2021-06-07
 * 수정자   : 정은영
 * ID       : #001
 * 수정내용 : (현업요청)
 *            1. 잔액 안맞는 현상 쿼리수정           
 *            
 * 수정일자 : 2022-12-14
 * 수정자   : 정은영
 * ID       : #002
 * 수정내용 : 1. 잔액 안맞는 부분 수정
 */
namespace AccAdm
{
    public partial class AC06001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC06001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC06001F01_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            //DateEditFrom.EditValue = today.AddDays(1 - today.Day);
            DateEditFrom.EditValue = today;
            DateEditTo.EditValue = today;
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            InitControls();
            UpdateDropDownButton(BtnSlipOut);
            arrGrdView = new GridView[] { GridViewRetr };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }


            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            DataTable dt = GetCashInOutList(sYmdFrom, sYmdTo, ComnEtcFunc.CashCode);

            if (dt != null) GridRetr.DataSource = dt;
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ComnEtcFunc.ExportExcelFile("현금출납장", GridRetr);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        #region[Execute By Query]

        private DataTable GetCashInOutList(string sYmdFrom, string sYmdTo, string sCashCd)
        {
            /*
             * 2021-01-27 현업요청_
             * 일계/누계 추가
             */
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();

            strSql.AppendLine("SELECT Y1.*                                                                                                                                                                                         ");
            strSql.AppendLine("  FROM(                                                                                                                                                                                             ");
            strSql.AppendLine("                SELECT X1.SEQ     ");
            strSql.AppendLine("                     , X1.TDATE   ");
            strSql.AppendLine("                     , X1.TDATE1  ");
            strSql.AppendLine("                     , X1.SEQNO   ");
            strSql.AppendLine("                     , X1.LINNO   ");
            strSql.AppendLine("                     , X1.ACCOD   ");
            strSql.AppendLine("                     , X1.ACNAM   ");
            strSql.AppendLine("                     , X1.ATEXT   ");
            strSql.AppendLine("                     , X1.CVCOD   ");
            strSql.AppendLine("                     , X1.CVNAM   ");
            strSql.AppendLine("                     , X1.ACRDR   ");
            strSql.AppendLine("                     , X1.ACAMT   ");
            strSql.AppendLine("                     , X1.ADAMT   ");
            strSql.AppendLine("                     , X1.BAL_AMT ");
            strSql.AppendLine("                     , SUM((ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0)) + BAL_AMT) OVER(ORDER BY TDATE1, SEQNO, LINNO) JJAMT");
            strSql.AppendLine("                     , X1.RK");
            strSql.AppendLine("           FROM(                                                                                                                                                                                    ");
            strSql.AppendLine("                  SELECT '1' AS SEQ                                                                                                                                                                 ");
            strSql.AppendLine("                       , '[ 이월]' AS TDATE                                                                                                                                                         ");
            strSql.AppendLine("                       , '0000-00-00' AS TDATE1                                                                                                                                                     ");
            strSql.AppendLine("                       , '' AS SEQNO                                                                                                                                                                ");
            strSql.AppendLine("                       , 0 AS LINNO                                                                                                                                                                 ");
            strSql.AppendLine("                       , '' AS ACCOD                                                                                                                                                                ");
            strSql.AppendLine("                       , '' AS ACNAM                                                                                                                                                                ");
            strSql.AppendLine("                       , '' AS ATEXT                                                                                                                                                                ");
            strSql.AppendLine("                       , NULL AS CVCOD                                                                                                                                                              ");
            strSql.AppendLine("                       , '' AS CVNAM                                                                                                                                                                ");
            strSql.AppendLine("                       , '' AS ACRDR                                                                                                                                                                ");
            strSql.AppendLine("                       , 0 AS ACAMT                                                                                                                                                                 ");
            strSql.AppendLine("                       , 0 AS ADAMT                                                                                                                                                                 ");
            strSql.AppendLine("                       , SUM(ISNULL(ACJANF_AMT, 0)) - SUM(ISNULL(ACTRAN_AMT, 0)) AS BAL_AMT                                                                                                         ");
            strSql.AppendLine("                              , '' AS RK");
            strSql.AppendLine("                   FROM(                                                                                                                                                                            ");
            strSql.AppendLine("                        SELECT CASE WHEN ROWNUM = '1' THEN ISNULL(CARRY_OVER_AMT1, 0) END AS ACJANF_AMT                                                                                             ");
            strSql.AppendLine("                                , CASE WHEN ROWNUM = '2' THEN ISNULL(CARRY_OVER_AMT2, 0) END AS ACTRAN_AMT                                                                                          ");
            strSql.AppendLine("                          FROM(                                                                                                                                                                     ");
            strSql.AppendLine("                                 SELECT '1' AS ROWNUM                                                                                                                                               ");
            strSql.AppendLine("                                      , ISNULL(SUM(ISNULL(A.ACDRJN, 0) + ISNULL(A.ACCRJN, 0)), 0) AS CARRY_OVER_AMT1                                                                                ");
            strSql.AppendLine("                                       , 0 AS CARRY_OVER_AMT2                                                                                                                                       ");
            strSql.AppendLine("                                   FROM ACJANF A                                                                                                                                                    ");
            strSql.AppendLine("                                   LEFT OUTER JOIN ACMSTF B                                                                                                                                         ");
            strSql.AppendLine("                                     ON A.ACCOD = B.ACCOD                                                                                                                                           ");
            strSql.AppendLine("                                  WHERE A.ACYEAR = '"+ sYmdFrom.Substring(0, 4) + "'                                                                                                                                           ");
            strSql.AppendLine("                                    AND A.ACCOD = '"+ sCashCd + "'                                                                                                                                            ");
            strSql.AppendLine("                                  GROUP BY A.ACCOD                                                                                                                                                  ");
            strSql.AppendLine("                                  UNION ALL                                                                                                                                                         ");
            strSql.AppendLine("                                 SELECT '2' AS ROWNUM                                                                                                                                               ");
            strSql.AppendLine("                                      , 0 AS CARRY_OVER_AMT1                                                                                                                                        ");
            //strSql.AppendLine("                                      , (CASE WHEN ACRDR = '1' THEN SUM(ISNULL(ACAMT, 0)) - SUM(ISNULL(ADAMT, 0)) ELSE SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) END) AS CARRY_OVER_AMT2        ");
            strSql.AppendLine("                                      , SUM(ISNULL(ACAMT, 0)) - SUM(ISNULL(ADAMT, 0)) AS CARRY_OVER_AMT2        ");//#002
            strSql.AppendLine("                                   FROM ACTRAN C                                                                                                                                                    ");
            strSql.AppendLine("                                   LEFT OUTER JOIN ACMSTF D                                                                                                                                         ");
            strSql.AppendLine("                                     ON C.ACCOD = D.ACCOD                                                                                                                                           ");
            strSql.AppendLine("                                  WHERE TDATE BETWEEN SUBSTRING('"+ sYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '"+ sYmdFrom + "')), 112)                                ");
            strSql.AppendLine("                                    AND C.ATGUB IN('1','2')"); 
            strSql.AppendLine("                                    AND C.ACCOD <> '" + sCashCd + "'                                                                                                                                           ");
            strSql.AppendLine("                                  GROUP BY D.ASMCD, ACRDR                                                                                                                                           ");
            strSql.AppendLine("                                  UNION ALL                                                                                                                                                         ");
            strSql.AppendLine("                                 SELECT '2' AS ROWNUM                                                                                                                                               ");
            strSql.AppendLine("                                      , 0 AS CARRY_OVER_AMT1                                                                                                                                        ");
            strSql.AppendLine("                                      , (CASE WHEN ACRDR = '1' THEN SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) ELSE SUM(ISNULL(ACAMT, 0)) - SUM(ISNULL(ADAMT, 0))END) AS CARRY_OVER_AMT2         ");
            strSql.AppendLine("                                   FROM ACTRAN C                                                                                                                                                    ");
            strSql.AppendLine("                                   LEFT OUTER JOIN ACMSTF D                                                                                                                                         ");
            strSql.AppendLine("                                     ON C.ACCOD = D.ACCOD                                                                                                                                           ");
            strSql.AppendLine("                                  WHERE TDATE BETWEEN SUBSTRING('"+ sYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '"+ sYmdFrom + "')), 112)                                ");
            strSql.AppendLine("                                    AND C.ATGUB = '3'                                                                                                                                               ");
            strSql.AppendLine("                                    AND C.ACCOD = '" + sCashCd + "'                                                                                                                                            ");
            strSql.AppendLine("                                  GROUP BY D.ASMCD, ACRDR                                                                                                                                           ");
            strSql.AppendLine("                               ) X1                                                                                                                                                                 ");
            strSql.AppendLine("                      ) X2                                                                                                                                                                          ");
            strSql.AppendLine("                  UNION ALL                                                                                                                                                                         ");
            strSql.AppendLine("                 SELECT YY1.SEQ                                                                                                                                                                     ");
            strSql.AppendLine("                       , YY1.TDATE                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.TDATE1                                                                                                                                                                 ");
            strSql.AppendLine("                       , YY1.SEQNO                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.LINNO                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.ACCOD                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.ACNAM                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.ATEXT                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.CVCOD                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.CVNAM                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.ACRDR                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.ACAMT                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.ADAMT                                                                                                                                                                  ");
            strSql.AppendLine("                       , YY1.JAMT                                                                                                                                                                   ");
            strSql.AppendLine("                       , YY1.RK");
            strSql.AppendLine("                    FROM(                                                                                                                                                                           ");
            strSql.AppendLine("                           SELECT '2' AS SEQ                                                                                                                                                        ");
            strSql.AppendLine("                                , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE                                                                                                        ");
            strSql.AppendLine("                                , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE1                                                                                                       ");
            strSql.AppendLine("                                , A1.SEQNO                                                                                                                                                          ");
            strSql.AppendLine("                                , A1.LINNO                                                                                                                                                          ");
            strSql.AppendLine("                                , A1.ACCOD                                                                                                                                                          ");
            strSql.AppendLine("                                , B1.ACNAM                                                                                                                                                          ");
            strSql.AppendLine("                                , A1.ATEXT AS ATEXT                                                                                                                                                 ");
            strSql.AppendLine("                                , A1.CVCOD AS CVCOD                                                                                                                                                 ");
            strSql.AppendLine("                                , A1.CVNAM AS CVNAM                                                                                                                                                 ");
            strSql.AppendLine("                                , B1.ACRDR                                                                                                                                                          ");
            strSql.AppendLine("                                , ISNULL(A1.ADAMT, 0) AS ACAMT                                                                                                                                      ");
            strSql.AppendLine("                                , ISNULL(A1.ACAMT, 0) AS ADAMT                                                                                                                                      ");
            strSql.AppendLine("                                , 0 AS JAMT                                                                                                                                                         ");
            strSql.AppendLine("                                , A1.RK");
            strSql.AppendLine("                             FROM ACTRAN A1                                                                                                                                                         ");
            strSql.AppendLine("                             LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                                                                             ");
            strSql.AppendLine("                            WHERE A1.TDATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'                                                                                                                        ");
            strSql.AppendLine("                              AND A1.ATGUB IN('1', '2')                                                                                                                                             ");
            strSql.AppendLine("                              AND A1.ACCOD <> '"+ sCashCd + "'                                                                                                                                                ");
            strSql.AppendLine("                            UNION ALL                                                                                                                                                               ");
            strSql.AppendLine("                           SELECT '2' AS SEQ                                                                                                                                                        ");
            strSql.AppendLine("                                , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE                                                                                                        ");
            strSql.AppendLine("                                , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE1                                                                                                       ");
            strSql.AppendLine("                                , A1.SEQNO                                                                                                                                                          ");
            strSql.AppendLine("                                , A1.LINNO                                                                                                                                                          ");
            strSql.AppendLine("                                , A1.ACCOD                                                                                                                                                          ");
            strSql.AppendLine("                                , B1.ACNAM                                                                                                                                                          ");
            strSql.AppendLine("                                , A1.ATEXT AS ATEXT                                                                                                                                                 ");
            strSql.AppendLine("                                , A1.CVCOD AS CVCOD                                                                                                                                                 ");
            strSql.AppendLine("                                , A1.CVNAM AS CVNAM                                                                                                                                                 ");
            strSql.AppendLine("                                , B1.ACRDR                                                                                                                                                          ");
            strSql.AppendLine("                                , ISNULL(A1.ACAMT, 0) AS ACAMT                                                                                                                                      ");
            strSql.AppendLine("                                , ISNULL(A1.ADAMT, 0) AS ADAMT                                                                                                                                      ");
            strSql.AppendLine("                                , 0 AS JAMT                                                                                                                                                         ");
            strSql.AppendLine("                                , A1.RK");
            strSql.AppendLine("                             FROM ACTRAN A1                                                                                                                                                         ");
            strSql.AppendLine("                             LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                                                                             ");
            strSql.AppendLine("                            WHERE A1.TDATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'                                                                                                                        ");
            strSql.AppendLine("                              AND A1.ATGUB = '3'                                                                                                                                                    ");
            strSql.AppendLine("                              AND A1.ACCOD = '"+ sCashCd + "'                                                                                                                                                 ");
            strSql.AppendLine("                         ) YY1                                                                                                                                                                      ");
            strSql.AppendLine("                 ) X1                                                                                                                                                                               ");
            strSql.AppendLine("             UNION ALL                                                                                                                                                                              ");
            strSql.AppendLine("             SELECT YY1.SEQ                                                                                                                                                                         ");
            strSql.AppendLine("                  , YY1.TDATE                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.TDATE1                                                                                                                                                                      ");
            strSql.AppendLine("                  , YY1.SEQNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.LINNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ATEXT                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACRDR                                                                                                                                                                       ");
            strSql.AppendLine("                  , SUM(YY1.ACAMT) AS ACAMT                                                                                                                                                         ");
            strSql.AppendLine("                  , SUM(YY1.ADAMT) AS ADAMT                                                                                                                                                         ");
            strSql.AppendLine("                  , SUM(YY1.BAL_AMT) AS BAL_AMT                                                                                                                                                     ");
            strSql.AppendLine("                  , YY1.JJAMT                                                                                                                                                                       ");
            strSql.AppendLine("                  , '' AS RK");
            strSql.AppendLine("               FROM(                                                                                                                                                                                ");
            strSql.AppendLine("                      SELECT '3' AS SEQ                                                                                                                                                             ");
            strSql.AppendLine("                           , '[ 일계]' AS TDATE                                                                                                                                                     ");
            strSql.AppendLine("                           , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE1                                                                                                            ");
            strSql.AppendLine("                           , '' AS SEQNO                                                                                                                                                            ");
            strSql.AppendLine("                           , 0 AS LINNO                                                                                                                                                             ");
            strSql.AppendLine("                           , '' AS ACCOD                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ATEXT                                                                                                                                                            ");
            strSql.AppendLine("                           , NULL AS CVCOD                                                                                                                                                          ");
            strSql.AppendLine("                           , '' AS CVNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACRDR                                                                                                                                                            ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) AS ACAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ACAMT, 0)) AS ADAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) AS BAL_AMT                                                                                                               ");
            strSql.AppendLine("                           , 0 AS JJAMT                                                                                                                                                             ");
            strSql.AppendLine("                        FROM ACTRAN A1                                                                                                                                                              ");
            strSql.AppendLine("                       WHERE A1.TDATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'                                                                                                                             ");
            strSql.AppendLine("                         AND A1.ATGUB IN('1', '2')                                                                                                                                                  ");
            strSql.AppendLine("                         AND A1.ACCOD <> '"+ sCashCd + "'                                                                                                                                                     ");
            strSql.AppendLine("                       GROUP BY TDATE                                                                                                                                                               ");
            strSql.AppendLine("                       UNION ALL                                                                                                                                                                    ");
            strSql.AppendLine("                      SELECT '3' AS SEQ                                                                                                                                                             ");
            strSql.AppendLine("                           , '[ 일계]' AS TDATE                                                                                                                                                     ");
            strSql.AppendLine("                           , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE1                                                                                                            ");
            strSql.AppendLine("                           , '' AS SEQNO                                                                                                                                                            ");
            strSql.AppendLine("                           , 0 AS LINNO                                                                                                                                                             ");
            strSql.AppendLine("                           , '' AS ACCOD                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ATEXT                                                                                                                                                            ");
            strSql.AppendLine("                           , NULL AS CVCOD                                                                                                                                                          ");
            strSql.AppendLine("                           , '' AS CVNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACRDR                                                                                                                                                            ");
            strSql.AppendLine("                           , SUM(ISNULL(ACAMT, 0)) AS ACAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) AS ADAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) AS BAL_AMT                                                                                                               ");
            strSql.AppendLine("                           , 0 AS JJAMT                                                                                                                                                             ");
            strSql.AppendLine("                        FROM ACTRAN A1                                                                                                                                                              ");
            strSql.AppendLine("                       WHERE A1.TDATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'                                                                                                                             ");
            strSql.AppendLine("                         AND A1.ATGUB = '3'                                                                                                                                                         ");
            strSql.AppendLine("                         AND A1.ACCOD = '"+ sCashCd + "'                                                                                                                                                      ");
            strSql.AppendLine("                       GROUP BY TDATE                                                                                                                                                               ");
            strSql.AppendLine("                    ) YY1                                                                                                                                                                           ");
            strSql.AppendLine("             GROUP BY YY1.SEQ                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.TDATE                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.TDATE1                                                                                                                                                                      ");
            strSql.AppendLine("                  , YY1.SEQNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.LINNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ATEXT                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACRDR                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.JJAMT                                                                                                                                                                       ");
            strSql.AppendLine("             UNION ALL                                                                                                                                                                              ");
            strSql.AppendLine("             SELECT YY1.SEQ                                                                                                                                                                         ");
            strSql.AppendLine("                  , YY1.TDATE                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.TDATE1                                                                                                                                                                      ");
            strSql.AppendLine("                  , YY1.SEQNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.LINNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ATEXT                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACRDR                                                                                                                                                                       ");
            strSql.AppendLine("                  , SUM(ISNULL(ACAMT, 0)) AS ACAMT                                                                                                                                                  ");
            strSql.AppendLine("                  , SUM(ISNULL(ADAMT, 0)) AS ADAMT                                                                                                                                                  ");
            strSql.AppendLine("                  , SUM(YY1.BAL_AMT) AS BAL_AMT                                                                                                                                                     ");
            strSql.AppendLine("                  , 0 AS JJAMT                                                                                                                                                                      ");
            strSql.AppendLine("                  , '' AS RK");
            strSql.AppendLine("               FROM(                                                                                                                                                                                ");
            strSql.AppendLine("                      SELECT '4' AS SEQ                                                                                                                                                             ");
            strSql.AppendLine("                           , '[ 누계]' AS TDATE                                                                                                                                                     ");
            strSql.AppendLine("                           , CONVERT(VARCHAR(10), CONVERT(DATE, '99991231'), 23) AS TDATE1                                                                                                          ");
            strSql.AppendLine("                           , '' AS SEQNO                                                                                                                                                            ");
            strSql.AppendLine("                           , 0 AS LINNO                                                                                                                                                             ");
            strSql.AppendLine("                           , '' AS ACCOD                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ATEXT                                                                                                                                                            ");
            strSql.AppendLine("                           , NULL AS CVCOD                                                                                                                                                          ");
            strSql.AppendLine("                           , '' AS CVNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACRDR                                                                                                                                                            ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) AS ACAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ACAMT, 0)) AS ADAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) AS BAL_AMT                                                                                                               ");
            strSql.AppendLine("                           , 0 AS JJAMT                                                                                                                                                             ");
            strSql.AppendLine("                        FROM ACTRAN A1                                                                                                                                                              ");
            strSql.AppendLine("                       WHERE A1.TDATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'                                                                                                                             ");
            strSql.AppendLine("                         AND A1.ATGUB IN('1', '2')                                                                                                                                                  ");
            strSql.AppendLine("                         AND A1.ACCOD <> '"+ sCashCd + "'                                                                                                                                                     ");
            strSql.AppendLine("                       UNION ALL                                                                                                                                                                    ");
            strSql.AppendLine("                      SELECT '4' AS SEQ                                                                                                                                                             ");
            strSql.AppendLine("                           , '[ 누계]' AS TDATE                                                                                                                                                     ");
            strSql.AppendLine("                           , CONVERT(VARCHAR(10), CONVERT(DATE, '99991231'), 23) AS TDATE1                                                                                                          ");
            strSql.AppendLine("                           , '' AS SEQNO                                                                                                                                                            ");
            strSql.AppendLine("                           , 0 AS LINNO                                                                                                                                                             ");
            strSql.AppendLine("                           , '' AS ACCOD                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ATEXT                                                                                                                                                            ");
            strSql.AppendLine("                           , NULL AS CVCOD                                                                                                                                                          ");
            strSql.AppendLine("                           , '' AS CVNAM                                                                                                                                                            ");
            strSql.AppendLine("                           , '' AS ACRDR                                                                                                                                                            ");
            strSql.AppendLine("                           , SUM(ISNULL(ACAMT, 0)) AS ACAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) AS ADAMT                                                                                                                                         ");
            strSql.AppendLine("                           , SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) AS BAL_AMT                                                                                                               ");
            strSql.AppendLine("                           , 0 AS JJAMT                                                                                                                                                             ");
            strSql.AppendLine("                        FROM ACTRAN A1                                                                                                                                                              ");
            strSql.AppendLine("                       WHERE A1.TDATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'                                                                                                                             ");
            strSql.AppendLine("                         AND A1.ATGUB = '3'                                                                                                                                                         ");
            strSql.AppendLine("                         AND A1.ACCOD = '"+ sCashCd + "'                                                                                                                                                      ");
            strSql.AppendLine("                    ) YY1                                                                                                                                                                           ");
            strSql.AppendLine("                GROUP BY YY1.SEQ                                                                                                                                                                    ");
            strSql.AppendLine("                  , YY1.TDATE                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.TDATE1                                                                                                                                                                      ");
            strSql.AppendLine("                  , YY1.SEQNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.LINNO                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ATEXT                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVCOD                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.CVNAM                                                                                                                                                                       ");
            strSql.AppendLine("                  , YY1.ACRDR                                                                                                                                                                       ");
            strSql.AppendLine("        ) Y1                                                                                                                                                                                        ");
            strSql.AppendLine("    ORDER BY Y1.TDATE1, Y1.SEQ, Y1.SEQNO, Y1.LINNO                                                                                                                                                  ");

            #region mariaDB
            //strSql.AppendFormat("\r\n ");                                                                                                                                                                                         
            //strSql.AppendFormat("\r\n SELECT Y1.*   ");
            //strSql.AppendFormat("\r\n   FROM (  ");
            //strSql.AppendLine("                SELECT X1.SEQ     ");
            //strSql.AppendLine("                     , X1.TDATE   ");
            //strSql.AppendLine("                     , X1.TDATE1  ");
            //strSql.AppendLine("                     , X1.SEQNO   ");
            //strSql.AppendLine("                     , X1.LINNO   ");
            //strSql.AppendLine("                     , X1.ACCOD   ");
            //strSql.AppendLine("                     , X1.ACNAM   ");
            //strSql.AppendLine("                     , X1.ATEXT   ");
            //strSql.AppendLine("                     , X1.CVCOD   ");
            //strSql.AppendLine("                     , X1.CVNAM   ");
            //strSql.AppendLine("                     , X1.ACRDR   ");
            //strSql.AppendLine("                     , X1.ACAMT   ");
            //strSql.AppendLine("                     , X1.ADAMT   ");
            //strSql.AppendLine("                     , X1.BAL_AMT ");
            //strSql.AppendLine("                     , SUM((IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0)) + BAL_AMT) OVER(ORDER BY TDATE1, SEQNO, LINNO) JJAMT");
            //strSql.AppendLine("                     , X1.RK");
            //strSql.AppendFormat("\r\n            FROM (   ");
            //strSql.AppendFormat("\r\n                   SELECT '1' AS SEQ   ");
            //strSql.AppendFormat("\r\n                        , '[ 이월]' AS TDATE  ");
            //strSql.AppendFormat("\r\n                        , '0000-00-00' AS TDATE1  ");
            //strSql.AppendFormat("\r\n                        , '' AS SEQNO   ");
            //strSql.AppendFormat("\r\n                        , 0 AS LINNO   ");
            //strSql.AppendFormat("\r\n                        , '' AS ACCOD     ");
            //strSql.AppendFormat("\r\n                        , '' AS ACNAM     ");
            //strSql.AppendFormat("\r\n                        , '' AS ATEXT     ");
            //strSql.AppendFormat("\r\n                        , NULL AS CVCOD    ");
            //strSql.AppendFormat("\r\n                        , '' AS CVNAM    ");
            //strSql.AppendFormat("\r\n                        , '' AS ACRDR   ");
            //strSql.AppendFormat("\r\n                        , 0 AS ACAMT    ");
            //strSql.AppendFormat("\r\n                        , 0 AS ADAMT    ");
            //strSql.AppendFormat("\r\n                        , SUM(IFNULL(ACJANF_AMT, 0)) - SUM(IFNULL(ACTRAN_AMT, 0)) AS BAL_AMT    ");
            //strSql.AppendLine("                              , '' AS RK");
            //strSql.AppendFormat("\r\n                    FROM(    ");
            //strSql.AppendFormat("\r\n                         SELECT CASE WHEN ROWNUM = '1' THEN IFNULL(CARRY_OVER_AMT1, 0) END AS ACJANF_AMT    ");
            //strSql.AppendFormat("\r\n                         	    , CASE WHEN ROWNUM = '2' THEN IFNULL(CARRY_OVER_AMT2, 0) END AS ACTRAN_AMT    ");
            //strSql.AppendFormat("\r\n                           FROM (    ");
            //strSql.AppendFormat("\r\n                                  SELECT '1' AS ROWNUM    ");
            //strSql.AppendFormat("\r\n                                       , IFNULL(SUM(IFNULL(A.ACDRJN, 0) + IFNULL(A.ACCRJN, 0)), 0) AS CARRY_OVER_AMT1    ");
            //strSql.AppendFormat("\r\n                                  	  , 0 AS CARRY_OVER_AMT2    ");
            //strSql.AppendFormat("\r\n                                    FROM ACJANF A     ");
            //strSql.AppendFormat("\r\n                                    LEFT OUTER JOIN ACMSTF B    ");
            //strSql.AppendFormat("\r\n                                      ON A.ACCOD = B.ACCOD    ");
            //strSql.AppendFormat("\r\n                                   WHERE A.ACYEAR = '{0}'       ", sYmdFrom.Substring(0, 4));
            //strSql.AppendFormat("\r\n                                     AND A.ACCOD = '{0}'    ", sCashCd);
            //strSql.AppendFormat("\r\n                                   GROUP BY A.ACCOD    ");
            //strSql.AppendFormat("\r\n                                   UNION ALL    ");
            //strSql.AppendFormat("\r\n                                  SELECT '2' AS ROWNUM    ");
            //strSql.AppendFormat("\r\n                                       , 0 AS CARRY_OVER_AMT1    ");
            //strSql.AppendFormat("\r\n                                       , (CASE WHEN ACRDR='1' THEN SUM(IFNULL(ACAMT, 0))-SUM(IFNULL(ADAMT, 0)) ELSE SUM(IFNULL(ADAMT, 0))-SUM(IFNULL(ACAMT, 0)) END) AS CARRY_OVER_AMT2    ");
            //strSql.AppendFormat("\r\n                                    FROM ACTRAN C    ");
            //strSql.AppendFormat("\r\n                                    LEFT OUTER JOIN ACMSTF D   ");
            //strSql.AppendFormat("\r\n                                      ON C.ACCOD = D.ACCOD   ");
            //strSql.AppendFormat("\r\n                                   WHERE TDATE BETWEEN DATE_FORMAT('{0}', '%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('{0}', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')   ", sYmdFrom);
            //strSql.AppendFormat("\r\n                                     AND C.ATGUB IN ( '1', '2' )   ");
            //strSql.AppendFormat("\r\n                                     AND C.ACCOD <> '{0}'       ", sCashCd);
            //strSql.AppendFormat("\r\n                                   GROUP BY D.ASMCD  ");
            //strSql.AppendFormat("\r\n                                   UNION ALL    ");
            //strSql.AppendFormat("\r\n                                  SELECT '2' AS ROWNUM    ");
            //strSql.AppendFormat("\r\n                                       , 0 AS CARRY_OVER_AMT1    ");
            ///* #001 */
            ////strSql.AppendFormat("\r\n                                       , (CASE WHEN ACRDR='1' THEN SUM(IFNULL(ACAMT, 0))-SUM(IFNULL(ADAMT, 0)) ELSE SUM(IFNULL(ADAMT, 0))-SUM(IFNULL(ACAMT, 0)) END) AS CARRY_OVER_AMT2    ");
            //strSql.AppendFormat("\r\n                                       , (CASE WHEN ACRDR='1' THEN SUM(IFNULL(ADAMT, 0))-SUM(IFNULL(ACAMT, 0)) ELSE SUM(IFNULL(ACAMT, 0))-SUM(IFNULL(ADAMT, 0))END) AS CARRY_OVER_AMT2      ");
            //strSql.AppendFormat("\r\n                                    FROM ACTRAN C    ");
            //strSql.AppendFormat("\r\n                                    LEFT OUTER JOIN ACMSTF D   ");
            //strSql.AppendFormat("\r\n                                      ON C.ACCOD = D.ACCOD   ");
            //strSql.AppendFormat("\r\n                                   WHERE TDATE BETWEEN DATE_FORMAT('{0}', '%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('{0}', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')   ", sYmdFrom);
            //strSql.AppendFormat("\r\n                                     AND C.ATGUB = '3' ");
            //strSql.AppendFormat("\r\n                                     AND C.ACCOD = '{0}'       ", sCashCd);
            //strSql.AppendFormat("\r\n                                   GROUP BY D.ASMCD  ");
            //strSql.AppendFormat("\r\n                                ) X1       ");
            //strSql.AppendFormat("\r\n                       ) X2     ");
            //strSql.AppendFormat("\r\n                   UNION ALL   ");
            //strSql.AppendFormat("\r\n                  SELECT YY1.SEQ    ");
            //strSql.AppendFormat("\r\n                       , YY1.TDATE    ");
            //strSql.AppendFormat("\r\n                       , YY1.TDATE1    ");
            //strSql.AppendFormat("\r\n                       , YY1.SEQNO   ");
            //strSql.AppendFormat("\r\n                       , YY1.LINNO   ");
            //strSql.AppendFormat("\r\n                       , YY1.ACCOD   ");
            //strSql.AppendFormat("\r\n                       , YY1.ACNAM     ");
            //strSql.AppendFormat("\r\n                       , YY1.ATEXT     ");
            //strSql.AppendFormat("\r\n                       , YY1.CVCOD    ");
            //strSql.AppendFormat("\r\n                       , YY1.CVNAM    ");
            //strSql.AppendFormat("\r\n                       , YY1.ACRDR   ");
            //strSql.AppendFormat("\r\n                       , YY1.ACAMT    ");
            //strSql.AppendFormat("\r\n                       , YY1.ADAMT    ");
            //strSql.AppendFormat("\r\n                       , YY1.JAMT   ");
            //strSql.AppendLine("                             , YY1.RK");
            //strSql.AppendFormat("\r\n                    FROM ( ");
            //strSql.AppendFormat("\r\n                           SELECT '2' AS SEQ    ");
            //strSql.AppendFormat("\r\n                                , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE    ");
            //strSql.AppendFormat("\r\n                                , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1    ");
            //strSql.AppendFormat("\r\n                                , A1.SEQNO   ");
            //strSql.AppendFormat("\r\n                                , A1.LINNO   ");
            //strSql.AppendFormat("\r\n                                , A1.ACCOD   ");
            //strSql.AppendFormat("\r\n                                , B1.ACNAM     ");
            //strSql.AppendFormat("\r\n                                , A1.ATEXT AS ATEXT     ");
            //strSql.AppendFormat("\r\n                                , A1.CVCOD AS CVCOD    ");
            //strSql.AppendFormat("\r\n                                , A1.CVNAM AS CVNAM    ");
            //strSql.AppendFormat("\r\n                                , B1.ACRDR   ");
            //strSql.AppendFormat("\r\n                                , IFNULL(A1.ADAMT, 0) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                                , IFNULL(A1.ACAMT, 0) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                                , 0 AS JAMT   ");
            //strSql.AppendLine("                                      , A1.RK");
            //strSql.AppendFormat("\r\n                             FROM ACTRAN A1    ");
            //strSql.AppendFormat("\r\n                             LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD    ");
            //strSql.AppendFormat("\r\n                            WHERE A1.TDATE BETWEEN '{0}' AND '{1}'    ", sYmdFrom, sYmdTo);
            //strSql.AppendFormat("\r\n                              AND A1.ATGUB IN ( '1', '2' )   ");
            //strSql.AppendFormat("\r\n                              AND A1.ACCOD <> '{0}' ", sCashCd);
            //strSql.AppendFormat("\r\n                            UNION ALL ");
            //strSql.AppendFormat("\r\n                           SELECT '2' AS SEQ    ");
            //strSql.AppendFormat("\r\n                                , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE    ");
            //strSql.AppendFormat("\r\n                                , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1    ");
            //strSql.AppendFormat("\r\n                                , A1.SEQNO   ");
            //strSql.AppendFormat("\r\n                                , A1.LINNO   ");
            //strSql.AppendFormat("\r\n                                , A1.ACCOD   ");
            //strSql.AppendFormat("\r\n                                , B1.ACNAM     ");
            //strSql.AppendFormat("\r\n                                , A1.ATEXT AS ATEXT     ");
            //strSql.AppendFormat("\r\n                                , A1.CVCOD AS CVCOD    ");
            //strSql.AppendFormat("\r\n                                , A1.CVNAM AS CVNAM    ");
            //strSql.AppendFormat("\r\n                                , B1.ACRDR   ");
            //strSql.AppendFormat("\r\n                                , IFNULL(A1.ACAMT, 0) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                                , IFNULL(A1.ADAMT, 0) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                                , 0 AS JAMT   ");
            //strSql.AppendLine("                                      , A1.RK");
            //strSql.AppendFormat("\r\n                             FROM ACTRAN A1    ");
            //strSql.AppendFormat("\r\n                             LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD    ");
            //strSql.AppendFormat("\r\n                            WHERE A1.TDATE BETWEEN '{0}' AND '{1}'    ", sYmdFrom, sYmdTo);
            //strSql.AppendFormat("\r\n                              AND A1.ATGUB = '3' ");
            //strSql.AppendFormat("\r\n                              AND A1.ACCOD ='{0}' ", sCashCd);
            //strSql.AppendFormat("\r\n                         ) YY1 ");
            //strSql.AppendFormat("\r\n                 ) X1   ");
            //strSql.AppendFormat("\r\n             UNION ALL  ");
            //strSql.AppendFormat("\r\n             SELECT YY1.SEQ   ");
            //strSql.AppendFormat("\r\n                  , YY1.TDATE    ");
            //strSql.AppendFormat("\r\n                  , YY1.TDATE1  ");
            //strSql.AppendFormat("\r\n                  , YY1.SEQNO   ");
            //strSql.AppendFormat("\r\n                  , YY1.LINNO   ");
            //strSql.AppendFormat("\r\n                  , YY1.ACCOD     ");
            //strSql.AppendFormat("\r\n                  , YY1.ACNAM     ");
            //strSql.AppendFormat("\r\n                  , YY1.ATEXT     ");
            //strSql.AppendFormat("\r\n                  , YY1.CVCOD    ");
            //strSql.AppendFormat("\r\n                  , YY1.CVNAM    ");
            //strSql.AppendFormat("\r\n                  , YY1.ACRDR   ");
            //strSql.AppendFormat("\r\n                  , SUM(YY1.ACAMT) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                  , SUM(YY1.ADAMT) AS ADAMT ");
            //strSql.AppendFormat("\r\n                  , SUM(YY1.BAL_AMT) AS BAL_AMT    ");
            //strSql.AppendFormat("\r\n                  , YY1.JJAMT  ");
            //strSql.AppendLine("                        , '' AS RK");
            //strSql.AppendFormat("\r\n               FROM ( ");
            //strSql.AppendFormat("\r\n                      SELECT '3' AS SEQ   ");
            //strSql.AppendFormat("\r\n                           , '[ 일계]' AS TDATE    ");
            //strSql.AppendFormat("\r\n                           , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1  ");
            //strSql.AppendFormat("\r\n                           , '' AS SEQNO   ");
            //strSql.AppendFormat("\r\n                           , 0 AS LINNO   ");
            //strSql.AppendFormat("\r\n                           , '' AS ACCOD     ");
            //strSql.AppendFormat("\r\n                           , '' AS ACNAM     ");
            //strSql.AppendFormat("\r\n                           , '' AS ATEXT     ");
            //strSql.AppendFormat("\r\n                           , NULL AS CVCOD    ");
            //strSql.AppendFormat("\r\n                           , '' AS CVNAM    ");
            //strSql.AppendFormat("\r\n                           , '' AS ACRDR   ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ACAMT, 0)) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT    ");
            //strSql.AppendFormat("\r\n                           , 0 AS JJAMT  ");
            //strSql.AppendFormat("\r\n                        FROM ACTRAN A1  ");
            //strSql.AppendFormat("\r\n                       WHERE A1.TDATE BETWEEN '{0}' AND '{1}'  ", sYmdFrom, sYmdTo);
            //strSql.AppendFormat("\r\n                         AND A1.ATGUB IN ('1', '2')  ");
            //strSql.AppendFormat("\r\n                         AND A1.ACCOD <> '{0}'  ", sCashCd);
            //strSql.AppendFormat("\r\n                       GROUP BY TDATE  ");
            //strSql.AppendFormat("\r\n                       UNION ALL ");
            //strSql.AppendFormat("\r\n                      SELECT '3' AS SEQ   ");
            //strSql.AppendFormat("\r\n                           , '[ 일계]' AS TDATE    ");
            //strSql.AppendFormat("\r\n                           , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1  ");
            //strSql.AppendFormat("\r\n                           , '' AS SEQNO   ");
            //strSql.AppendFormat("\r\n                           , 0 AS LINNO   ");
            //strSql.AppendFormat("\r\n                           , '' AS ACCOD     ");
            //strSql.AppendFormat("\r\n                           , '' AS ACNAM     ");
            //strSql.AppendFormat("\r\n                           , '' AS ATEXT     ");
            //strSql.AppendFormat("\r\n                           , NULL AS CVCOD    ");
            //strSql.AppendFormat("\r\n                           , '' AS CVNAM    ");
            //strSql.AppendFormat("\r\n                           , '' AS ACRDR   ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ACAMT, 0)) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT    ");
            //strSql.AppendFormat("\r\n                           , 0 AS JJAMT  ");
            //strSql.AppendFormat("\r\n                        FROM ACTRAN A1  ");
            //strSql.AppendFormat("\r\n                       WHERE A1.TDATE BETWEEN '{0}' AND '{1}'  ", sYmdFrom, sYmdTo);
            //strSql.AppendFormat("\r\n                         AND A1.ATGUB = '3' ");
            //strSql.AppendFormat("\r\n                         AND A1.ACCOD = '{0}'  ", sCashCd);
            //strSql.AppendFormat("\r\n                       GROUP BY TDATE  ");
            //strSql.AppendFormat("\r\n                    ) YY1 ");
            //strSql.AppendFormat("\r\n             GROUP BY YY1.TDATE1 ");
            //strSql.AppendFormat("\r\n             UNION ALL  ");
            //strSql.AppendFormat("\r\n             SELECT YY1.SEQ   ");
            //strSql.AppendFormat("\r\n                  , YY1.TDATE  ");
            //strSql.AppendFormat("\r\n                  , YY1.TDATE1  ");
            //strSql.AppendFormat("\r\n                  , YY1.SEQNO   ");
            //strSql.AppendFormat("\r\n                  , YY1.LINNO   ");
            //strSql.AppendFormat("\r\n                  , YY1.ACCOD     ");
            //strSql.AppendFormat("\r\n                  , YY1.ACNAM     ");
            //strSql.AppendFormat("\r\n                  , YY1.ATEXT     ");
            //strSql.AppendFormat("\r\n                  , YY1.CVCOD    ");
            //strSql.AppendFormat("\r\n                  , YY1.CVNAM    ");
            //strSql.AppendFormat("\r\n                  , YY1.ACRDR   ");
            //strSql.AppendFormat("\r\n                  , SUM(IFNULL(ACAMT, 0)) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                  , SUM(IFNULL(ADAMT, 0)) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                  , SUM(YY1.BAL_AMT) AS BAL_AMT    ");
            //strSql.AppendFormat("\r\n                  , 0 AS JJAMT  ");
            //strSql.AppendLine("                        , '' AS RK");
            //strSql.AppendFormat("\r\n               FROM ( ");
            //strSql.AppendFormat("\r\n                      SELECT '4' AS SEQ   ");
            //strSql.AppendFormat("\r\n                           , '[ 누계]' AS TDATE  ");
            //strSql.AppendFormat("\r\n                           , DATE_FORMAT('{0}', '%Y-%m-%d') AS TDATE1  ", sYmdTo);
            //strSql.AppendFormat("\r\n                           , '' AS SEQNO   ");
            //strSql.AppendFormat("\r\n                           , 0 AS LINNO   ");
            //strSql.AppendFormat("\r\n                           , '' AS ACCOD     ");
            //strSql.AppendFormat("\r\n                           , '' AS ACNAM     ");
            //strSql.AppendFormat("\r\n                           , '' AS ATEXT     ");
            //strSql.AppendFormat("\r\n                           , NULL AS CVCOD    ");
            //strSql.AppendFormat("\r\n                           , '' AS CVNAM    ");
            //strSql.AppendFormat("\r\n                           , '' AS ACRDR   ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ACAMT, 0)) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT    ");
            //strSql.AppendFormat("\r\n                           , 0 AS JJAMT  ");
            //strSql.AppendFormat("\r\n                        FROM ACTRAN A1  ");
            //strSql.AppendFormat("\r\n                       WHERE A1.TDATE BETWEEN '{0}' AND '{1}'  ", sYmdFrom, sYmdTo);
            //strSql.AppendFormat("\r\n                         AND A1.ATGUB IN ('1', '2')  ");
            //strSql.AppendFormat("\r\n                         AND A1.ACCOD <> '{0}'  ", sCashCd);
            //strSql.AppendFormat("\r\n                       UNION ALL ");
            //strSql.AppendFormat("\r\n                      SELECT '4' AS SEQ   ");
            //strSql.AppendFormat("\r\n                           , '[ 누계]' AS TDATE  ");
            //strSql.AppendFormat("\r\n                           , DATE_FORMAT('{0}', '%Y-%m-%d') AS TDATE1  ", sYmdFrom);
            //strSql.AppendFormat("\r\n                           , '' AS SEQNO   ");
            //strSql.AppendFormat("\r\n                           , 0 AS LINNO   ");
            //strSql.AppendFormat("\r\n                           , '' AS ACCOD     ");
            //strSql.AppendFormat("\r\n                           , '' AS ACNAM     ");
            //strSql.AppendFormat("\r\n                           , '' AS ATEXT     ");
            //strSql.AppendFormat("\r\n                           , NULL AS CVCOD    ");
            //strSql.AppendFormat("\r\n                           , '' AS CVNAM    ");
            //strSql.AppendFormat("\r\n                           , '' AS ACRDR   ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ACAMT, 0)) AS ACAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) AS ADAMT    ");
            //strSql.AppendFormat("\r\n                           , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT    ");
            //strSql.AppendFormat("\r\n                           , 0 AS JJAMT  ");
            //strSql.AppendFormat("\r\n                        FROM ACTRAN A1  ");
            //strSql.AppendFormat("\r\n                       WHERE A1.TDATE BETWEEN '{0}' AND '{1}'  ", sYmdFrom, sYmdTo);
            //strSql.AppendFormat("\r\n                         AND A1.ATGUB = '3' ");
            //strSql.AppendFormat("\r\n                         AND A1.ACCOD = '{0}'  ", sCashCd);
            //strSql.AppendFormat("\r\n                    ) YY1 ");
            //strSql.AppendFormat("\r\n        ) Y1  ");
            //strSql.AppendFormat("\r\n    ORDER BY Y1.TDATE1, Y1.SEQ, Y1.SEQNO, Y1.LINNO  ");
            #endregion

            #region[2021-04-20 이전쿼리]

            //strSql.Clear();
            //strSql.AppendLine("SELECT Y1.*  ");
            //strSql.AppendLine("  FROM ( ");
            //strSql.AppendLine("         SELECT X1.*  ");
            //strSql.AppendLine("              , SUM((CASE WHEN ACRDR='1' THEN IFNULL(ACAMT, 0)-IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0)-IFNULL(ACAMT, 0) END)+BAL_AMT) OVER(ORDER BY SEQ, TDATE, SEQNO, LINNO) JJAMT  ");
            //strSql.AppendLine("           FROM (  ");
            //strSql.AppendLine("                  SELECT '1' AS SEQ  ");
            //strSql.AppendLine("                       , '[ 이월]' AS TDATE ");
            //strSql.AppendLine("                       , '0000-00-00' AS TDATE1 ");
            //strSql.AppendLine("                       , '' AS SEQNO  ");
            //strSql.AppendLine("                       , 0 AS LINNO  ");
            //strSql.AppendLine("                       , '' AS ACCOD    ");
            //strSql.AppendLine("                       , '' AS ACNAM    ");
            //strSql.AppendLine("                       , '' AS ATEXT    ");
            //strSql.AppendLine("                       , NULL AS CVCOD   ");
            //strSql.AppendLine("                       , '' AS CVNAM   ");
            //strSql.AppendLine("                       , '' AS ACRDR  ");
            //strSql.AppendLine("                       , 0 AS ACAMT   ");
            //strSql.AppendLine("                       , 0 AS ADAMT   ");
            //strSql.AppendLine("                       , SUM(IFNULL(ACJANF_AMT, 0)) - SUM(IFNULL(ACTRAN_AMT, 0)) AS BAL_AMT   ");
            //strSql.AppendLine("                   FROM(   ");
            //strSql.AppendLine("                        SELECT CASE WHEN ROWNUM = '1' THEN IFNULL(CARRY_OVER_AMT1, 0) END AS ACJANF_AMT   ");
            //strSql.AppendLine("                        	    , CASE WHEN ROWNUM = '2' THEN IFNULL(CARRY_OVER_AMT2, 0) END AS ACTRAN_AMT   ");
            //strSql.AppendLine("                          FROM (   ");
            //strSql.AppendLine("                                 SELECT '1' AS ROWNUM   ");
            //strSql.AppendLine("                                      , IFNULL(SUM(IFNULL(A.ACDRJN, 0) + IFNULL(A.ACCRJN, 0)), 0) AS CARRY_OVER_AMT1   ");
            //strSql.AppendLine("                                 	  , 0 AS CARRY_OVER_AMT2   ");
            //strSql.AppendLine("                                   FROM ACJANF A    ");
            //strSql.AppendLine("                                   LEFT OUTER JOIN ACMSTF B   ");
            //strSql.AppendLine("                                     ON A.ACCOD = B.ACCOD   ");
            //strSql.AppendLine("                                  WHERE A.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'      ");
            //strSql.AppendLine("                                    AND A.ACCOD = '" + sCashCd + "'   ");
            //strSql.AppendLine("                                  GROUP BY A.ACCOD   ");
            //strSql.AppendLine("                                  UNION ALL   ");
            //strSql.AppendLine("                                 SELECT '2' AS ROWNUM   ");
            //strSql.AppendLine("                                      , 0 AS CARRY_OVER_AMT1   ");
            //strSql.AppendLine("                                      , (CASE WHEN ACRDR='1' THEN SUM(IFNULL(ACAMT, 0))-SUM(IFNULL(ADAMT, 0)) ELSE SUM(IFNULL(ADAMT, 0))-SUM(IFNULL(ACAMT, 0)) END) AS CARRY_OVER_AMT2   ");
            //strSql.AppendLine("                                   FROM ACTRAN C   ");
            //strSql.AppendLine("                                   LEFT OUTER JOIN ACMSTF D  ");
            //strSql.AppendLine("                                     ON C.ACCOD = D.ACCOD  ");
            //strSql.AppendLine("                                  WHERE TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "', '%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')  ");
            //strSql.AppendLine("                                    AND C.ATGUB IN ( '1', '2' )  ");
            //strSql.AppendLine("                                    AND C.ACCOD <> '" + sCashCd + "'      ");
            //strSql.AppendLine("                                  GROUP BY D.ASMCD   ");
            //strSql.AppendLine("                               ) X1      ");
            //strSql.AppendLine("                      ) X2    ");
            //strSql.AppendLine("                  UNION ALL  ");
            //strSql.AppendLine("                 SELECT '2' AS SEQ   ");
            //strSql.AppendLine("                      , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE   ");
            //strSql.AppendLine("                      , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1   ");
            //strSql.AppendLine("                      , A1.SEQNO  ");
            //strSql.AppendLine("                      , A1.LINNO  ");
            //strSql.AppendLine("                      , A1.ACCOD  ");
            //strSql.AppendLine("                      , B1.ACNAM    ");
            //strSql.AppendLine("                      , A1.ATEXT AS ATEXT    ");
            //strSql.AppendLine("                      , A1.CVCOD AS CVCOD   ");
            //strSql.AppendLine("                      , A1.CVNAM AS CVNAM   ");
            //strSql.AppendLine("                      , B1.ACRDR  ");
            //strSql.AppendLine("                      , IFNULL(A1.ADAMT, 0) AS ACAMT   ");
            //strSql.AppendLine("                      , IFNULL(A1.ACAMT, 0) AS ADAMT   ");
            //strSql.AppendLine("                      , 0 AS JAMT  ");
            //strSql.AppendLine("                   FROM ACTRAN A1   ");
            //strSql.AppendLine("                   LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD   ");
            //strSql.AppendLine("                  WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'   ");
            //strSql.AppendLine("                    AND A1.ATGUB IN ( '1', '2' )  ");
            //strSql.AppendLine("                    AND A1.ACCOD <> '" + sCashCd + "'   ");
            //strSql.AppendLine("                ) X1  ");
            //strSql.AppendLine("            UNION ALL ");
            //strSql.AppendLine("            SELECT '3' AS SEQ  ");
            //strSql.AppendLine("                 , '[ 일계]' AS TDATE   ");
            //strSql.AppendLine("                 , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1 ");
            //strSql.AppendLine("                 , '' AS SEQNO  ");
            //strSql.AppendLine("                 , 0 AS LINNO  ");
            //strSql.AppendLine("                 , '' AS ACCOD    ");
            //strSql.AppendLine("                 , '' AS ACNAM    ");
            //strSql.AppendLine("                 , '' AS ATEXT    ");
            //strSql.AppendLine("                 , NULL AS CVCOD   ");
            //strSql.AppendLine("                 , '' AS CVNAM   ");
            //strSql.AppendLine("                 , '' AS ACRDR  ");
            //strSql.AppendLine("                 , SUM(IFNULL(ADAMT, 0)) AS ACAMT   ");
            //strSql.AppendLine("                 , SUM(IFNULL(ACAMT, 0)) AS ADAMT   ");
            //strSql.AppendLine("                 , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT   ");
            //strSql.AppendLine("                 , 0 AS JJAMT ");
            //strSql.AppendLine("              FROM ACTRAN A1 ");
            //strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("               AND A1.ATGUB IN ('1', '2') ");
            //strSql.AppendLine("               AND A1.ACCOD <> '" + sCashCd + "' ");
            //strSql.AppendLine("             GROUP BY TDATE ");
            //strSql.AppendLine("            UNION ALL ");
            //strSql.AppendLine("            SELECT '4' AS SEQ  ");
            //strSql.AppendLine("                 , '[ 누계]' AS TDATE ");
            //strSql.AppendLine("                 , DATE_FORMAT('" + sYmdTo + "', '%Y-%m-%d') AS TDATE1 ");
            //strSql.AppendLine("                 , '' AS SEQNO  ");
            //strSql.AppendLine("                 , 0 AS LINNO  ");
            //strSql.AppendLine("                 , '' AS ACCOD    ");
            //strSql.AppendLine("                 , '' AS ACNAM    ");
            //strSql.AppendLine("                 , '' AS ATEXT    ");
            //strSql.AppendLine("                 , NULL AS CVCOD   ");
            //strSql.AppendLine("                 , '' AS CVNAM   ");
            //strSql.AppendLine("                 , '' AS ACRDR  ");
            //strSql.AppendLine("                 , SUM(IFNULL(ADAMT, 0)) AS ACAMT   ");
            //strSql.AppendLine("                 , SUM(IFNULL(ACAMT, 0)) AS ADAMT   ");
            //strSql.AppendLine("                 , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT   ");
            //strSql.AppendLine("                 , 0 AS JJAMT ");
            //strSql.AppendLine("              FROM ACTRAN A1 ");
            //strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("               AND A1.ATGUB IN ('1', '2') ");
            //strSql.AppendLine("               AND A1.ACCOD <> '" + sCashCd + "' ");
            //strSql.AppendLine("       ) Y1 ");
            //strSql.AppendLine("   ORDER BY Y1.TDATE1, Y1.SEQ ");

            #endregion[2021-04-20 이전쿼리]

            #region[2021-01-27 이전쿼리]

            //strSql.AppendLine(" SELECT X1.* ");
            //strSql.AppendLine("      , SUM((CASE WHEN ACRDR='1' THEN IFNULL(ACAMT, 0)-IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0)-IFNULL(ACAMT, 0) END)+BAL_AMT) OVER(ORDER BY SEQ, TDATE, SEQNO, LINNO) JJAMT ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("      SELECT '1' AS SEQ ");
            //strSql.AppendLine("           , '[ 이월]' AS TDATE  ");
            //strSql.AppendLine("           , '' AS SEQNO ");
            //strSql.AppendLine("           , 0 AS LINNO ");
            //strSql.AppendLine("           , '' AS ACCOD   ");
            //strSql.AppendLine("           , '' AS ACNAM   ");
            //strSql.AppendLine("           , '' AS ATEXT   ");
            //strSql.AppendLine("           , NULL AS CVCOD  ");
            //strSql.AppendLine("           , '' AS CVNAM  ");
            //strSql.AppendLine("           , '' AS ACRDR ");
            //strSql.AppendLine("           , 0 AS ACAMT  ");
            //strSql.AppendLine("           , 0 AS ADAMT  ");
            //strSql.AppendLine("           , SUM(IFNULL(ACJANF_AMT, 0)) - SUM(IFNULL(ACTRAN_AMT, 0)) AS BAL_AMT  ");
            //strSql.AppendLine("       FROM(  ");
            //strSql.AppendLine("             SELECT CASE WHEN ROWNUM = '1' THEN IFNULL(CARRY_OVER_AMT1, 0) END AS ACJANF_AMT  ");
            //strSql.AppendLine("             	  , CASE WHEN ROWNUM = '2' THEN IFNULL(CARRY_OVER_AMT2, 0) END AS ACTRAN_AMT  ");
            //strSql.AppendLine("               FROM (  ");
            //strSql.AppendLine("                      SELECT '1' AS ROWNUM  ");
            //strSql.AppendLine("                           , IFNULL(SUM(IFNULL(A.ACDRJN, 0) + IFNULL(A.ACCRJN, 0)), 0) AS CARRY_OVER_AMT1  ");
            //strSql.AppendLine("                      	  , 0 AS CARRY_OVER_AMT2  ");
            //strSql.AppendLine("                        FROM ACJANF A   ");
            //strSql.AppendLine("                        LEFT OUTER JOIN ACMSTF B  ");
            //strSql.AppendLine("                          ON A.ACCOD = B.ACCOD  ");
            //strSql.AppendLine("                       WHERE A.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'     ");
            //strSql.AppendLine("                         AND A.ACCOD = '" + sCashCd + "'  ");
            //strSql.AppendLine("                       GROUP BY A.ACCOD  ");
            //strSql.AppendLine("                       UNION ALL  ");
            //strSql.AppendLine("                      SELECT '2' AS ROWNUM  ");
            //strSql.AppendLine("                           , 0 AS CARRY_OVER_AMT1  ");
            //strSql.AppendLine("                           , (CASE WHEN ACRDR='1' THEN SUM(IFNULL(ACAMT, 0))-SUM(IFNULL(ADAMT, 0)) ELSE SUM(IFNULL(ADAMT, 0))-SUM(IFNULL(ACAMT, 0)) END) AS CARRY_OVER_AMT2  ");
            //strSql.AppendLine("                        FROM ACTRAN C  ");
            //strSql.AppendLine("                        LEFT OUTER JOIN ACMSTF D ");
            //strSql.AppendLine("                          ON C.ACCOD = D.ACCOD ");
            //strSql.AppendLine("                       WHERE TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "', '%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d') ");
            //strSql.AppendLine("                         AND C.ATGUB IN ( '1', '2' ) ");
            //strSql.AppendLine("                         AND C.ACCOD <> '" + sCashCd + "'     ");
            //strSql.AppendLine("                       GROUP BY D.ASMCD  ");
            //strSql.AppendLine("                    ) X1     ");
            //strSql.AppendLine("           ) X2   ");
            //strSql.AppendLine("       UNION ALL ");
            //strSql.AppendLine("      SELECT '2' AS SEQ  ");
            //strSql.AppendLine("           , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE  ");
            //strSql.AppendLine("           , A1.SEQNO ");
            //strSql.AppendLine("           , A1.LINNO ");
            //strSql.AppendLine("           , A1.ACCOD ");
            //strSql.AppendLine("           , B1.ACNAM   ");
            //strSql.AppendLine("           , A1.ATEXT AS ATEXT   ");
            //strSql.AppendLine("           , A1.CVCOD AS CVCOD  ");
            //strSql.AppendLine("           , A1.CVNAM AS CVNAM  ");
            //strSql.AppendLine("           , B1.ACRDR ");
            //strSql.AppendLine("           , IFNULL(A1.ADAMT, 0) AS ACAMT  ");
            //strSql.AppendLine("           , IFNULL(A1.ACAMT, 0) AS ADAMT  ");
            //strSql.AppendLine("           , 0 AS JAMT ");
            //strSql.AppendLine("        FROM ACTRAN A1  ");
            //strSql.AppendLine("        LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD  ");
            //strSql.AppendLine("       WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //strSql.AppendLine("         AND A1.ATGUB IN ( '1', '2' ) ");
            //strSql.AppendLine("         AND A1.ACCOD <> '" + sCashCd + "'  ");
            //strSql.AppendLine("       ) X1 ");

            #endregion[2021-01-27 이전쿼리]

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #endregion[Execute By Query]

        #region[GridView's Design]

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            /*
             * 2021-01-27
             * 일계 및 누계는 색깔 달리함 (IDX : 1, 3, 4)
             * 쿼리참조
             */
            string sIDX = GridViewRetr.GetRowCellValue(e.RowHandle, GridColSEQ)?.ToString();
            if (!string.IsNullOrEmpty(sIDX) && (sIDX.Equals("3") || sIDX.Equals("4") || sIDX.Equals("1")))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
            //ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        #endregion[GridView's Design]

        private void AC06001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
            {
                
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void AC06001F01_TextChanged(object sender, EventArgs e)
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

        private void DropBtnPrint_Click(object sender, EventArgs e)
        {
            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "출금전표")
            {
                Print("1");
            }
            else if (tag == "입금전표")
            {
                Print("2");
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnSlipOut;
        BarButtonItem BtnSlipIn;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnSlipOut = new BarButtonItem(barManager1, "출금전표");
            BtnSlipIn = new BarButtonItem(barManager1, "입금전표");
            popupMenu1.AddItem(BtnSlipOut);
            popupMenu1.AddItem(BtnSlipIn);

            DropBtnPrint.DropDownControl = popupMenu1;

            BtnSlipOut.Tag = "출금전표";
            BtnSlipOut.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlipOut_ItemClick);

            BtnSlipIn.Tag = "입금전표";
            BtnSlipIn.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlipIn_ItemClick);
        }

        private void BtnSlipOut_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            Print("1");
        }

        private void BtnSlipIn_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            Print("2");
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnPrint.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnPrint.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnPrint.Tag = submenuItem.Tag;
            DropBtnPrint.Text = string.Format("{0}", submenuItem.Tag);
        }

        private void Print(string sGb)
        {
            try
            {
                if (GridViewRetr.RowCount == 0)
                {
                    XtraMessageBox.Show("리스트에 내용이 존재하지 않습니다.");
                    return;
                }

                string sSEQ = GridViewRetr.GetFocusedRowCellValue(GridColSEQ)?.ToString();
                if (!sSEQ.Equals("2"))
                {
                    XtraMessageBox.Show("해당 건은 입출금전표가 아닙니다.");
                    return;
                }

                string sYmd = GridViewRetr.GetFocusedRowCellValue(GridColSlipDt)?.ToString().Replace("-", "").Trim();

                Cursor = Cursors.WaitCursor;

                /*
                 * 2021-02-04 (현업요청)
                 * 1. 입출금 전표는 Focusing된 Row 건별로 발행
                 */

                //DataTable dt1 = GetInfo(sGb);
                //DataTable dt = GetInfo(sGb);
                //DataTable dt1 = GetInfo(sGb);

                /*
                 * 2021-02-05(현업요청)
                 * 1. 전표출력은 Focusing 행의 전표일자를 기준하여 출력
                 */
                //DataTable dt1 = GetInfo(sGb, sYmd);
                DataTable dt = GetInfo(sGb, sYmd);
                dt.Columns["ACAMT"].DataType = typeof(string);
                dt.Columns["ADAMT"].DataType = typeof(string);

                ////한건씩만 출력되기 때문에 행번은 1로 고정
                //DataRow rowCopy = GridViewRetr.GetFocusedDataRow();
                
                //dt.ImportRow(rowCopy);
                //dt.Rows[0]["SEQ"] = 1;
                //dt.Rows[0]["ACAMT"] = string.Format("{0:n0}", Convert.ToDouble(dt.Rows[0]["ACAMT"]));
                //dt.Rows[0]["ADAMT"] = string.Format("{0:n0}", Convert.ToDouble(dt.Rows[0]["ADAMT"]));
                if (dt.Rows.Count < 16)
                {
                    for(int i = dt.Rows.Count - 1; i < 16; i++)
                    {
                        DataRow row = dt.NewRow();
                        dt.Rows.Add(row);
                    }
                }

                DataTable dt2 = GetInfoSummary(sGb, sYmd);
                List<string> list = new List<string>();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CONCAT(B.EMP_NM, ' ', C.COM_NM) AS EMP_GD ");
                strSql.AppendLine("      , B.EMP_NM ");
                strSql.AppendLine("   FROM ZUSRLST A ");
                strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B  ");
                strSql.AppendLine("     ON A.INSANO = B.EMP_ID ");
                strSql.AppendLine("   LEFT JOIN COM_BASE_CD C  ");
                strSql.AppendLine("     ON B.GRADE_CD = C.COM_CD ");
                strSql.AppendLine("    AND C.CD_GB = 'GRADE_CD' ");
                strSql.AppendLine("  WHERE A.USRCD = " + FmMainToolBar2.UserID + "");

                DataTable dtEmp = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (sGb.Equals("1"))
                {
                    list.Add("출 금");
                }
                else
                {
                    list.Add("입 금");
                }

                if (dtEmp.Rows.Count > 0)
                {
                    list.Add(dtEmp.Rows[0]["EMP_GD"]?.ToString());
                    list.Add(dtEmp.Rows[0]["EMP_NM"]?.ToString());
                }
                else
                {
                    list.Add("");
                    list.Add("");
                }

                if (dt2.Rows.Count > 0)
                {
                    list.Add(dt2.Rows[0]["ACAMT"]?.ToString());
                    list.Add(dt2.Rows[0]["ADAMT"]?.ToString());
                }
                else
                {
                    list.Add("0");
                    list.Add("0");
                }

                list.Add(GridViewRetr.GetFocusedRowCellValue(GridColSlipDt)?.ToString());

                Cursor = Cursors.Default;
                ReportViewer fm = new ReportViewer(dt, list, "RptInOutSlip");
                fm.Show();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
           
        }

        private DataTable GetInfo(string sGB, string sYmd)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sCashCd = ComnEtcFunc.CashCode;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine(" SELECT  @ROWNUM := @ROWNUM + 1 AS SEQ");
            //strSql.AppendLine("      , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE ");
            //strSql.AppendLine("      , A1.SEQNO ");
            //strSql.AppendLine("      , A1.LINNO ");
            //strSql.AppendLine("      , A1.ACCOD ");
            //strSql.AppendLine("      , B1.ACNAM   ");
            //strSql.AppendLine("      , A1.ATEXT AS ATEXT   ");
            //strSql.AppendLine("      , A1.CVCOD AS CVCOD  ");
            //strSql.AppendLine("      , A1.CVNAM AS CVNAM  ");
            //strSql.AppendLine("      , CASE WHEN LENGTH(C.IDT_NO) = 10 THEN CONCAT(SUBSTRING(C.IDT_NO, 1, 3), '-', SUBSTRING(C.IDT_NO, 4, 2), '-', SUBSTRING(C.IDT_NO, 6, 5)) ");
            //strSql.AppendLine("                                        ELSE C.IDT_NO END AS IDT_NO");
            //strSql.AppendLine("      , B1.ACRDR ");
            //strSql.AppendLine("      , FORMAT(IFNULL(A1.ACAMT, 0), 0) AS ACAMT  ");
            //strSql.AppendLine("      , FORMAT(IFNULL(A1.ADAMT, 0), 0) AS ADAMT  ");
            //strSql.AppendLine("      , 0 AS JAMT ");
            //strSql.AppendLine("   FROM ACTRAN A1  ");
            //strSql.AppendLine("   LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD  ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C ");
            //strSql.AppendLine("     ON A1.CVCOD = C.DEALER_CD ");
            //strSql.AppendLine("  WHERE A1.TDATE BETWEEN '" + sYmd + "' AND '" + sYmd + "'  ");
            //strSql.AppendLine("    AND A1.ATGUB IN ( '" + sGB + "' ) ");
            //strSql.AppendLine("    AND A1.ACCOD <> '" + sCashCd + "'  ");
            //strSql.AppendLine("    AND ( @ROWNUM:= 0 ) = 0 ");
            //strSql.AppendLine("       ) Y1 ");
            #endregion

            strSql.AppendLine(" SELECT*                                                                 ");
            strSql.AppendLine("    FROM(                                                                ");
            strSql.AppendLine("  SELECT  ROW_NUMBER() OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) AS SEQ");
            strSql.AppendLine("       , CONVERT(DATE, A1.TDATE) AS TDATE                                ");
            strSql.AppendLine("       , A1.SEQNO                                                        ");
            strSql.AppendLine("       , A1.LINNO                                                        ");
            strSql.AppendLine("       , A1.ACCOD                                                        ");
            strSql.AppendLine("       , B1.ACNAM                                                        ");
            strSql.AppendLine("       , A1.ATEXT AS ATEXT                                               ");
            strSql.AppendLine("       , A1.CVCOD AS CVCOD                                               ");
            strSql.AppendLine("       , A1.CVNAM AS CVNAM                                               ");
            strSql.AppendLine("       , CASE WHEN LEN(C.IDT_NO) = 10 THEN CONCAT(SUBSTRING(C.IDT_NO, 1, 3), '-', SUBSTRING(C.IDT_NO, 4, 2), '-', SUBSTRING(C.IDT_NO, 6, 5))");
            strSql.AppendLine("                                         ELSE C.IDT_NO END AS IDT_NO");
            strSql.AppendLine("       , B1.ACRDR                                                   ");
            strSql.AppendLine("       , FORMAT(ISNULL(A1.ACAMT, 0), N'#,0') AS ACAMT               ");
            strSql.AppendLine("       , FORMAT(ISNULL(A1.ADAMT, 0), N'#,0') AS ADAMT               ");
            strSql.AppendLine("       , 0 AS JAMT                                                  ");
            strSql.AppendLine("    FROM ACTRAN A1                                                  ");
            strSql.AppendLine("    LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                      ");
            strSql.AppendLine("    LEFT JOIN ACC_DEALER_CD C                                       ");
            strSql.AppendLine("      ON A1.CVCOD = C.DEALER_CD                                     ");
            strSql.AppendLine("   WHERE A1.TDATE BETWEEN '" + sYmd + "' AND '" + sYmd + "'  ");
            strSql.AppendLine("     AND A1.ATGUB IN ( '" + sGB + "' ) ");
            strSql.AppendLine("     AND A1.ACCOD <> '" + sCashCd + "'  ");
            strSql.AppendLine("        ) Y1                                                        ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetInfoSummary(string sGB, string sYmd)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sCashCd = ComnEtcFunc.CashCode;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT FORMAT(SUM(IFNULL(A1.ACAMT, 0)), 0) AS ACAMT  ");
            //strSql.AppendLine("      , FORMAT(SUM(IFNULL(A1.ADAMT, 0)), 0) AS ADAMT  ");
            //strSql.AppendLine("   FROM ACTRAN A1  ");
            //strSql.AppendLine("   LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD  ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C ");
            //strSql.AppendLine("     ON A1.CVCOD = C.DEALER_CD ");
            //strSql.AppendLine("  WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //strSql.AppendLine("    AND A1.ATGUB IN ( '" + sGB + "' ) ");
            //strSql.AppendLine("    AND A1.ACCOD <> '" + sCashCd + "'  ");
            #endregion

            strSql.AppendLine(" SELECT FORMAT(SUM(ISNULL(A1.ACAMT, 0)), N'#,0') AS ACAMT ");
            strSql.AppendLine("      , FORMAT(SUM(ISNULL(A1.ADAMT, 0)), N'#,0')AS ADAMT  ");
            strSql.AppendLine("   FROM ACTRAN A1                                         ");
            strSql.AppendLine("   LEFT JOIN ACMSTF B1                                    ");
            strSql.AppendLine("     ON A1.ACCOD = B1.ACCOD                               ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C                              ");
            strSql.AppendLine("     ON A1.CVCOD = C.DEALER_CD                            ");
            strSql.AppendLine("  WHERE A1.TDATE BETWEEN '" + sYmd + "' AND '" + sYmd + "'  ");
            strSql.AppendLine("    AND A1.ATGUB IN ( '" + sGB + "' ) ");
            strSql.AppendLine("    AND A1.ACCOD <> '" + sCashCd + "'  ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //입출금액 중 0 이상인 것에 따라 DropDown버튼 자동세팅
        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sSEQ = GridViewRetr.GetFocusedRowCellValue(GridColSEQ)?.ToString() ?? string.Empty;
            if (sSEQ.Equals("2"))
                DropBtnPrint.Enabled = true;
            else
                DropBtnPrint.Enabled = false;

            int iDeposit = 0;
            int iWithDraw = 0;
            int.TryParse(GridViewRetr.GetFocusedRowCellValue(GridColDeposit)?.ToString(), out iDeposit);
            int.TryParse(GridViewRetr.GetFocusedRowCellValue(GridColWithDraw)?.ToString(), out iWithDraw);
            if(iDeposit > 0)
            {
                UpdateDropDownButton(BtnSlipIn);
            }
            else if(iWithDraw > 0)
            {
                UpdateDropDownButton(BtnSlipOut);
            }
        }

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}