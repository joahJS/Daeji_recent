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
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using ComLib;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Helpers;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Data.SqlClient;
using DevExpress.XtraGrid.Views.BandedGrid;

namespace AccAdm
{
    public partial class PurTeamRecord : DevExpress.XtraEditors.XtraForm
    {
        public PurTeamRecord()
        {
            InitializeComponent();
        }

        private void PurTeamRecord_Load(object sender, EventArgs e)
        {

            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);

            SetLoadFormLayout();
            SetBaseData(DateTo.EditValue?.ToString().Substring(0,10));
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { gridView1, gridView2, gridView3, gridView4, gridView5 };
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

        //기초자료 세팅 인센티브, 어음할인율, 어음이자, 페기물처리비, 회수율
        private void SetBaseData(string sToDate)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" DECLARE @INCEN FLOAT = NULL, @UMHIN FLOAT = NULL, @PEGIMUL FLOAT = NULL, @UMIJA FLOAT = NULL                                                     ");
            strSql.AppendLine("       , @H1 FLOAT = NULL, @H2 FLOAT = NULL, @H3 FLOAT = NULL,@H4 FLOAT = NULL,@H5 FLOAT = NULL,@H6 FLOAT = NULL,@H7 FLOAT = NULL,@H8 FLOAT = NULL,@H9 FLOAT = NULL ");

            strSql.AppendLine("  SELECT @INCEN = DATAVALUE                                                                                                          ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                                                 ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '인센티브' AND J_DATE <= '" + sToDate + "')");
            strSql.AppendLine("     AND DAEGUBUN = '기초자료'                                                                                                       ");
            strSql.AppendLine("     AND GUBUN = '인센티브'                                                                                                          ");
                                 
            strSql.AppendLine("  SELECT @UMHIN = DATAVALUE                                                                                                          ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                                                 ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '어음할인' AND J_DATE <= '" + sToDate + "')");
		    strSql.AppendLine("     AND DAEGUBUN = '기초자료'                                                                                                       ");
            strSql.AppendLine("     AND GUBUN = '어음할인'                                                                                                          ");
                                 
            strSql.AppendLine("  SELECT @UMIJA = DATAVALUE                                                                                                          ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                                                 ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '어음이자' AND J_DATE <= '" + sToDate + "')");
            strSql.AppendLine("     AND DAEGUBUN = '기초자료'                                                                                                       ");
            strSql.AppendLine("     AND GUBUN = '어음이자'                                                                                                          ");
                                 
            strSql.AppendLine("  SELECT @PEGIMUL = DATAVALUE                                                                                                          ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                                                 ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '폐기물처리비' AND J_DATE <= '" + sToDate + "')");
            strSql.AppendLine("     AND DAEGUBUN = '기초자료'                                                                                                       ");
            strSql.AppendLine("     AND GUBUN = '폐기물처리비'                                                                                                          ");

            strSql.AppendLine("  SELECT @H1 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재특급A' AND J_DATE <= '" + sToDate + "')       ");
           	strSql.AppendLine("  AND GUBUN = '모재특급A'                                                                                          ");
                                 
            strSql.AppendLine("  SELECT @H2 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재특급B' AND J_DATE <= '" + sToDate + "')       ");
           	strSql.AppendLine("  AND GUBUN = '모재특급B'                                                                                          ");
                                 
            strSql.AppendLine("  SELECT @H3 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재A' AND J_DATE <= '" + sToDate + "')           ");
           	strSql.AppendLine("  AND GUBUN = '모재A'                                                                                              ");
                                 
            strSql.AppendLine("  SELECT @H4 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재B' AND J_DATE <= '" + sToDate + "')           ");
           	strSql.AppendLine("  AND GUBUN = '모재B'                                                                                              ");
                                 
            strSql.AppendLine("  SELECT @H5 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재C' AND J_DATE <= '" + sToDate + "')           ");
           	strSql.AppendLine("  AND GUBUN = '모재C'                                                                                              ");
                                 
            strSql.AppendLine("  SELECT @H6 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '슈레더C' AND J_DATE <= '" + sToDate + "')         ");
           	strSql.AppendLine("  AND GUBUN = '슈레더C'                                                                                            ");
                                 
            strSql.AppendLine("  SELECT @H7 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재AA' AND J_DATE <= '" + sToDate + "')          ");
           	strSql.AppendLine("  AND GUBUN = '모재AA'                                                                                             ");
                                 
            strSql.AppendLine("  SELECT @H8 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '슈레더A' AND J_DATE <= '" + sToDate + "')         ");
           	strSql.AppendLine("  AND GUBUN = '슈레더A'                                                                                            ");
                                 
            strSql.AppendLine("  SELECT @H9 = RECOVERY                                                                                            ");
            strSql.AppendLine("    FROM MEAIPSILJUK                                                                                               ");
            strSql.AppendLine("   WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = '모재P' AND J_DATE <= '" + sToDate + "')           ");
            strSql.AppendLine("  AND GUBUN = '모재P'                                                                                              ");

            strSql.AppendLine(" SELECT @INCEN AS INCEN                                                                                                    ");
            strSql.AppendLine("      , @UMHIN AS UMHIN                                                                                                    ");
            strSql.AppendLine("      , @UMIJA AS UMIJA                                                                                                    ");
            strSql.AppendLine("      , @PEGIMUL AS PEGIMUL");
            strSql.AppendLine("      , @H1 AS H1 ");
            strSql.AppendLine("      , @H2 AS H2 ");
            strSql.AppendLine("      , @H3 AS H3 ");
            strSql.AppendLine("      , @H4 AS H4 ");
            strSql.AppendLine("      , @H5 AS H5 ");
            strSql.AppendLine("      , @H6 AS H6 ");
            strSql.AppendLine("      , @H7 AS H7 ");
            strSql.AppendLine("      , @H8 AS H8 ");
            strSql.AppendLine("      , @H9 AS H9 ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt != null)
            {
                TxtIncen.EditValue = dt.Rows[0]["INCEN"]?.ToString();
                TxtAumH.EditValue = dt.Rows[0]["UMHIN"]?.ToString();
                TxtUmIja.EditValue = dt.Rows[0]["UMIJA"]?.ToString();
                TxtPagimul.EditValue = dt.Rows[0]["PEGIMUL"]?.ToString();
                TxtH1.EditValue = dt.Rows[0]["H1"]?.ToString();
                TxtH2.EditValue = dt.Rows[0]["H2"]?.ToString();
                TxtH3.EditValue = dt.Rows[0]["H3"]?.ToString();
                TxtH4.EditValue = dt.Rows[0]["H4"]?.ToString();
                TxtH5.EditValue = dt.Rows[0]["H5"]?.ToString();
                TxtH6.EditValue = dt.Rows[0]["H6"]?.ToString();
                TxtH7.EditValue = dt.Rows[0]["H7"]?.ToString();
                TxtH8.EditValue = dt.Rows[0]["H8"]?.ToString();
                TxtH9.EditValue = dt.Rows[0]["H9"]?.ToString();
            }
            else
            {
                TxtIncen.EditValue = null;
                TxtAumH.EditValue  = null;
                TxtUmIja.EditValue = null;
                TxtPagimul.EditValue = null;
                TxtH1.EditValue = null;
                TxtH2.EditValue = null;
                TxtH3.EditValue = null;
                TxtH4.EditValue = null;
                TxtH5.EditValue = null;
                TxtH6.EditValue = null;
                TxtH7.EditValue = null;
                TxtH8.EditValue = null;
                TxtH9.EditValue = null;
            }

            ResetEditYn();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sGb = "조회";
            GetRetrData(sGb);
        }

        private void GetRetrData(string sGb)
        {
            //string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);
            //string sFromDate = sDate.Substring(0, 7) + "-01";
            //string sYM = sDate.Substring(0, 7).Replace("-", "");

            string sFromDate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sToDate = DateTo.EditValue?.ToString().Substring(0, 10);
            string sYM = sFromDate.Substring(0, 7).Replace("-", "");

            SetBaseData(sToDate);

            string sIncen = TxtIncen.EditValue?.ToString();
            string sUmhin = TxtAumH.EditValue?.ToString();
            string sUmija = TxtUmIja.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            if (sGb.Equals("조회"))
            {
                #region 매출평균단가
                strSql.Clear();
                strSql.AppendLine(" SET ANSI_WARNINGS OFF");
                strSql.AppendLine(" SET ARITHIGNORE ON   ");
                strSql.AppendLine(" SET ARITHABORT OFF;  ");
                strSql.AppendLine(" DECLARE @AVDAN1 FLOAT, @AVDAN2 FLOAT, @AVDAN3 FLOAT                                                 ");
                strSql.AppendLine(" --매출단가 중량                                                                                     ");
                strSql.AppendLine(" SELECT @AVDAN1 = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))-- 매출평균단가");
                strSql.AppendLine("         FROM INLIST A, JAJAE B                                                                      ");
                strSql.AppendLine("     WHERE A.J_SERIAL = B.J_SERIAL                                                                   ");
                strSql.AppendLine("     AND A.KERATYPE = '매출'                                                                         ");
                strSql.AppendLine("     AND A.J_DATE BETWEEN'" + sFromDate + "' AND '" + sToDate + "'                                               ");
                strSql.AppendLine("     AND A.J_LOTNO <> '4'                                                                            ");
                strSql.AppendLine("     AND B.DAEGUBUN = '고철A'                                                                        ");
                strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                                                                      ");
                strSql.AppendLine(" --경량매출단가 구매팀실적 매출 와이케이스틸(코드:6531121044) j_lotno: 직납구분                      ");
                strSql.AppendLine(" SELECT @AVDAN2 = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))-- 매출평균단가");
                strSql.AppendLine("         FROM INLIST A, JAJAE B                                                                      ");
                strSql.AppendLine("     WHERE A.J_SERIAL = B.J_SERIAL                                                                   ");
                strSql.AppendLine("     AND A.KERATYPE = '매출'                                                                         ");
                strSql.AppendLine("     AND A.J_DATE BETWEEN '" + sFromDate + "' AND '" + sToDate + "'                                              ");
                strSql.AppendLine("     AND A.J_LOTNO <> '4'                                                                            ");
                strSql.AppendLine("     AND B.DAEGUBUN = '고철B'                                                                        ");
                strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                                                                      ");
                strSql.AppendLine(" --슈레더                                                                                            ");
                strSql.AppendLine(" SELECT @AVDAN3 = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))-- 매출평균단가");
                strSql.AppendLine("         FROM INLIST A, JAJAE B                                                                      ");
                strSql.AppendLine("     WHERE A.J_SERIAL = B.J_SERIAL                                                                   ");
                strSql.AppendLine("     AND A.KERATYPE = '매출'                                                                         ");
                strSql.AppendLine("     AND A.J_DATE BETWEEN '" + sFromDate + "' AND '" + sToDate + "'                                              ");
                strSql.AppendLine("     AND A.J_LOTNO <> '4'                                                                            ");
                strSql.AppendLine("     AND A.J_ID1 = '6531121044'                                                                      ");
                strSql.AppendLine("     AND B.DAEGUBUN = '슈레더'                                                                       ");
                strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                                                                      ");
                strSql.AppendLine("                                                                                                     ");
                strSql.AppendLine(" SELECT @AVDAN1 AS AVDAN1                                                                            ");
                strSql.AppendLine("     , @AVDAN2 AS AVDAN2                                                                             ");
                strSql.AppendLine("     , @AVDAN3 AS AVDAN3                                                                             ");

                DataTable dtAvDan = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dtAvDan != null)
                {
                    TeditMeachul.EditValue = dtAvDan.Rows[0]["AVDAN1"]?.ToString();
                    TeditMeachul1.EditValue = dtAvDan.Rows[0]["AVDAN2"]?.ToString();
                    TeditMeachul2.EditValue = dtAvDan.Rows[0]["AVDAN3"]?.ToString();
                }
                else
                {
                    TeditMeachul.EditValue = null;
                    TeditMeachul1.EditValue = null;
                    TeditMeachul2.EditValue = null;
                }
                #endregion
            }

            double dAvDan1 = 0; //중량매출단가
            double dAvDan2 = 0; //경량매출단가
            double dAvDan3 = 0; //슈레더매출단가
            double dPDan = 0; //폐기물처리비

            string sPDan = TxtPagimul.EditValue?.ToString();//폐기물
            double.TryParse(sPDan, out dPDan);

            //매출단가
            string sAvDan1 = TeditMeachul.EditValue?.ToString();//중량
            string sAvDan2 = TeditMeachul1.EditValue?.ToString();//경량
            string sAvDan3 = TeditMeachul2.EditValue?.ToString();//슈레더

            double.TryParse(sAvDan1, out dAvDan1);
            double.TryParse(sAvDan2, out dAvDan2);
            double.TryParse(sAvDan3, out dAvDan3);

            #region 담당자별 스크랩 중량 실적
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" DECLARE @AVDAN1 FLOAT = "+ dAvDan1 + ";");
            strSql.AppendLine(" --담당자별 스크랩 중량 실적              ");
            strSql.AppendLine(" WITH TEMP1 AS(                           ");
            strSql.AppendLine("     SELECT '손상영' AS NAME              ");
            strSql.AppendLine("      UNION ALL                           ");
            strSql.AppendLine("     SELECT '이우택'                      ");
            strSql.AppendLine("      UNION ALL                           ");
            strSql.AppendLine("     SELECT '박봉섭'                      ");
            strSql.AppendLine("      UNION ALL                           ");
            strSql.AppendLine("     SELECT '오상훈'                      ");
            strSql.AppendLine("      UNION ALL                           ");
            strSql.AppendLine("     SELECT '김명철'                      ");
            strSql.AppendLine(" ), TEMP2 AS(                             ");
            strSql.AppendLine("     --마진 : 매출평균단가 - 도착매입단가 매출이익: 중량* 마진 ");
            strSql.AppendLine("     SELECT A1.NAME                                            ");
            strSql.AppendLine("          , A2.DANJUNG                                         ");
            strSql.AppendLine("          , A2.HALIN                                           ");
            strSql.AppendLine("          , A2.DOAVDAN                                         ");
            strSql.AppendLine("          , CASE WHEN A2.DOAVDAN IS NULL THEN NULL ELSE @AVDAN1 - A2.DOAVDAN END AS MAGIN");
            strSql.AppendLine("          , CASE WHEN A2.DANJUNG IS NULL THEN NULL ELSE A2.DANJUNG * (@AVDAN1 - A2.DOAVDAN) END AS MCHEIC");
            strSql.AppendLine("       FROM TEMP1 A1                                      ");
            strSql.AppendLine("       LEFT JOIN(SELECT E.EMP_NM                          ");
            strSql.AppendLine("                       , SUM(A.DANJUNG) AS DANJUNG --중량 ");
            strSql.AppendLine("                       , SUM(A.HALIN) AS HALIN--감량      ");
            strSql.AppendLine("                       , SUM(A.IKONGKEP) AS IKONGKEP      ");
            strSql.AppendLine("                       , SUM(A.CKONGKEP) AS CKONGKEP      ");
            strSql.AppendLine("                       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ SUM(A.DANJUNG) AS DOAVDAN--도착매입평균단가(반올림)");
            strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E");
            strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                           ");
            strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                             ");
            strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                              ");
            strSql.AppendLine("                     AND A.KERATYPE = '매입'                               ");
            strSql.AppendLine("                     AND A.J_DATE BETWEEN '" + sFromDate + "' AND '" + sToDate + "'    ");
            strSql.AppendLine("                     AND A.J_LOTNO <> '4'                                  ");
            strSql.AppendLine("                     AND B.DAEGUBUN = '고철A'                              ");
            strSql.AppendLine("                     AND E.EMP_NM NOT IN('김홍철', '기타')                 ");
            strSql.AppendLine("                   GROUP BY E.EMP_NM ) A2                                  ");
            strSql.AppendLine("         ON A1.NAME = A2.EMP_NM                                            ");
            strSql.AppendLine(" )                                                                         ");
            strSql.AppendLine("                                                                           ");
            strSql.AppendLine(" SELECT NAME                                                               ");
            strSql.AppendLine("      , DANJUNG                                                            ");
            strSql.AppendLine("      , HALIN                                                              ");
            strSql.AppendLine("      , DOAVDAN                                                            ");
            strSql.AppendLine("      , MAGIN                                                              ");
            strSql.AppendLine("      , MCHEIC                                                             ");
            strSql.AppendLine("   FROM TEMP2                                                              ");
            strSql.AppendLine("  WHERE DANJUNG IS NOT NULL");
            strSql.AppendLine("  UNION ALL                                                                ");
            strSql.AppendLine(" SELECT '합 계'                                                            ");
            strSql.AppendLine("      , SUM(DANJUNG)                                                       ");
            strSql.AppendLine("      , SUM(HALIN)                                                         ");
            strSql.AppendLine("      , NULL                                                               ");
            strSql.AppendLine("      , SUM(MCHEIC) / SUM(DANJUNG)                                         ");
            strSql.AppendLine("      , SUM(MCHEIC)                                                        ");
            strSql.AppendLine("   FROM TEMP2                                                              ");

            DataTable dtScrapA = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl1.DataSource = dtScrapA;

            #endregion

            #region 담당자별 스크랩 경량 실적
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" DECLARE @AVDAN2 FLOAT = "+ dAvDan2 + ";                                                    ");
            strSql.AppendLine(" --담당자별 스크랩 경량 실적                                                                  ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                                               ");
            strSql.AppendLine("        SELECT '손상영' AS NAME                                                               ");
            strSql.AppendLine("         UNION ALL                                                                            ");
            strSql.AppendLine("        SELECT '이우택'                                                                       ");
            strSql.AppendLine("         UNION ALL                                                                            ");
            strSql.AppendLine("        SELECT '박봉섭'                                                                       ");
            strSql.AppendLine("         UNION ALL                                                                            ");
            strSql.AppendLine("        SELECT '오상훈'                                                                       ");
            strSql.AppendLine("         UNION ALL                                                                            ");
            strSql.AppendLine("        SELECT '김명철'                                                                       ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                                 ");
            strSql.AppendLine("         SELECT A1.NAME                                                                       ");
            strSql.AppendLine("              , CASE WHEN A2.GUBUN1 IS NULL THEN '합 계' ELSE A2.GUBUN1 END GUBUN1            ");
            strSql.AppendLine("              , A2.DANJUNG                                                                    ");
            strSql.AppendLine("              , A2.HALIN                                                                      ");
            strSql.AppendLine("              , A2.DOAVDAN                                                                    ");
            strSql.AppendLine("              , CASE WHEN A2.DOAVDAN IS NULL THEN NULL ELSE @AVDAN2 - A2.DOAVDAN END AS MAGIN ");
            strSql.AppendLine("              , CASE WHEN A2.DANJUNG IS NULL THEN NULL ELSE A2.DANJUNG * (@AVDAN2 - A2.DOAVDAN) END AS MCHEIC");
            strSql.AppendLine("              , A2.IKONGKEP                                                                               ");
            strSql.AppendLine("           FROM TEMP1 A1                                                                                  ");
            strSql.AppendLine("           LEFT JOIN(SELECT E.EMP_NM                                                                      ");
            strSql.AppendLine("                          , B.GUBUN1                                                                      ");
            strSql.AppendLine("                          , SUM(A.DANJUNG) AS DANJUNG                                                     ");
            strSql.AppendLine("                          , SUM(A.HALIN) AS HALIN                                                         ");
            strSql.AppendLine("                          , SUM(A.IKONGKEP) AS IKONGKEP                                                   ");
            strSql.AppendLine("                          , SUM(A.CKONGKEP) AS CKONGKEP                                                   ");
            strSql.AppendLine("                          , (SUM(A.IKONGKEP)+SUM(A.CKONGKEP))/ SUM(A.DANJUNG) AS DOAVDAN--도착매입평균단가");
            strSql.AppendLine("                       FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                            ");
            strSql.AppendLine("                      WHERE A.J_SERIAL = B.J_SERIAL                                                       ");
            strSql.AppendLine("                        AND A.J_ID1 = D.DEALER_CD                                                         ");
            strSql.AppendLine("                        AND D.CHRG_ID = E.EMP_ID                                                          ");
            strSql.AppendLine("                        AND A.KERATYPE = '매입'                                                           ");
            strSql.AppendLine("                        AND A.J_DATE BETWEEN '" + sFromDate + "' AND '" + sToDate + "'                                ");
            strSql.AppendLine("                        AND A.J_LOTNO <> '4'                                                              ");
            strSql.AppendLine("                        AND B.DAEGUBUN = '고철B'                                                          ");
            strSql.AppendLine("                        AND E.EMP_NM NOT IN('기타', '김홍철')                                             ");
            strSql.AppendLine("                      GROUP BY E.EMP_NM, B.GUBUN1                                                         ");
            strSql.AppendLine("                       WITH ROLLUP ) A2                                                                   ");
            strSql.AppendLine("             ON A1.NAME = A2.EMP_NM                                                                       ");
            strSql.AppendLine(" )                                                                                                        ");
            strSql.AppendLine(" SELECT NAME                                                                                              ");
            strSql.AppendLine("      , GUBUN1                                                                                            ");
            strSql.AppendLine("      , DANJUNG                                                                                           ");
            strSql.AppendLine("      , HALIN                                                                                             ");
            strSql.AppendLine("      , DOAVDAN                                                                                           ");
            strSql.AppendLine("      , MAGIN                                                                                             ");
            strSql.AppendLine("      , MCHEIC                                                                                            ");
            strSql.AppendLine("   FROM TEMP2                                                                                             ");
            strSql.AppendLine("  WHERE DANJUNG IS NOT NULL");
            strSql.AppendLine("  UNION ALL                                                                                               ");
            strSql.AppendLine(" SELECT '총 합 계'                                                                                        ");
            strSql.AppendLine("      , NULL                                                                                              ");
            strSql.AppendLine("      , SUM(DANJUNG)                                                                                      ");
            strSql.AppendLine("      , SUM(HALIN)                                                                                        ");
            strSql.AppendLine("      , NULL                                                                                              ");
            strSql.AppendLine("      , SUM(MCHEIC) / SUM(DANJUNG)                                                                        ");
            strSql.AppendLine("      , SUM(MCHEIC)                                                                                       ");
            strSql.AppendLine("   FROM TEMP2                                                                                             ");
            strSql.AppendLine("  WHERE GUBUN1 = '합 계'                                                                                  ");

            DataTable dtScrapKeong = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl3.DataSource = dtScrapKeong;

            #endregion

            #region 담당자별 슈레더 실적
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" DECLARE @AVDAN3 FLOAT = "+ dAvDan3 + ", @PEGIMUL FLOAT = "+ dPDan + ";");
            strSql.AppendLine(" WITH TEMP1 AS(                         ");
            strSql.AppendLine("     SELECT 1 AS NUM, '손상영' AS NAME  ");
            strSql.AppendLine("      UNION ALL                         ");
            strSql.AppendLine("     SELECT 2 AS NUM, '이우택'          ");
            strSql.AppendLine("      UNION ALL                         ");
            strSql.AppendLine("     SELECT 3 AS NUM, '박봉섭'          ");
            strSql.AppendLine("      UNION ALL                         ");
            strSql.AppendLine("     SELECT 4 AS NUM, '오상훈'          ");
            strSql.AppendLine("      UNION ALL                         ");
            strSql.AppendLine("     SELECT 5 AS NUM, '김명철'          ");
            strSql.AppendLine(" ), TEMP2 AS(                            ");
            strSql.AppendLine("     SELECT 1 AS NUM                     ");
            strSql.AppendLine("          , E.EMP_NM AS NAME                    ");
            strSql.AppendLine("          , B.GUBUN1                     ");
            strSql.AppendLine("          , SUM(A.DANJUNG) AS DANJUNG    ");
            strSql.AppendLine("          , SUM(A.HALIN) AS HALIN        ");
            strSql.AppendLine("          , SUM(A.IKONGKEP) AS IKONGKEP  ");
            strSql.AppendLine("          , SUM(A.CKONGKEP) AS CKONGKEP  ");
            strSql.AppendLine("          , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ SUM(A.DANJUNG) AS MIPDAN ");
            strSql.AppendLine("          , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) AS DOAMT                  ");
            strSql.AppendLine("          --((중량 * 회수율) * 매출평균단가) - (중량 * (((금액 + 운반비) / 중량) + 36.6)) - (중량 - (중량 * 회수율)) * 페기물처리비");
            strSql.AppendLine("          , ((SUM(A.DANJUNG) * MAX(B1.RECOVERY)) * @AVDAN3)                                                                        ");
            strSql.AppendLine("            - (SUM(A.DANJUNG) * (((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) + 36.6))                                   ");
            strSql.AppendLine("            - (SUM(A.DANJUNG) - (SUM(A.DANJUNG) * MAX(B1.RECOVERY))) * @PEGIMUL AS MACHEIC                                         ");
            strSql.AppendLine("       FROM INLIST A                                                                                                               ");
            strSql.AppendLine("       LEFT JOIN JAJAE B                                                                                                           ");
            strSql.AppendLine("         ON A.J_SERIAL = B.J_SERIAL                                                                                                ");
            strSql.AppendLine("        AND B.DAEGUBUN = '슈레더'                                                                                                  ");
            strSql.AppendLine("      LEFT JOIN MEAIPSILJUK B1                                                                                                     ");
            strSql.AppendLine("         ON B.GUBUN1 = B1.GUBUN                                                                                                    ");
            strSql.AppendLine("        AND B1.DAEGUBUN = '슈레더'                                                                                                 ");
            strSql.AppendLine("        AND B1.J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '슈레더' AND J_DATE <= '" + sToDate + "')");
            strSql.AppendLine("       LEFT JOIN ACC_DEALER_CD D                                                                                                   ");
            strSql.AppendLine("         ON A.J_ID1 = D.DEALER_CD                                                                                                  ");
            strSql.AppendLine("       LEFT JOIN HR_EMP_BASIS E                                                                                                    ");
            strSql.AppendLine("         ON D.CHRG_ID = E.EMP_ID                                                                                                   ");
            strSql.AppendLine("      WHERE A.KERATYPE = '매입'                                                                                                    ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                                                         ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                                                                       ");
            strSql.AppendLine("        AND E.EMP_NM NOT IN ('기타','김홍철')                                                                     ");
            strSql.AppendLine("        AND B.GUBUN1 IN('모재특급A','모재특급B','모재A','모재B','모재C','슈레더C','모재AA')                                        ");
            strSql.AppendLine("      GROUP BY E.EMP_NM, B.GUBUN1                                                                                                  ");
            strSql.AppendLine(" )                                                                                                                                 ");
            strSql.AppendLine("                                                                                                                                   ");
            strSql.AppendLine(" SELECT T1.NUM                                                                                                                     ");
            strSql.AppendLine("      , Z1.*                                                                                                                       ");
            strSql.AppendLine("   FROM ( SELECT NUM                                                                                                               ");
            strSql.AppendLine("               , NAME                                                                                                            ");
            strSql.AppendLine("               , GUBUN1                                                                                                            ");
            strSql.AppendLine("               , DANJUNG                                                                                                           ");
            strSql.AppendLine("               , MIPDAN                                                                                                            ");
            strSql.AppendLine("               , DOAMT                                                                                                             ");
            strSql.AppendLine("               , MACHEIC                                                                                                           ");
            strSql.AppendLine("            FROM TEMP2                                                                                                             ");
            strSql.AppendLine("           UNION ALL                                                                                                               ");
            strSql.AppendLine("          SELECT 99                                                                                                                ");
            strSql.AppendLine("               , NAME                                                                                                            ");
            strSql.AppendLine("               , '합 계'                                                                                                           ");
            strSql.AppendLine("               , SUM(DANJUNG)                                                                                                      ");
            strSql.AppendLine("               , SUM(DOAMT) / SUM(DANJUNG)                                                                                         ");
            strSql.AppendLine("               , SUM(DOAMT)                                                                                                        ");
            strSql.AppendLine("               , SUM(MACHEIC)                                                                                                      ");
            strSql.AppendLine("            FROM TEMP2                                                                                                             ");
            strSql.AppendLine("           WHERE DANJUNG IS NOT NULL                                                                                               ");
            strSql.AppendLine("           GROUP BY NAME) Z1                                                                                                     ");
            strSql.AppendLine("   LEFT JOIN TEMP1 T1                                                                                                              ");
            strSql.AppendLine("     ON Z1.NAME = T1.NAME                                                                                                        ");
            strSql.AppendLine("  ORDER BY T1.NUM, Z1.NUM                                                                                                          ");

            DataTable dtShureder = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl4.DataSource = dtShureder;
            #endregion

            #region 직납 실적
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" --직납실적                                                        ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                    ");
            strSql.AppendLine("         SELECT 1 AS NUM, '손상영' AS NAME                         ");
            strSql.AppendLine("          UNION ALL                                                ");
            strSql.AppendLine("         SELECT 2, '이우택'                                        ");
            strSql.AppendLine("          UNION ALL                                                ");
            strSql.AppendLine("         SELECT 3, '박봉섭'                                        ");
            strSql.AppendLine("          UNION ALL                                                ");
            strSql.AppendLine("         SELECT 4, '오상훈'                                        ");
            strSql.AppendLine("          UNION ALL                                                ");
            strSql.AppendLine("         SELECT 5, '김명철'                                        ");
            strSql.AppendLine("  ), TEMP2 AS(                                                     ");
            strSql.AppendLine("       SELECT A1.NUM                                               ");
            strSql.AppendLine("           , A1.NAME                                               ");
            strSql.AppendLine("           , A2.DANJUNG                                            ");
            strSql.AppendLine("           , A2.J_KONGKEP                                          ");
            strSql.AppendLine("           , A2.DOAMT                                              ");
            strSql.AppendLine("           , A2.MAGIN                                              ");
            strSql.AppendLine("           , A2.MACHEIC                                            ");
            strSql.AppendLine("        FROM TEMP1 A1                                              ");
            strSql.AppendLine("        LEFT JOIN (SELECT E.EMP_NM                                 ");
            strSql.AppendLine("                         , SUM(A.DANJUNG) AS DANJUNG               ");
            strSql.AppendLine("                         , SUM(A.KONGKEP) AS KONGKEP               ");
            strSql.AppendLine("                         , SUM(A.IKONGKEP) AS IKONGKEP             ");
            strSql.AppendLine("                         , SUM(A.CKONGKEP) AS CKONGKEP             ");
            strSql.AppendLine("                         , SUM(F.J_KONGKEP) AS J_KONGKEP           ");
            strSql.AppendLine("                         , SUM(A.IKONGKEP)+SUM(A.CKONGKEP) AS DOAMT");
            strSql.AppendLine("                         , (SUM(F.J_KONGKEP) - (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))) / SUM(A.DANJUNG) AS MAGIN");
            strSql.AppendLine("                         , SUM(F.J_KONGKEP) -(SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) AS MACHEIC                  ");
            strSql.AppendLine("                      FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E, IPCHULGO F                    ");
            strSql.AppendLine("                     WHERE A.J_SERIAL = B.J_SERIAL                                                           ");
            strSql.AppendLine("                       AND A.J_ID1 = D.DEALER_CD                                                             ");
            strSql.AppendLine("                       AND D.CHRG_ID = E.EMP_ID                                                              ");
            strSql.AppendLine("                       AND A.J_ID = F.J_ID - 1                                                               ");
            strSql.AppendLine("                       AND A.KERATYPE = '매입'                                                               ");
            strSql.AppendLine("                       AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                    ");
            strSql.AppendLine("                       AND A.J_LOTNO = '4'                                                                   ");
            strSql.AppendLine("                       AND E.EMP_NM NOT IN('기타', '김홍철')                                                 ");
            strSql.AppendLine("                     GROUP BY E.EMP_NM ) A2                                                                  ");
            strSql.AppendLine("          ON A1.NAME = A2.EMP_NM                                                                             ");
            strSql.AppendLine("       WHERE A2.DANJUNG IS NOT NULL                                                                          ");
		    strSql.AppendLine("  )                                                                                                          ");
		    strSql.AppendLine("                                                                                                             ");
		    strSql.AppendLine("  SELECT A1.*                                                                                                ");
            strSql.AppendLine("    FROM( SELECT NUM                                                                                         ");
            strSql.AppendLine("               , NAME                                                                                        ");
            strSql.AppendLine("               , DANJUNG                                                                                     ");
            strSql.AppendLine("               , J_KONGKEP                                                                                   ");
            strSql.AppendLine("               , DOAMT                                                                                       ");
            strSql.AppendLine("               , MAGIN                                                                                       ");
            strSql.AppendLine("               , MACHEIC                                                                                     ");
            strSql.AppendLine("            FROM TEMP2                                                                                       ");
            strSql.AppendLine("           UNION ALL                                                                                         ");
            strSql.AppendLine("          SELECT 6                                                                                           ");
            strSql.AppendLine("               , '합 계'                                                                                     ");
            strSql.AppendLine("               , SUM(DANJUNG)                                                                                ");
            strSql.AppendLine("               , SUM(J_KONGKEP)                                                                              ");
            strSql.AppendLine("               , SUM(DOAMT)                                                                                  ");
            strSql.AppendLine("               , (SUM(J_KONGKEP) - SUM(DOAMT)) / SUM(DANJUNG)                                                ");
            strSql.AppendLine("               , SUM(MACHEIC)                                                                                ");
            strSql.AppendLine("            FROM TEMP2) A1                                                                                   ");
            strSql.AppendLine("  ORDER BY A1.NUM                                                                                            ");

            DataTable dtJK = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl2.DataSource = dtJK;
            #endregion

            #region 실적통계
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF                                                                                                       ");
            strSql.AppendLine(" SET ARITHIGNORE ON                                                                                                          ");
            strSql.AppendLine(" SET ARITHABORT OFF;                                                                                                         ");
            strSql.AppendLine(" DECLARE @AVDAN1 FLOAT, @AVDAN2 FLOAT, @AVDAN3 FLOAT, @PEGIMUL FLOAT                                  ");
            strSql.AppendLine(" --매출단가 중량                                                                                      ");
            strSql.AppendLine(" SELECT @AVDAN1 = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))-- 매출평균단가 ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                             ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                       ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                           ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                                 ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                              ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철A'                                                                          ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브'                                                                        ");
            strSql.AppendLine(" --경량매출단가 구매팀실적 매출 와이케이스틸(코드:6531121044) j_lotno: 직납구분                       ");
            strSql.AppendLine(" SELECT @AVDAN2 = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))-- 매출평균단가 ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                             ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                       ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                           ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                                ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                              ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철B'                                                                          ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브'                                                                        ");
            strSql.AppendLine(" --슈레더                                                                                             ");
            strSql.AppendLine(" SELECT @AVDAN3 = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))-- 매출평균단가 ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                             ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                       ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                           ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                                ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                              ");
            strSql.AppendLine("    AND A.J_ID1 = '6531121044'                                                                        ");
            strSql.AppendLine("    AND B.DAEGUBUN = '슈레더'                                                                         ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브'                                                                        ");
            strSql.AppendLine(" --폐기물처리비                                                                                       ");
            strSql.AppendLine("     IF EXISTS(SELECT DATAVALUE FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '폐기물처리비' AND J_DATE = '"+ sFromDate + "')");
            strSql.AppendLine("        BEGIN                                                                                   ");
            strSql.AppendLine("              SELECT @PEGIMUL = DATAVALUE                                                       ");
            strSql.AppendLine("                FROM MEAIPSILJUK                                                                ");
            strSql.AppendLine("               WHERE DAEGUBUN = '기초자료' AND GUBUN = '폐기물처리비' AND J_DATE = '"+ sFromDate + "' ");
            strSql.AppendLine("          END                                                                                   ");
            strSql.AppendLine("   ELSE                                                                                         ");
            strSql.AppendLine("        BEGIN                                                                                   ");
            strSql.AppendLine("              SELECT @PEGIMUL = DATAVALUE                                                       ");
            strSql.AppendLine("                FROM MEAIPSILJUK                                                                ");
            strSql.AppendLine("               WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '폐기물처리비')");
            strSql.AppendLine(" 		        AND DAEGUBUN = '기초자료' AND GUBUN = '폐기물처리비'                                                     ");
            strSql.AppendLine("          END;                                                                                                            ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                                                                           ");
            strSql.AppendLine("     --매입 월 목표량                                                                                                     ");
            strSql.AppendLine("     SELECT NAME                                                                                                          ");
            strSql.AppendLine("          , SUM(I_WEIGHT) * 1000 AS I_WEIGHT-- 중량                                                                       ");
            strSql.AppendLine("          , SUM(I_LIGHT) * 1000 AS I_LIGHT-- 경량                                                                         ");
            strSql.AppendLine("          , SUM(I_CHAPI) * 1000 AS I_CHAPI-- 폐압                                                                         ");
            strSql.AppendLine("          , SUM(I_YK) * 1000 AS I_YK-- 직납                                                                               ");
            strSql.AppendLine("       FROM SALEMAEIP                                                                                                     ");
            strSql.AppendLine("      WHERE YYMM = '"+ sYM + "'                                                                                               ");
            strSql.AppendLine("        AND NAME NOT IN('기타')                                                                                           ");
            strSql.AppendLine("       GROUP BY NAME                                                                                                      ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                                                             ");
            strSql.AppendLine("     --담당자별 스크랩 중량 마진: 매출평균단가 - 도착매입단가 매출이익: 중량* 마진                                        ");
            strSql.AppendLine("      SELECT CASE WHEN A1.NAME IS NULL THEN '합 계' ELSE A1.NAME END AS NAME                                              ");
            strSql.AppendLine("           , A2.DANJUNG                                                                                                   ");
            strSql.AppendLine(" 	      , A2.HALIN                                                                                                     ");
            strSql.AppendLine(" 	      , A2.DOAVDAN                                                                                                   ");
            strSql.AppendLine(" 	      , CASE WHEN A2.DOAVDAN IS NULL THEN NULL ELSE @AVDAN1 - A2.DOAVDAN END AS MAGIN                                ");
            strSql.AppendLine(" 	 	  , CASE WHEN A2.DANJUNG IS NULL THEN NULL ELSE A2.DANJUNG * (@AVDAN1 - A2.DOAVDAN) END AS MCHEIC                ");
            strSql.AppendLine("        FROM TEMP1 A1                                                                                                     ");
            strSql.AppendLine("        LEFT JOIN(SELECT E.EMP_NM                                                                                         ");
            strSql.AppendLine("                       , SUM(A.DANJUNG) AS DANJUNG --중량                                                                 ");
            strSql.AppendLine(" 		              , SUM(A.HALIN) AS HALIN--감량                                                                      ");
            strSql.AppendLine(" 		              , SUM(A.IKONGKEP) AS IKONGKEP                                                                      ");
            strSql.AppendLine("                       , SUM(A.CKONGKEP) AS CKONGKEP                                                                      ");
            strSql.AppendLine("                       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ SUM(A.DANJUNG) AS DOAVDAN--도착매입평균단가(반올림)         ");
            strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                                               ");
            strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                                                                          ");
            strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                                                                            ");
            strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                                                                             ");
            strSql.AppendLine("                     AND A.KERATYPE = '매입'                                                                              ");
            strSql.AppendLine("                     AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                                   ");
            strSql.AppendLine("                     AND A.J_LOTNO <> '4'                                                                                 ");
            strSql.AppendLine("                     AND B.DAEGUBUN = '고철A'                                                                             ");
            strSql.AppendLine("                     AND E.EMP_NM NOT IN('김홍철', '기타')                                                                ");
            strSql.AppendLine("                   GROUP BY E.EMP_NM WITH ROLLUP) A2                                                                      ");
            strSql.AppendLine("          ON A1.NAME = A2.EMP_NM                                                                                          ");
            strSql.AppendLine(" ), TEMP3 AS(                                                                                                             ");
            strSql.AppendLine("      --스크랩 경량                                                                                                       ");
            strSql.AppendLine("      SELECT A1.NAME                                                                                                      ");
            strSql.AppendLine("           , A2.GUBUN1                                                                                                    ");
            strSql.AppendLine("           , A2.DANJUNG                                                                                                   ");
            strSql.AppendLine("           , A2.HALIN                                                                                                     ");
            strSql.AppendLine("           , A2.DOAVDAN                                                                                                   ");
            strSql.AppendLine("           , CASE WHEN A2.DOAVDAN IS NULL THEN NULL ELSE @AVDAN2 - A2.DOAVDAN END AS MAGIN                                ");
            strSql.AppendLine("           , CASE WHEN A2.DANJUNG IS NULL THEN NULL ELSE A2.DANJUNG * (@AVDAN2 - A2.DOAVDAN) END AS MCHEIC                ");
            strSql.AppendLine("           , A2.IKONGKEP                                                                                                  ");
            strSql.AppendLine("        FROM TEMP1 A1                                                                                                     ");
            strSql.AppendLine("        LEFT JOIN(SELECT E.EMP_NM                                                                                         ");
            strSql.AppendLine("                       , CASE WHEN B.GUBUN1 IS NULL THEN '합 계' ELSE B.GUBUN1 END GUBUN1                                 ");
            strSql.AppendLine("  	  	              , SUM(A.DANJUNG) AS DANJUNG                                                                        ");
            strSql.AppendLine("                       , SUM(A.HALIN) AS HALIN                                                                            ");
            strSql.AppendLine("                       , SUM(A.IKONGKEP) AS IKONGKEP                                                                      ");
            strSql.AppendLine("                       , SUM(A.CKONGKEP) AS CKONGKEP                                                                      ");
            strSql.AppendLine("                       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ SUM(A.DANJUNG) AS DOAVDAN--도착매입평균단가                 ");
            strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                                               ");
            strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                                                                          ");
            strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                                                                            ");
            strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                                                                             ");
            strSql.AppendLine("                     AND A.KERATYPE = '매입'                                                                              ");
            strSql.AppendLine("                     AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                                   ");
            strSql.AppendLine("                     AND A.J_LOTNO <> '4'                                                                                 ");
            strSql.AppendLine("                     AND B.DAEGUBUN = '고철B'                                                                             ");
            strSql.AppendLine("                     AND E.EMP_NM NOT IN('기타', '김홍철')                                                                ");
            strSql.AppendLine("                   GROUP BY E.EMP_NM, B.GUBUN1                                                                            ");
            strSql.AppendLine("                    WITH ROLLUP ) A2                                                                                      ");
            strSql.AppendLine("          ON A1.NAME = A2.EMP_NM                                                                                          ");
            strSql.AppendLine(" ), TEMP4 AS(                                                                                                             ");
            strSql.AppendLine(" SELECT Z1.NAME                                                                  ");
            strSql.AppendLine("      , SUM(Z1.DANJUNG) AS DANJUNG                                               ");
            strSql.AppendLine("      , SUM(Z1.DOAMT)/ SUM(Z1.DANJUNG) AS MIPDAN                                 ");
            strSql.AppendLine("      , SUM(Z1.DOAMT) AS DOAMT                                                   ");
            strSql.AppendLine("      , SUM(Z1.MACHEIC) AS MACHEIC                                               ");
            strSql.AppendLine("   FROM  ( --슈레더                                                              ");
            strSql.AppendLine("           SELECT E.EMP_NM AS NAME                                               ");
            strSql.AppendLine("                 , B.GUBUN1                                                      ");
            strSql.AppendLine("                 , SUM(A.DANJUNG) AS DANJUNG                                     ");
            strSql.AppendLine("                 , SUM(A.HALIN) AS HALIN                                         ");
            strSql.AppendLine("                 , SUM(A.IKONGKEP) AS IKONGKEP                                   ");
            strSql.AppendLine("                 , SUM(A.CKONGKEP) AS CKONGKEP                                   ");
            strSql.AppendLine("                 , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MIPDAN");
            strSql.AppendLine("                 , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) AS DOAMT                  ");
            strSql.AppendLine("                 --((중량 * 회수율) * 매출평균단가) - (중량 * (((금액 + 운반비) / 중량) + 36.6)) - (중량 - (중량 * 회수율)) * 페기물처리비");
            strSql.AppendLine("                 , ((SUM(A.DANJUNG) * MAX(B1.RECOVERY)) * @AVDAN3)                                      ");
            strSql.AppendLine("                   - (SUM(A.DANJUNG) * (((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) + 36.6)) ");
            strSql.AppendLine("                   - (SUM(A.DANJUNG) - (SUM(A.DANJUNG) * MAX(B1.RECOVERY))) * @PEGIMUL AS MACHEIC       ");
            strSql.AppendLine("              FROM INLIST A                                                                             ");
            strSql.AppendLine("              LEFT JOIN JAJAE B                                                                         ");
            strSql.AppendLine("                ON A.J_SERIAL = B.J_SERIAL                                                              ");
            strSql.AppendLine("               AND B.DAEGUBUN = '슈레더'                                                                ");
            strSql.AppendLine("              LEFT JOIN MEAIPSILJUK B1                                                                  ");
            strSql.AppendLine("                ON B.GUBUN1 = B1.GUBUN                                                                  ");
            strSql.AppendLine("               AND B1.DAEGUBUN = '슈레더'                                                               ");
            strSql.AppendLine("              LEFT JOIN ACC_DEALER_CD D                                                                 ");
            strSql.AppendLine("                ON A.J_ID1 = D.DEALER_CD                                                                ");
            strSql.AppendLine("              LEFT JOIN HR_EMP_BASIS E                                                                  ");
            strSql.AppendLine("                ON D.CHRG_ID = E.EMP_ID                                                                 ");
            strSql.AppendLine("             WHERE A.KERATYPE = '매입'                                                                  ");
            strSql.AppendLine("               AND A.J_DATE BETWEEN '" + sFromDate + "' AND '" + sToDate + "'                                       ");
            strSql.AppendLine("               AND A.J_LOTNO <> '4'                                                                     ");
            strSql.AppendLine("               AND E.EMP_NM NOT IN ('기타','김홍철')                                                                     ");
            strSql.AppendLine("               AND B.GUBUN1 IN('모재특급A', '모재특급B', '모재A', '모재B', '모재C', '슈레더C', '모재AA')");
            strSql.AppendLine("             GROUP BY E.EMP_NM, B.GUBUN1) Z1                                                            ");
            strSql.AppendLine("  WHERE Z1.DANJUNG IS NOT NULL                                                                          ");
            strSql.AppendLine("  GROUP BY Z1.NAME                                                                                      ");
            strSql.AppendLine(" ), TEMP5 AS(                                                                                                   ");
            strSql.AppendLine("      --직납                                                                                                    ");
            strSql.AppendLine("      SELECT A1.NAME                                                                                            ");
            strSql.AppendLine(" 	      , A2.DANJUNG                                                                                         ");
            strSql.AppendLine(" 	      , A2.J_KONGKEP                                                                                       ");
            strSql.AppendLine(" 	      , A2.DOAMT                                                                                           ");
            strSql.AppendLine(" 	      , A2.MAGIN                                                                                           ");
            strSql.AppendLine(" 	      , A2.MACHEIC                                                                                         ");
            strSql.AppendLine("        FROM TEMP1 A1                                                                                           ");
            strSql.AppendLine("        LEFT JOIN(SELECT E.EMP_NM                                                                               ");
            strSql.AppendLine("                       , SUM(A.DANJUNG) AS DANJUNG                                                              ");
            strSql.AppendLine("                       , SUM(A.KONGKEP) AS KONGKEP                                                              ");
            strSql.AppendLine("                       , SUM(A.IKONGKEP) AS IKONGKEP                                                            ");
            strSql.AppendLine("                       , SUM(A.CKONGKEP) AS CKONGKEP                                                            ");
            strSql.AppendLine("                       , SUM(F.J_KONGKEP) AS J_KONGKEP                                                          ");
            strSql.AppendLine("                       , SUM(A.IKONGKEP)+SUM(A.CKONGKEP) AS DOAMT                                               ");
            strSql.AppendLine("                       , (SUM(F.J_KONGKEP) - (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))) / SUM(A.DANJUNG) AS MAGIN     ");
            strSql.AppendLine("                       , SUM(F.J_KONGKEP) -(SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) AS MACHEIC                       ");
            strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E, IPCHULGO F                         ");
            strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                                                                ");
            strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                                                                  ");
            strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                                                                   ");
            strSql.AppendLine("                     AND A.J_ID = F.J_ID - 1                                                                    ");
            strSql.AppendLine("                     AND A.KERATYPE = '매입'                                                                    ");
            strSql.AppendLine("                     AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sToDate + "'                                         ");
            strSql.AppendLine("                     AND A.J_LOTNO = '4'                                                                        ");
            strSql.AppendLine("                     AND E.EMP_NM NOT IN('기타', '김홍철')                                                      ");
            strSql.AppendLine("                   GROUP BY E.EMP_NM ) A2                                                                       ");
            strSql.AppendLine("          ON A1.NAME = A2.EMP_NM                                                                                ");
            strSql.AppendLine("       WHERE A2.DANJUNG IS NOT NULL                                                                             ");
            strSql.AppendLine(" ), TEMP6 AS(                                                                                                   ");
            strSql.AppendLine("     SELECT A1.NAME                                                                                             ");
            strSql.AppendLine("          , A1.I_WEIGHT --중량목표                                                                              ");
            strSql.AppendLine(" 		 , A1.I_LIGHT-- 경량목표                                                                               ");
            strSql.AppendLine(" 		 , A1.I_CHAPI-- 폐압목표                                                                               ");
            strSql.AppendLine(" 		 , A1.I_YK-- 직납목표                                                                                  ");
            strSql.AppendLine(" 		 , A2.DANJUNG AS W1--중량실적                                                                          ");
            strSql.AppendLine(" 		 , A3.DANJUNG AS W2--경량실적                                                                          ");
            strSql.AppendLine(" 		 , A4.DANJUNG AS W3--폐압실적                                                                          ");
            strSql.AppendLine(" 		 , A5.DANJUNG AS W4--직납실적                                                                          ");
            strSql.AppendLine(" 		 , A2.MAGIN AS M1--중량마진                                                                            ");
            strSql.AppendLine(" 		 , @AVDAN2 - A3.DOAVDAN AS M2 --경량마진                                                               ");
            strSql.AppendLine(" 		 , A4.MACHEIC / A4.DANJUNG AS M3 --폐압마진                                                            ");
            strSql.AppendLine(" 		 , A5.MAGIN AS M4--직납마진                                                                            ");
            strSql.AppendLine(" 		 , A2.DANJUNG / A1.I_WEIGHT AS WP1 --중량달성율                                                        ");
            strSql.AppendLine(" 		 , A3.DANJUNG / A1.I_LIGHT AS WP2 --경량달성율                                                         ");
            strSql.AppendLine(" 		 , A4.DANJUNG / A1.I_CHAPI AS WP3 --폐압달성율                                                         ");
            strSql.AppendLine(" 		 , A5.DANJUNG / A1.I_YK AS WP4 --직납달성율                                                            ");
            strSql.AppendLine("       FROM TEMP1 A1                                                                                            ");
            strSql.AppendLine("       LEFT JOIN TEMP2 A2                                                                                       ");
            strSql.AppendLine("         ON A1.NAME = A2.NAME                                                                                   ");
            strSql.AppendLine("       LEFT JOIN TEMP3 A3                                                                                       ");
            strSql.AppendLine("         ON A1.NAME = A3.NAME                                                                                   ");
            strSql.AppendLine("        AND A3.GUBUN1 = '합 계'                                                                                 ");
            strSql.AppendLine("       LEFT JOIN TEMP4 A4                                                                                       ");
            strSql.AppendLine("         ON A1.NAME = A4.NAME                                                                                   ");
            strSql.AppendLine("       LEFT JOIN TEMP5 A5                                                                                       ");
            strSql.AppendLine("         ON A1.NAME = A5.NAME                                                                                   ");
            strSql.AppendLine(" )                                                                                                              ");
            strSql.AppendLine("                                                                                                                ");
            strSql.AppendLine(" SELECT *                                                                                                       ");
            strSql.AppendLine("      , ISNULL(I_WEIGHT, 0) + ISNULL(I_LIGHT, 0) + ISNULL(I_CHAPI, 0) AS TOTM                                                   ");
            strSql.AppendLine("      , ISNULL(W1, 0)+ISNULL(W2, 0) + ISNULL(W3, 0) AS TOTW                                                                     ");
            strSql.AppendLine("      , (ISNULL(W1, 0) + ISNULL(W2, 0) + ISNULL(W3, 0))/ (ISNULL(I_WEIGHT, 0) + ISNULL(I_LIGHT, 0) + ISNULL(I_CHAPI, 0)) AS TOTP");
            strSql.AppendLine("   FROM TEMP6                                                                                                   ");
            strSql.AppendLine("  UNION ALL                                                                                                     ");
            strSql.AppendLine(" SELECT '합계'                                                                                                  ");
            strSql.AppendLine("      , SUM(I_WEIGHT)-- 중량목표                                                                                ");
            strSql.AppendLine(" 	 , SUM(I_LIGHT)-- 경량목표                                                                                 ");
            strSql.AppendLine(" 	 , SUM(I_CHAPI)-- 폐압목표                                                                                 ");
            strSql.AppendLine(" 	 , SUM(I_YK)-- 직납목표                                                                                    ");
            strSql.AppendLine(" 	 , SUM(W1)--중량실적                                                                                       ");
            strSql.AppendLine(" 	 , SUM(W2)--경량실적                                                                                       ");
            strSql.AppendLine(" 	 , SUM(W3)--폐압실적                                                                                       ");
            strSql.AppendLine(" 	 , SUM(W4)--직납실적                                                                                       ");
            strSql.AppendLine(" 	 , (SELECT SUM(MCHEIC) FROM TEMP2)/ SUM(W1)--중량마진                                                      ");
            strSql.AppendLine(" 	 , (SELECT SUM(MCHEIC) FROM TEMP3 WHERE GUBUN1 = '합 계')/ SUM(W2)--경량마진                               ");
            strSql.AppendLine(" 	 , (SELECT SUM(MACHEIC) FROM TEMP4 WHERE NAME IN('손상영', '이우택', '박봉섭'))/ SUM(W3)--폐압마진");
            strSql.AppendLine(" 	 , (SELECT SUM(MACHEIC) FROM TEMP5)/ SUM(W4)--직납마진                                                                 ");
            strSql.AppendLine(" 	 , SUM(W1) / SUM(I_WEIGHT)--중량달성율                                                                                 ");
            strSql.AppendLine(" 	 , SUM(W2) / SUM(I_LIGHT)-- 경량달성율                                                                                 ");
            strSql.AppendLine(" 	 , SUM(W3) / SUM(I_CHAPI)-- 폐압달성율                                                                                 ");
            strSql.AppendLine(" 	 , SUM(W4) / SUM(I_YK)-- 직납달성율                                                                                    ");
            strSql.AppendLine("      , NULL ");
	        strSql.AppendLine("      , NULL ");
            strSql.AppendLine("      , NULL ");
            strSql.AppendLine("   FROM TEMP6                                                                                                               ");

            DataTable dtSilTot = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtSilTot != null)
            {
                DataTable dtTot = new DataTable();

                dtTot.Columns.Add("GUBUN");//구분

                for (int i =0;i< dtSilTot.Rows.Count; i++)
                {
                    dtTot.Columns.Add("MOK"+ (i + 1));//목표
                    dtTot.Columns.Add("SIL" + (i + 1));//실적
                    dtTot.Columns.Add("MAGIN" + (i + 1));//마진
                    dtTot.Columns.Add("YUL" + (i + 1));//달성율
                }

                DataRow row1 = dtTot.NewRow();
                row1["GUBUN"] = "스크랩 중량";

                DataRow row2 = dtTot.NewRow();
                row2["GUBUN"] = "스크랩 경량";

                DataRow row3 = dtTot.NewRow();
                row3["GUBUN"] = "폐 압";

                DataRow row4 = dtTot.NewRow();
                row4["GUBUN"] = "합 계";

                DataRow row5 = dtTot.NewRow();
                row5["GUBUN"] = "직 납";

                gridView5.Bands.Clear();

                GridBand gridBandGB = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
                BandedGridColumn bandedGridColumnGB = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();//구분

                gridView5.Bands.AddRange(new GridBand[] { gridBandGB });

                gridBandGB.AppearanceHeader.Options.UseTextOptions = true;
                gridBandGB.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridBandGB.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                gridBandGB.Caption = "구분";
                gridBandGB.Columns.Add(bandedGridColumnGB);
                gridBandGB.Name = "gridBandGB";
                gridBandGB.OptionsBand.AllowMove = false;
                gridBandGB.VisibleIndex = 0;
                gridBandGB.Width = 75;

                bandedGridColumnGB.AppearanceCell.Options.UseTextOptions = true;
                bandedGridColumnGB.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandedGridColumnGB.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                bandedGridColumnGB.FieldName = "GUBUN";
                bandedGridColumnGB.OptionsColumn.AllowEdit = false;
                bandedGridColumnGB.Visible = true;

                for (int i = 0; i < dtSilTot.Rows.Count; i++)
                {
                    string sNAME = dtSilTot.Rows[i]["NAME"]?.ToString();
                    string sI_WEIGHT = dtSilTot.Rows[i]["I_WEIGHT"]?.ToString();// 중량목표
                    string sI_LIGHT = dtSilTot.Rows[i]["I_LIGHT"]?.ToString();// 경량목표
                    string sI_CHAPI = dtSilTot.Rows[i]["I_CHAPI"]?.ToString();// 폐압목표
                    string sI_YK = dtSilTot.Rows[i]["I_YK"]?.ToString();// 직납목표
                    string sW1 = dtSilTot.Rows[i]["W1"]?.ToString();//중량실적
                    string sW2 = dtSilTot.Rows[i]["W2"]?.ToString();//경량실적
                    string sW3 = dtSilTot.Rows[i]["W3"]?.ToString();//폐압실적
                    string sW4 = dtSilTot.Rows[i]["W4"]?.ToString();//직납실적
                    string sM1 = dtSilTot.Rows[i]["M1"]?.ToString();//중량마진
                    string sM2 = dtSilTot.Rows[i]["M2"]?.ToString();//경량마진
                    string sM3 = dtSilTot.Rows[i]["M3"]?.ToString();//폐압마진
                    string sM4 = dtSilTot.Rows[i]["M4"]?.ToString();//직납마진
                    string sWP1 = dtSilTot.Rows[i]["WP1"]?.ToString();//중량달성율
                    string sWP2 = dtSilTot.Rows[i]["WP2"]?.ToString();//경량달성율
                    string sWP3 = dtSilTot.Rows[i]["WP3"]?.ToString();//폐압달성율
                    string sWP4 = dtSilTot.Rows[i]["WP4"]?.ToString();//직납달성율
                    string sTOTM = dtSilTot.Rows[i]["TOTM"]?.ToString();//목표합계
                    string sTOTW = dtSilTot.Rows[i]["TOTW"]?.ToString();//실적합계
                    string sTOTP = dtSilTot.Rows[i]["TOTP"]?.ToString();//합계 달성율

                    double.TryParse(sI_WEIGHT, out double dI_WEIGHT);
                    double.TryParse(sI_LIGHT, out double dI_LIGHT);
                    double.TryParse(sI_CHAPI, out double dI_CHAPI);
                    double.TryParse(sI_YK, out double dI_YK);
                    double.TryParse(sW1, out double dW1);
                    double.TryParse(sW2, out double dW2);
                    double.TryParse(sW3, out double dW3);
                    double.TryParse(sW4, out double dW4);
                    double.TryParse(sM1, out double dM1);
                    double.TryParse(sM2, out double dM2);
                    double.TryParse(sM3, out double dM3);
                    double.TryParse(sM4, out double dM4);
                    double.TryParse(sWP1, out double dWP1);
                    double.TryParse(sWP2, out double dWP2);
                    double.TryParse(sWP3, out double dWP3);
                    double.TryParse(sWP4, out double dWP4);
                    double.TryParse(sTOTM, out double dTOTM);
                    double.TryParse(sTOTW, out double dTOTW);
                    double.TryParse(sTOTP, out double dTOTP);

                    sI_WEIGHT = dI_WEIGHT.ToString("n0");
                    sI_LIGHT = dI_LIGHT.ToString("n0");
                    sI_CHAPI = dI_CHAPI.ToString("n0");
                    sI_YK = dI_YK.ToString("n0");
                    sW1 = dW1.ToString("n0");
                    sW2 = dW2.ToString("n0");
                    sW3 = dW3.ToString("n0");
                    sW4 = dW4.ToString("n0");
                    sM1 = dM1.ToString("n1");
                    sM2 = dM2.ToString("n1");
                    sM3 = dM3.ToString("n1");
                    sM4 = dM4.ToString("n1");
                    sWP1 = dWP1.ToString("P1");
                    sWP2 = dWP2.ToString("P1");
                    sWP3 = dWP3.ToString("P1");
                    sWP4 = dWP4.ToString("P1");
                    sTOTM = dTOTM.ToString("n0");
                    sTOTW = dTOTW.ToString("n0");
                    sTOTP = dTOTP.ToString("P1");

                    //중량
                    row1["MOK" + (i + 1)] = sI_WEIGHT;//목표
                    row1["SIL" + (i + 1)] = sW1;//실적
                    row1["MAGIN" + (i + 1)] = sM1;//마진
                    row1["YUL" + (i + 1)] = sWP1;//달성율

                    //경량
                    row2["MOK" + (i + 1)] = sI_LIGHT;//목표
                    row2["SIL" + (i + 1)] = sW2;//실적
                    row2["MAGIN" + (i + 1)] = sM2;//마진
                    row2["YUL" + (i + 1)] = sWP2;//달성율

                    //폐압
                    row3["MOK" + (i + 1)] = sI_CHAPI;//목표
                    row3["SIL" + (i + 1)] = sW3;//실적
                    row3["MAGIN" + (i + 1)] = sM3;//마진
                    row3["YUL" + (i + 1)] = sWP3;//달성율

                    //합계
                    row4["MOK" + (i + 1)] = sTOTM;//목표
                    row4["SIL" + (i + 1)] = sTOTW;//실적
                    row4["MAGIN" + (i + 1)] = null;//마진
                    row4["YUL" + (i + 1)] = sTOTP;//달성율

                    //직납
                    row5["MOK" + (i + 1)] = sI_YK;//목표
                    row5["SIL" + (i + 1)] = sW4;//실적
                    row5["MAGIN" + (i + 1)] = sM4;//마진
                    row5["YUL" + (i + 1)] = sWP4;//달성율

                    GridBand gridBand = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
                    GridBand gridBandMok = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();//목표
                    GridBand gridBandSil = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();//실적
                    GridBand gridBandMa = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();//마진
                    GridBand gridBandYul = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();//달성율

                    gridView5.Bands.AddRange(new GridBand[] { gridBand });

                    gridBand.AppearanceHeader.Options.UseTextOptions = true;
                    gridBand.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBand.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBand.Caption = sNAME;
                    gridBand.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] {
                    gridBandMok,
                    gridBandSil,
                    gridBandMa,
                    gridBandYul});
                    gridBand.Name = "gridBand"+(i + 1);
                    gridBand.OptionsBand.AllowMove = false;
                    gridBand.VisibleIndex = 1;
                    gridBand.Width = 300;

                    BandedGridColumn bandedGridColumnMok = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();//목표

                    bandedGridColumnMok.AppearanceCell.Options.UseTextOptions = true;
                    bandedGridColumnMok.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                    bandedGridColumnMok.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    bandedGridColumnMok.FieldName = "MOK" + (i + 1);
                    bandedGridColumnMok.OptionsColumn.AllowEdit = false;
                    bandedGridColumnMok.Visible = true;

                    gridBandMok.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandMok.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandMok.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandMok.Caption = "목표";
                    gridBandMok.Columns.Add(bandedGridColumnMok);
                    gridBandMok.OptionsBand.AllowMove = false;
                    gridBandMok.VisibleIndex = 2;
                    gridBandMok.Width = 75;

                    BandedGridColumn bandedGridColumnSil = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();//실적

                    bandedGridColumnSil.AppearanceCell.Options.UseTextOptions = true;
                    bandedGridColumnSil.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                    bandedGridColumnSil.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    bandedGridColumnSil.FieldName = "SIL" + (i + 1);
                    bandedGridColumnSil.OptionsColumn.AllowEdit = false;
                    bandedGridColumnSil.Visible = true;

                    gridBandSil.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandSil.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandSil.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandSil.Caption = "실적";
                    gridBandSil.Columns.Add(bandedGridColumnSil);
                    gridBandSil.OptionsBand.AllowMove = false;
                    gridBandSil.VisibleIndex = 2;
                    gridBandSil.Width = 75;

                    BandedGridColumn bandedGridColumnMA = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();//마진

                    bandedGridColumnMA.AppearanceCell.Options.UseTextOptions = true;
                    bandedGridColumnMA.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    bandedGridColumnMA.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    bandedGridColumnMA.FieldName = "MAGIN" + (i + 1);
                    bandedGridColumnMA.OptionsColumn.AllowEdit = false;
                    bandedGridColumnMA.Visible = true;

                    gridBandMa.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandMa.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandMa.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandMa.Caption = "마진";
                    gridBandMa.Columns.Add(bandedGridColumnMA);
                    gridBandMa.OptionsBand.AllowMove = false;
                    gridBandMa.VisibleIndex = 2;
                    gridBandMa.Width = 75;

                    BandedGridColumn bandedGridColumnYul = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();//달성율

                    bandedGridColumnYul.AppearanceCell.Options.UseTextOptions = true;
                    bandedGridColumnYul.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    bandedGridColumnYul.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    bandedGridColumnYul.FieldName = "YUL" + (i + 1);
                    bandedGridColumnYul.OptionsColumn.AllowEdit = false;

                    gridBandYul.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandYul.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandYul.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandYul.Caption = "달성율";
                    gridBandYul.Columns.Add(bandedGridColumnYul);
                    gridBandYul.OptionsBand.AllowMove = false;
                    gridBandYul.VisibleIndex = 2;
                    gridBandYul.Width = 75;

                }

                dtTot.Rows.Add(row1);
                dtTot.Rows.Add(row2);
                dtTot.Rows.Add(row3);
                dtTot.Rows.Add(row4);
                dtTot.Rows.Add(row5);
                gridControl5.DataSource = dtTot;
            }

            #region 2022-03-15 이전
            //DataTable dtTot = new DataTable();

            //dtTot.Columns.Add("GUBUN");//구분
            ////손상영
            //dtTot.Columns.Add("SMOK");//목표
            //dtTot.Columns.Add("SSIL");//실적
            //dtTot.Columns.Add("SMA");//마진
            //dtTot.Columns.Add("SYUL");//달성율
            ////이우택
            //dtTot.Columns.Add("LMOK");
            //dtTot.Columns.Add("LSIL");
            //dtTot.Columns.Add("LMA");
            //dtTot.Columns.Add("LYUL"); 
            ////박봉섭
            //dtTot.Columns.Add("PMOK");
            //dtTot.Columns.Add("PSIL");
            //dtTot.Columns.Add("PMA");
            //dtTot.Columns.Add("PYUL");
            ////오상훈
            //dtTot.Columns.Add("OMOK");
            //dtTot.Columns.Add("OSIL");
            //dtTot.Columns.Add("OMA");
            //dtTot.Columns.Add("OYUL");
            ////김명철
            //dtTot.Columns.Add("KMOK");
            //dtTot.Columns.Add("KSIL");
            //dtTot.Columns.Add("KMA");
            //dtTot.Columns.Add("KYUL");
            ////합계
            //dtTot.Columns.Add("HMOK");
            //dtTot.Columns.Add("HSIL");
            //dtTot.Columns.Add("HMA");
            //dtTot.Columns.Add("HYUL");

            //DataRow row1 = dtTot.NewRow();
            //row1["GUBUN"] = "스크랩 중량";

            //DataRow row2 = dtTot.NewRow();
            //row2["GUBUN"] = "스크랩 경량";

            //DataRow row3 = dtTot.NewRow();
            //row3["GUBUN"] = "폐 압";

            //DataRow row4 = dtTot.NewRow();
            //row4["GUBUN"] = "합 계";

            //DataRow row5 = dtTot.NewRow();
            //row5["GUBUN"] = "직 납";

            //if(dtSilTot != null)
            //{
            //    for(int i=0;i< dtSilTot.Rows.Count; i++)
            //    {
            //        string sNAME     = dtSilTot.Rows[i]["NAME"]?.ToString();
            //     string sI_WEIGHT = dtSilTot.Rows[i]["I_WEIGHT"]?.ToString();// 중량목표
            //  string sI_LIGHT  = dtSilTot.Rows[i]["I_LIGHT"]?.ToString();// 경량목표
            //  string sI_CHAPI  = dtSilTot.Rows[i]["I_CHAPI"]?.ToString();// 폐압목표
            //  string sI_YK     = dtSilTot.Rows[i]["I_YK"]?.ToString();// 직납목표
            //  string sW1       = dtSilTot.Rows[i]["W1"]?.ToString();//중량실적
            //  string sW2       = dtSilTot.Rows[i]["W2"]?.ToString();//경량실적
            //  string sW3       = dtSilTot.Rows[i]["W3"]?.ToString();//폐압실적
            //  string sW4       = dtSilTot.Rows[i]["W4"]?.ToString();//직납실적
            //  string sM1       = dtSilTot.Rows[i]["M1"]?.ToString();//중량마진
            //  string sM2       = dtSilTot.Rows[i]["M2"]?.ToString();//경량마진
            //  string sM3       = dtSilTot.Rows[i]["M3"]?.ToString();//폐압마진
            //  string sM4       = dtSilTot.Rows[i]["M4"]?.ToString();//직납마진
            //  string sWP1      = dtSilTot.Rows[i]["WP1"]?.ToString();//중량달성율
            //  string sWP2      = dtSilTot.Rows[i]["WP2"]?.ToString();//경량달성율
            //  string sWP3      = dtSilTot.Rows[i]["WP3"]?.ToString();//폐압달성율
            //        string sWP4      = dtSilTot.Rows[i]["WP4"]?.ToString();//직납달성율
            //        string sTOTM     = dtSilTot.Rows[i]["TOTM"]?.ToString();//목표합계
            //        string sTOTW     = dtSilTot.Rows[i]["TOTW"]?.ToString();//실적합계
            //        string sTOTP     = dtSilTot.Rows[i]["TOTP"]?.ToString();//합계 달성율

            //        double.TryParse(sI_WEIGHT, out double dI_WEIGHT);
            //        double.TryParse(sI_LIGHT , out double dI_LIGHT );
            //        double.TryParse(sI_CHAPI , out double dI_CHAPI );
            //        double.TryParse(sI_YK    , out double dI_YK    );
            //        double.TryParse(sW1      , out double dW1      );
            //        double.TryParse(sW2      , out double dW2      );
            //        double.TryParse(sW3      , out double dW3      );
            //        double.TryParse(sW4      , out double dW4      );
            //        double.TryParse(sM1      , out double dM1      );
            //        double.TryParse(sM2      , out double dM2      );
            //        double.TryParse(sM3      , out double dM3      );
            //        double.TryParse(sM4      , out double dM4      );
            //        double.TryParse(sWP1     , out double dWP1     );
            //        double.TryParse(sWP2     , out double dWP2     );
            //        double.TryParse(sWP3     , out double dWP3     );
            //        double.TryParse(sWP4     , out double dWP4     );
            //        double.TryParse(sTOTM    , out double dTOTM    );
            //        double.TryParse(sTOTW    , out double dTOTW    );
            //        double.TryParse(sTOTP, out double dTOTP);


            //        sI_WEIGHT = dI_WEIGHT.ToString("n0");
            //        sI_LIGHT  = dI_LIGHT.ToString("n0");
            //        sI_CHAPI  = dI_CHAPI.ToString("n0");
            //        sI_YK     = dI_YK.ToString("n0");
            //        sW1       = dW1.ToString("n0");
            //        sW2       = dW2.ToString("n0");
            //        sW3       = dW3.ToString("n0");
            //        sW4       = dW4.ToString("n0");
            //        sM1       = dM1.ToString("n1");
            //        sM2       = dM2.ToString("n1");
            //        sM3       = dM3.ToString("n1");
            //        sM4       = dM4.ToString("n1");
            //        sWP1      = dWP1.ToString("P1");
            //        sWP2      = dWP2.ToString("P1");
            //        sWP3      = dWP3.ToString("P1");
            //        sWP4      = dWP4.ToString("P1");
            //        sTOTM     = dTOTM.ToString("n0");
            //        sTOTW     = dTOTW.ToString("n0");
            //        sTOTP = dTOTP.ToString("P1");

            //        if (sNAME.Equals("손상영"))
            //        {
            //            //중량
            //            row1["SMOK"] = sI_WEIGHT;
            //            row1["SSIL"] = sW1;
            //            row1["SMA" ] = sM1;
            //            row1["SYUL"] = sWP1;

            //            //경량
            //            row2["SMOK"] = sI_LIGHT;
            //            row2["SSIL"] = sW2;
            //            row2["SMA"] = sM2;
            //            row2["SYUL"] = sWP2;

            //            //폐압
            //            row3["SMOK"] = sI_CHAPI;
            //            row3["SSIL"] = sW3;
            //            row3["SMA"] = sM3;
            //            row3["SYUL"] = sWP3;

            //            //합계
            //            row4["SMOK"] = sTOTM;
            //            row4["SSIL"] = sTOTW;
            //            row4["SMA"] = null;
            //            row4["SYUL"] = sTOTP;

            //            //직납
            //            row5["SMOK"] = sI_YK;
            //            row5["SSIL"] = sW4;
            //            row5["SMA"] = sM4;
            //            row5["SYUL"] = sWP4;
            //        }
            //        else if (sNAME.Equals("이우택"))
            //        {
            //            //중량
            //            row1["LMOK"] = sI_WEIGHT;
            //            row1["LSIL"] = sW1;
            //            row1["LMA"] = sM1;
            //            row1["LYUL"] = sWP1;

            //            //경량
            //            row2["LMOK"] = sI_LIGHT;
            //            row2["LSIL"] = sW2;
            //            row2["LMA"] = sM2;
            //            row2["LYUL"] = sWP2;

            //            //폐압
            //            row3["LMOK"] = sI_CHAPI;
            //            row3["LSIL"] = sW3;
            //            row3["LMA"] = sM3;
            //            row3["LYUL"] = sWP3;

            //            //합계
            //            row4["LMOK"] = sTOTM;
            //            row4["LSIL"] = sTOTW;
            //            row4["LMA"] = null;
            //            row4["LYUL"] = sTOTP;

            //            //직납
            //            row5["LMOK"] = sI_YK;
            //            row5["LSIL"] = sW4;
            //            row5["LMA"] = sM4;
            //            row5["LYUL"] = sWP4;
            //        }
            //        else if (sNAME.Equals("박봉섭"))
            //        {
            //            //중량
            //            row1["PMOK"] = sI_WEIGHT;
            //            row1["PSIL"] = sW1;
            //            row1["PMA"] = sM1;
            //            row1["PYUL"] = sWP1;

            //            //경량
            //            row2["PMOK"] = sI_LIGHT;
            //            row2["PSIL"] = sW2;
            //            row2["PMA"] = sM2;
            //            row2["PYUL"] = sWP2;

            //            //폐압
            //            row3["PMOK"] = sI_CHAPI;
            //            row3["PSIL"] = sW3;
            //            row3["PMA"] = sM3;
            //            row3["PYUL"] = sWP3;

            //            //합계
            //            row4["PMOK"] = sTOTM;
            //            row4["PSIL"] = sTOTW;
            //            row4["PMA"] = null;
            //            row4["PYUL"] = sTOTP;

            //            //직납
            //            row5["PMOK"] = sI_YK;
            //            row5["PSIL"] = sW4;
            //            row5["PMA"] = sM4;
            //            row5["PYUL"] = sWP4;
            //        }
            //        else if (sNAME.Equals("오상훈"))
            //        {
            //            //중량
            //            row1["OMOK"] = sI_WEIGHT;
            //            row1["OSIL"] = sW1;
            //            row1["OMA"] = sM1;
            //            row1["OYUL"] = sWP1;

            //            //경량
            //            row2["OMOK"] = sI_LIGHT;
            //            row2["OSIL"] = sW2;
            //            row2["OMA"] = sM2;
            //            row2["OYUL"] = sWP2;

            //            //폐압
            //            row3["OMOK"] = sI_CHAPI;
            //            row3["OSIL"] = sW3;
            //            row3["OMA"] = sM3;
            //            row3["OYUL"] = sWP3;

            //            //합계
            //            row4["OMOK"] = sTOTM;
            //            row4["OSIL"] = sTOTW;
            //            row4["OMA"] = null;
            //            row4["OYUL"] = sTOTP;

            //            //직납
            //            row5["OMOK"] = sI_YK;
            //            row5["OSIL"] = sW4;
            //            row5["OMA"] = sM4;
            //            row5["OYUL"] = sWP4;
            //        }
            //        else if (sNAME.Equals("김명철"))
            //        {
            //            //중량
            //            row1["KMOK"] = sI_WEIGHT;
            //            row1["KSIL"] = sW1;
            //            row1["KMA"] = sM1;
            //            row1["KYUL"] = sWP1;

            //            //경량
            //            row2["KMOK"] = sI_LIGHT;
            //            row2["KSIL"] = sW2;
            //            row2["KMA"] = sM2;
            //            row2["KYUL"] = sWP2;

            //            //폐압
            //            row3["KMOK"] = sI_CHAPI;
            //            row3["KSIL"] = sW3;
            //            row3["KMA"] = sM3;
            //            row3["KYUL"] = sWP3;

            //            //합계
            //            row4["KMOK"] = sTOTM;
            //            row4["KSIL"] = sTOTW;
            //            row4["KMA"] = null;
            //            row4["KYUL"] = sTOTP;

            //            //직납
            //            row5["KMOK"] = sI_YK;
            //            row5["KSIL"] = sW4;
            //            row5["KMA"] = sM4;
            //            row5["KYUL"] = sWP4;
            //        }
            //        else if (sNAME.Equals("합계"))
            //        {
            //            //중량
            //            row1["HMOK"] = sI_WEIGHT;
            //            row1["HSIL"] = sW1;
            //            row1["HMA"] = sM1;
            //            row1["HYUL"] = sWP1;

            //            //경량
            //            row2["HMOK"] = sI_LIGHT;
            //            row2["HSIL"] = sW2;
            //            row2["HMA"] = sM2;
            //            row2["HYUL"] = sWP2;

            //            //폐압
            //            row3["HMOK"] = sI_CHAPI;
            //            row3["HSIL"] = sW3;
            //            row3["HMA"] = sM3;
            //            row3["HYUL"] = sWP3;

            //            //합계
            //            row4["HMOK"] = sTOTM;
            //            row4["HSIL"] = sTOTW;
            //            row4["HMA"] = null;
            //            row4["HYUL"] = sTOTP;

            //            //직납
            //            row5["HMOK"] = sI_YK;
            //            row5["HSIL"] = sW4;
            //            row5["HMA"] = sM4;
            //            row5["HYUL"] = sWP4;
            //        }
            //    }
            //}
            //dtTot.Rows.Add(row1);
            //dtTot.Rows.Add(row2);
            //dtTot.Rows.Add(row3);
            //dtTot.Rows.Add(row4);
            //dtTot.Rows.Add(row5);

            //gridControl5.DataSource = dtTot;
            #endregion

            #endregion
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            FileInfo_1 fileInfo = new FileInfo_1("7");

            Cursor = Cursors.WaitCursor;
            string[] sPath = fileInfo.CheckFileInfo();
            Cursor = Cursors.Default;

            if (sPath != null)
            {
                SetWeeklyReportForm(sPath[0], sPath[1]);
            }
        }

        Excel.Application ExcelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;
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

                string sFromDate = DateFrom.EditValue?.ToString().Substring(0, 10);
                string sToDate = DateTo.EditValue?.ToString().Substring(0, 10);

                ExcelApp = new Excel.Application();
                wb = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                ws = wb.Worksheets.get_Item(1);

                //일자부분 세팅
                ws.Range["F5"].Value = sFromDate;
                ws.Range["G5"].Value = sToDate;

                //기준세팅
                string sIncen = TxtIncen.EditValue?.ToString();//인센티브
                string sAumh = TxtAumH.EditValue?.ToString();//어음할인율
                string sUmija = TxtUmIja.EditValue?.ToString();//어음할인이자
                string sPagimul = TxtPagimul.EditValue?.ToString();//페기물처리비

                ws.Range["C5"].Value = sIncen;
                ws.Range["D5"].Value = sAumh;
                ws.Range["E5"].Value = sUmija;
                ws.Range["M4"].Value = sPagimul;


                //매출평균단가 
                string sAvDan1 = TeditMeachul.EditValue?.ToString();//중량
                string sAvDan2 = TeditMeachul1.EditValue?.ToString();//경량
                string sAvDan3 = TeditMeachul2.EditValue?.ToString();//슈레더

                ws.Range["E7"].Value = sAvDan1;
                ws.Range["E17"].Value = sAvDan2;
                ws.Range["K4"].Value = sAvDan3;

                #region 스크랩(중량)
                DataTable dt1 = (DataTable)gridControl1.DataSource;

                if(dt1 != null)
                {
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        string sNAME    = dt1.Rows[i]["NAME"]?.ToString();
                        string sDANJUNG = dt1.Rows[i]["DANJUNG"]?.ToString();
                        string sHALIN   = dt1.Rows[i]["HALIN"]?.ToString();
                        string sDOAVDAN = dt1.Rows[i]["DOAVDAN"]?.ToString();
                        string sMAGIN   = dt1.Rows[i]["MAGIN"]?.ToString();
                        string sMCHEIC = dt1.Rows[i]["MCHEIC"]?.ToString();

                        if (sNAME.Equals("손상영"))
                        {
                            ws.Range["C10"].Value = sDANJUNG;
                            ws.Range["D10"].Value = sHALIN;
                            ws.Range["E10"].Value = sDOAVDAN;
                            ws.Range["F10"].Value = sMAGIN;
                            ws.Range["G10"].Value = sMCHEIC;
                        }
                        else if (sNAME.Equals("이우택"))
                        {
                            ws.Range["C11"].Value = sDANJUNG;
                            ws.Range["D11"].Value = sHALIN;
                            ws.Range["E11"].Value = sDOAVDAN;
                            ws.Range["F11"].Value = sMAGIN;
                            ws.Range["G11"].Value = sMCHEIC;
                        }
                        else if (sNAME.Equals("박봉섭"))
                        {
                            ws.Range["C12"].Value = sDANJUNG;
                            ws.Range["D12"].Value = sHALIN;
                            ws.Range["E12"].Value = sDOAVDAN;
                            ws.Range["F12"].Value = sMAGIN;
                            ws.Range["G12"].Value = sMCHEIC;
                        }
                        else if (sNAME.Equals("오상훈"))
                        {
                            ws.Range["C13"].Value = sDANJUNG;
                            ws.Range["D13"].Value = sHALIN;
                            ws.Range["E13"].Value = sDOAVDAN;
                            ws.Range["F13"].Value = sMAGIN;
                            ws.Range["G13"].Value = sMCHEIC;
                        }
                        else if (sNAME.Equals("김명철"))
                        {
                            ws.Range["C14"].Value = sDANJUNG;
                            ws.Range["D14"].Value = sHALIN;
                            ws.Range["E14"].Value = sDOAVDAN;
                            ws.Range["F14"].Value = sMAGIN;
                            ws.Range["G14"].Value = sMCHEIC;
                        }
                    }
                }
                #endregion

                #region 스크랩 (경량)
                DataTable dt3 = (DataTable)gridControl3.DataSource;

                if(dt3 != null)
                {
                    int iStartRow1 = 20;//손상영
                    int iStartRow2 = 28;//이우택
                    int iStartRow3 = 36;//박봉섭
                    int iStartRow4 = 44;//오상훈
                    int iStartRow5 = 52;//김명철

                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        string sNAME    = dt3.Rows[i]["NAME"]?.ToString();
                        string sGUBUN1  = dt3.Rows[i]["GUBUN1"]?.ToString();
                        string sDANJUNG = dt3.Rows[i]["DANJUNG"]?.ToString();
                        string sHALIN   = dt3.Rows[i]["HALIN"]?.ToString();
                        string sDOAVDAN = dt3.Rows[i]["DOAVDAN"]?.ToString();
                        string sMAGIN   = dt3.Rows[i]["MAGIN"]?.ToString();
                        string sMCHEIC = dt3.Rows[i]["MCHEIC"]?.ToString();

                        int iApplyRowIdx = 0;

                        if (sNAME.Equals("손상영"))
                        {
                            iApplyRowIdx = iStartRow1++;

                            if (sGUBUN1.Equals("합 계"))
                            {
                                ws.Range["E27"].Value = sDOAVDAN;

                                continue;
                            }
                        }
                        else if (sNAME.Equals("이우택"))
                        {
                            iApplyRowIdx = iStartRow2++;

                            if (sGUBUN1.Equals("합 계"))
                            {
                                ws.Range["E35"].Value = sDOAVDAN;

                                continue;
                            }
                        }
                        else if (sNAME.Equals("박봉섭"))
                        {
                            iApplyRowIdx = iStartRow3++;

                            if (sGUBUN1.Equals("합 계"))
                            {
                                ws.Range["E43"].Value = sDOAVDAN;

                                continue;
                            }
                        }
                        else if (sNAME.Equals("오상훈"))
                        {
                            iApplyRowIdx = iStartRow4++;

                            if (sGUBUN1.Equals("합 계"))
                            {
                                ws.Range["E51"].Value = sDOAVDAN;

                                continue;
                            }
                        }
                        else if (sNAME.Equals("김명철"))
                        {
                            iApplyRowIdx = iStartRow5++;

                            if (sGUBUN1.Equals("합 계"))
                            {
                                ws.Range["E59"].Value = sDOAVDAN;

                                continue;
                            }
                        }

                        if (iApplyRowIdx == 0)
                            continue;

                        ws.Range["B" + iApplyRowIdx].Value = sGUBUN1;
                        ws.Range["C" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["D" + iApplyRowIdx].Value = sHALIN;
                        ws.Range["E" + iApplyRowIdx].Value = sDOAVDAN;
                        ws.Range["F" + iApplyRowIdx].Value = sMAGIN;
                        ws.Range["G" + iApplyRowIdx].Value = sMCHEIC;
                    }
                }
                #endregion

                #region 슈레더
                DataTable dt4 = (DataTable)gridControl4.DataSource;

                if(dt4 != null)
                {
                    int iStartRow1 = 8;//손상영
                    int iStartRow2 = 16;//이우택
                    int iStartRow3 = 24;//박봉섭
                    int iStartRow4 = 32;//오상훈
                    int iStartRow5 = 40;//김명철

                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {
                        string sNAME    = dt4.Rows[i]["NAME"]?.ToString();
                        string sGUBUN1  = dt4.Rows[i]["GUBUN1"]?.ToString();
                        string sDANJUNG = dt4.Rows[i]["DANJUNG"]?.ToString();
                        string sMIPDAN  = dt4.Rows[i]["MIPDAN"]?.ToString();
                        string sDOAMT   = dt4.Rows[i]["DOAMT"]?.ToString();
                        string sMACHEIC = dt4.Rows[i]["MACHEIC"]?.ToString();

                        int iApplyRowIdx = 0;

                        if (sNAME.Equals("손상영"))
                        {
                            iApplyRowIdx = iStartRow1++;
                        }
                        else if (sNAME.Equals("이우택"))
                        {
                            iApplyRowIdx = iStartRow2++;
                        }
                        else if (sNAME.Equals("박봉섭"))
                        {
                            iApplyRowIdx = iStartRow3++;
                        }
                        else if (sNAME.Equals("오상훈"))
                        {
                            iApplyRowIdx = iStartRow4++;
                        }
                        else if (sNAME.Equals("김명철"))
                        {
                            iApplyRowIdx = iStartRow5++;
                        }

                        if (iApplyRowIdx == 0 || sGUBUN1.Equals("합 계"))
                            continue;

                        ws.Range["I" + iApplyRowIdx].Value = sGUBUN1;
                        ws.Range["J" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["K" + iApplyRowIdx].Value = sMIPDAN;
                        ws.Range["L" + iApplyRowIdx].Value = sDOAMT;
                        ws.Range["M" + iApplyRowIdx].Value = sMACHEIC;
                    }
                }
                #endregion

                #region 직송
                DataTable dt2 = (DataTable)gridControl2.DataSource;

                if (dt2 != null)
                {
                    for(int i = 0; i < dt2.Rows.Count; i++)
                    {
                        string sNAME      = dt2.Rows[i]["NAME"]?.ToString();
                        string sDANJUNG   = dt2.Rows[i]["DANJUNG"]?.ToString();
                        string sJ_KONGKEP = dt2.Rows[i]["J_KONGKEP"]?.ToString();
                        string sDOAMT     = dt2.Rows[i]["DOAMT"]?.ToString();
                        string sMAGIN     = dt2.Rows[i]["MAGIN"]?.ToString();
                        string sMACHEIC = dt2.Rows[i]["MACHEIC"]?.ToString();

                        if (sNAME.Equals("손상영"))
                        {
                            ws.Range["I52"].Value = sDANJUNG;
                            ws.Range["J52"].Value = sJ_KONGKEP;
                            ws.Range["K52"].Value = sDOAMT;
                            ws.Range["L52"].Value = sMAGIN;
                            ws.Range["M52"].Value = sMACHEIC;
                        }
                        else if (sNAME.Equals("이우택"))
                        {
                            ws.Range["I53"].Value = sDANJUNG;
                            ws.Range["J53"].Value = sJ_KONGKEP;
                            ws.Range["K53"].Value = sDOAMT;
                            ws.Range["L53"].Value = sMAGIN;
                            ws.Range["M53"].Value = sMACHEIC;
                        }
                        else if (sNAME.Equals("박봉섭"))
                        {
                            ws.Range["I54"].Value = sDANJUNG;
                            ws.Range["J54"].Value = sJ_KONGKEP;
                            ws.Range["K54"].Value = sDOAMT;
                            ws.Range["L54"].Value = sMAGIN;
                            ws.Range["M54"].Value = sMACHEIC;
                        }
                        else if (sNAME.Equals("오상훈"))
                        {
                            ws.Range["I55"].Value = sDANJUNG;
                            ws.Range["J55"].Value = sJ_KONGKEP;
                            ws.Range["K55"].Value = sDOAMT;
                            ws.Range["L55"].Value = sMAGIN;
                            ws.Range["M55"].Value = sMACHEIC;
                        }
                        else if (sNAME.Equals("김명철"))
                        {
                            ws.Range["I56"].Value = sDANJUNG;
                            ws.Range["J56"].Value = sJ_KONGKEP;
                            ws.Range["K56"].Value = sDOAMT;
                            ws.Range["L56"].Value = sMAGIN;
                            ws.Range["M56"].Value = sMACHEIC;
                        }
                    }
                }
                #endregion

                #region 집계
                DataTable dt5 = (DataTable)gridControl5.DataSource;

                if (dt5 != null)
                {
                    int iStartRow = 64;

                    for (int i = 0; i < dt5.Rows.Count; i++)
                    {
                        string sGUBUN = dt5.Rows[i]["GUBUN"]?.ToString();
                        //손상영
                        string sSMOK  = dt5.Rows[i]["SMOK"]?.ToString();
                        string sSSIL  = dt5.Rows[i]["SSIL"]?.ToString();
                        string sSMA   = dt5.Rows[i]["SMA"]?.ToString();
                        string sSYUL  = dt5.Rows[i]["SYUL"]?.ToString();
                        //이우택
                        string sLMOK  = dt5.Rows[i]["LMOK"]?.ToString();
                        string sLSIL  = dt5.Rows[i]["LSIL"]?.ToString();
                        string sLMA   = dt5.Rows[i]["LMA"]?.ToString();
                        string sLYUL  = dt5.Rows[i]["LYUL"]?.ToString();
                        //박봉섭
                        string sPMOK  = dt5.Rows[i]["PMOK"]?.ToString();
                        string sPSIL  = dt5.Rows[i]["PSIL"]?.ToString();
                        string sPMA   = dt5.Rows[i]["PMA"]?.ToString();
                        string sPYUL  = dt5.Rows[i]["PYUL"]?.ToString();
                        //오상훈
                        string sOMOK  = dt5.Rows[i]["OMOK"]?.ToString();
                        string sOSIL  = dt5.Rows[i]["OSIL"]?.ToString();
                        string sOMA   = dt5.Rows[i]["OMA"]?.ToString();
                        string sOYUL  = dt5.Rows[i]["OYUL"]?.ToString();
                        //김명철
                        string sKMOK  = dt5.Rows[i]["KMOK"]?.ToString();
                        string sKSIL  = dt5.Rows[i]["KSIL"]?.ToString();
                        string sKMA   = dt5.Rows[i]["KMA"]?.ToString();
                        string sKYUL  = dt5.Rows[i]["KYUL"]?.ToString();
                        //합계
                        string sHMOK  = dt5.Rows[i]["HMOK"]?.ToString();
                        string sHSIL  = dt5.Rows[i]["HSIL"]?.ToString();
                        string sHMA   = dt5.Rows[i]["HMA"]?.ToString();
                        string sHYUL = dt5.Rows[i]["HYUL"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["B" + iApplyRowIdx].Value = sSMOK;
                        ws.Range["F" + iApplyRowIdx].Value = sLMOK;
                        ws.Range["J" + iApplyRowIdx].Value = sPMOK;
                        ws.Range["N" + iApplyRowIdx].Value = sOMOK;
                        ws.Range["R" + iApplyRowIdx].Value = sKMOK;
                    }
                }
                #endregion

                if (File.Exists(SavePath))
                    File.Delete(SavePath);

                Cursor = Cursors.Default;
                //wb.SaveAs(SavePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //파일 닫기... 
                wb.SaveAs(SavePath);
                wb.Close(false, Type.Missing, Type.Missing);
                wb = null;
                ExcelApp.Quit();

                Process.Start(SavePath);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                XtraMessageBox.Show(ex.Message);

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

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //string sFromDate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sFromDate = DateTime.Today.ToString("yyyy-MM-dd");
            string sIncen = TxtIncen.EditValue?.ToString();//인센티브
            string sAumh = TxtAumH.EditValue?.ToString();//어음할인율
            string sUmija = TxtUmIja.EditValue?.ToString();//어음할인이자
            string sPagimul = TxtPagimul.EditValue?.ToString();//페기물처리비

            //회수율
            string sH1 = TxtH1.EditValue?.ToString();
            string sH2 = TxtH2.EditValue?.ToString();
            string sH3 = TxtH3.EditValue?.ToString();
            string sH4 = TxtH4.EditValue?.ToString();
            string sH5 = TxtH5.EditValue?.ToString();
            string sH6 = TxtH6.EditValue?.ToString();
            string sH7 = TxtH7.EditValue?.ToString();
            string sH8 = TxtH8.EditValue?.ToString();
            string sH9 = TxtH9.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                if (!string.IsNullOrEmpty(sIncen) && _EDIT_Incen == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '인센티브')");
                    strSql.AppendLine("    BEGIN                                                                             ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                          ");
                    strSql.AppendLine("             SET DATAVALUE = " + sIncen + "                                                        ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '인센티브'                         ");
                    strSql.AppendLine("      END                                                                             ");
                    strSql.AppendLine(" ELSE                                                                                 ");
                    strSql.AppendLine("    BEGIN                                                                             ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                              ");
                    strSql.AppendLine("                                 , DAEGUBUN                                           ");
                    strSql.AppendLine("                                 , GUBUN                                              ");
                    strSql.AppendLine("                                 , DATAVALUE)                                         ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                  ");
                    strSql.AppendLine("                                 , '기초자료'                                         ");
                    strSql.AppendLine("                                 , '인센티브'                                         ");
                    strSql.AppendLine("                                 , " + sIncen + ")                                                 ");
                    strSql.AppendLine("      END                                                                             ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sAumh) && _EDIT_AumH == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '어음할인') ");
                    strSql.AppendLine("    BEGIN                                                                              ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                           ");
                    strSql.AppendLine("             SET DATAVALUE = "+ sAumh + "                                                         ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '어음할인'                          ");
                    strSql.AppendLine("      END                                                                              ");
                    strSql.AppendLine(" ELSE                                                                                  ");
                    strSql.AppendLine("    BEGIN                                                                              ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                               ");
                    strSql.AppendLine("                                 , DAEGUBUN                                            ");
                    strSql.AppendLine("                                 , GUBUN                                               ");
                    strSql.AppendLine("                                 , DATAVALUE)                                          ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                   ");
                    strSql.AppendLine("                                 , '기초자료'                                          ");
                    strSql.AppendLine("                                 , '어음할인'                                          ");
                    strSql.AppendLine("                                 , "+ sAumh + ")                                                  ");
                    strSql.AppendLine("      END                                                                              ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sUmija) && _EDIT_UmIja == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '어음이자')");
                    strSql.AppendLine("    BEGIN                                                                             ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                          ");
                    strSql.AppendLine("             SET DATAVALUE = "+ sUmija + "                                                        ");
                    strSql.AppendLine("           WHERE J_DATE = '"+ sFromDate + "' AND GUBUN = '어음이자'                         ");
                    strSql.AppendLine("      END                                                                             ");
                    strSql.AppendLine(" ELSE                                                                                 ");
                    strSql.AppendLine("    BEGIN                                                                             ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                              ");
                    strSql.AppendLine("                                 , DAEGUBUN                                           ");
                    strSql.AppendLine("                                 , GUBUN                                              ");
                    strSql.AppendLine("                                 , DATAVALUE)                                         ");
                    strSql.AppendLine("                           VALUES('"+ sFromDate + "'                                                  ");
                    strSql.AppendLine("                                 , '기초자료'                                         ");
                    strSql.AppendLine("                                 , '어음이자'                                         ");
                    strSql.AppendLine("                                 , "+ sUmija + ")                                                 ");
                    strSql.AppendLine("      END                                                                             ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }                                                                                                            

                if (!string.IsNullOrEmpty(sPagimul) && _EDIT_Pagimul == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '폐기물처리비')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET DATAVALUE = "+ sPagimul + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '"+ sFromDate + "' AND GUBUN = '폐기물처리비'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , DATAVALUE)                                             ");
                    strSql.AppendLine("                           VALUES('"+ sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '기초자료'                                             ");
                    strSql.AppendLine("                                 , '폐기물처리비'                                         ");
                    strSql.AppendLine("                                 , "+ sPagimul + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH1) && _EDIT_Hl == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재특급A')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH1 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재특급A'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재특급A'                                         ");
                    strSql.AppendLine("                                 , " + sH1 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH2) && _EDIT_H2 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재특급B')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH2 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재특급B'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재특급B'                                         ");
                    strSql.AppendLine("                                 , " + sH2 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH3) && _EDIT_H3 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재A')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH3 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재A'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재A'                                         ");
                    strSql.AppendLine("                                 , " + sH3 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH4) && _EDIT_H4 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재B')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH4 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재B'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재B'                                         ");
                    strSql.AppendLine("                                 , " + sH4 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH5) && _EDIT_H5 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재C')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH5 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재C'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재C'                                         ");
                    strSql.AppendLine("                                 , " + sH5 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH6) && _EDIT_H6 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '슈레더C')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH6 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '슈레더C'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '슈레더C'                                         ");
                    strSql.AppendLine("                                 , " + sH6 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH7) && _EDIT_H7 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재AA')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH7 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재AA'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재AA'                                         ");
                    strSql.AppendLine("                                 , " + sH7 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH8) && _EDIT_H8 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '슈레더A')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH8 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '슈레더A'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '슈레더A'                                         ");
                    strSql.AppendLine("                                 , " + sH8 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrEmpty(sH9) && _EDIT_H9 == EDIT_YN.Editable)
                {
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재P')");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          UPDATE MEAIPSILJUK                                                              ");
                    strSql.AppendLine("             SET RECOVERY = " + sH9 + "                                                            ");
                    strSql.AppendLine("           WHERE J_DATE = '" + sFromDate + "' AND GUBUN = '모재P'                         ");
                    strSql.AppendLine("      END                                                                                 ");
                    strSql.AppendLine(" ELSE                                                                                     ");
                    strSql.AppendLine("    BEGIN                                                                                 ");
                    strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                                  ");
                    strSql.AppendLine("                                 , DAEGUBUN                                               ");
                    strSql.AppendLine("                                 , GUBUN                                                  ");
                    strSql.AppendLine("                                 , RECOVERY)                                             ");
                    strSql.AppendLine("                           VALUES('" + sFromDate + "'                                                      ");
                    strSql.AppendLine("                                 , '슈레더'                                             ");
                    strSql.AppendLine("                                 , '모재P'                                         ");
                    strSql.AppendLine("                                 , " + sH9 + ")                                                     ");
                    strSql.AppendLine("      END                                                                                 ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr.PerformClick();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            string sAvDan1 = TeditDanga1.EditValue?.ToString();//중량
            string sAvDan2 = TeditDanga2.EditValue?.ToString();//경량
            string sAvDan3 = TeditDanga3.EditValue?.ToString();//슈레더

            TeditMeachul.EditValue = sAvDan1;//중량
            TeditMeachul1.EditValue = sAvDan2;//경량
            TeditMeachul2.EditValue = sAvDan3;//슈레더

            GetRetrData("적용후조회");
        }

        private void PurTeamRecord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void TeditDanga3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnApply.Focus();
        }

        private void TxtPagimul_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSave.Focus();
        }

        private void TxtH9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSave.Focus();
        }

        private void gridView1_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sName = gridView1.GetRowCellValue(e.RowHandle, "NAME")?.ToString();

            if (!string.IsNullOrEmpty(sName))
            {
                if(sName.Equals("합 계"))
                {
                    e.Appearance.BackColor = Color.LightYellow;
                }
                else
                {
                    ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                }
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void gridView3_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            string sName = gridView3.GetRowCellValue(e.RowHandle, "NAME")?.ToString();
            string sGuBun = gridView3.GetRowCellValue(e.RowHandle, "GUBUN1")?.ToString();

            if ((!string.IsNullOrEmpty(sName) && sName.Equals("총 합 계"))|| (!string.IsNullOrEmpty(sGuBun) && sGuBun.Equals("합 계")))
            {
                e.Appearance.BackColor = Color.LightYellow;
            }
        }

        private void gridView4_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            string sGuBun = gridView4.GetRowCellValue(e.RowHandle, "GUBUN1")?.ToString();

            if (!string.IsNullOrEmpty(sGuBun) && sGuBun.Equals("합 계"))
            {
                e.Appearance.BackColor = Color.LightYellow;
            }
        }

        private void gridView2_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            string sNAME = gridView2.GetRowCellValue(e.RowHandle, "NAME")?.ToString();

            if (!string.IsNullOrEmpty(sNAME) && sNAME.Equals("합 계"))
            {
                e.Appearance.BackColor = Color.LightYellow;
            }
        }

        private void gridView5_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            string sNAME = gridView5.GetRowCellValue(e.RowHandle, "GUBUN")?.ToString();

            if (!string.IsNullOrEmpty(sNAME) && sNAME.Equals("합 계"))
            {
                e.Appearance.BackColor = Color.LightYellow;
            }
        }

        private void gridView5_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
                 
            if (e.Column.FieldName == "SYUL" || e.Column.FieldName == "LYUL"
                || e.Column.FieldName == "PYUL" || e.Column.FieldName == "OYUL"
                || e.Column.FieldName == "KYUL" || e.Column.FieldName == "HYUL")
            {
                ComnGridFunc.CellDrawHelper.DoDefaultDrawCell(gridView5, e);
                ComnGridFunc.CellDrawHelper.DrawCellBorder(e);
                e.Handled = true;
            }
        }

        //값 변경확인
        private enum EDIT_YN { Editable, NonEditable }

        //인센티브
        private EDIT_YN _EDIT_Incen = EDIT_YN.NonEditable;
        private void TxtIncen_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_Incen = EDIT_YN.Editable;
        }
        //어음할인
        private EDIT_YN _EDIT_AumH = EDIT_YN.NonEditable;
        private void TxtAumH_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_AumH = EDIT_YN.Editable;
        }
        //어음이자
        private EDIT_YN _EDIT_UmIja = EDIT_YN.NonEditable;
        private void TxtUmIja_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_UmIja = EDIT_YN.Editable;
        }
        //페기물처리비
        private EDIT_YN _EDIT_Pagimul = EDIT_YN.NonEditable;
        private void TxtPagimul_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_Pagimul = EDIT_YN.Editable;
        }
        //모재특급A
        private EDIT_YN _EDIT_Hl = EDIT_YN.NonEditable;
        private void TxtH1_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_Hl = EDIT_YN.Editable;
        }
        //모재특급B
        private EDIT_YN _EDIT_H2 = EDIT_YN.NonEditable;
        private void TxtH2_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H2 = EDIT_YN.Editable;
        }
        //모재A
        private EDIT_YN _EDIT_H3 = EDIT_YN.NonEditable;
        private void TxtH3_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H3 = EDIT_YN.Editable;
        }
        //모재B
        private EDIT_YN _EDIT_H4 = EDIT_YN.NonEditable;
        private void TxtH4_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H4 = EDIT_YN.Editable;
        }
        //모재C
        private EDIT_YN _EDIT_H5 = EDIT_YN.NonEditable;
        private void TxtH5_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H5 = EDIT_YN.Editable;
        }
        //슈레더C
        private EDIT_YN _EDIT_H6 = EDIT_YN.NonEditable;
        private void TxtH6_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H6 = EDIT_YN.Editable;
        }
        //모재AA
        private EDIT_YN _EDIT_H7 = EDIT_YN.NonEditable;
        private void TxtH7_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H7 = EDIT_YN.Editable;
        }
        //슈레더A
        private EDIT_YN _EDIT_H8 = EDIT_YN.NonEditable;
        private void TxtH8_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H8 = EDIT_YN.Editable;
        }
        //모재P
        private EDIT_YN _EDIT_H9 = EDIT_YN.NonEditable;
        private void TxtH9_EditValueChanged(object sender, EventArgs e)
        {
            _EDIT_H9 = EDIT_YN.Editable;
        }

        //EDIT_YN 초기화
        private void ResetEditYn()
        {
            //인센티브
            _EDIT_Incen = EDIT_YN.NonEditable;
            //어음할인
            _EDIT_AumH = EDIT_YN.NonEditable;
            //어음이자
            _EDIT_UmIja = EDIT_YN.NonEditable;
            //페기물처리비
            _EDIT_Pagimul = EDIT_YN.NonEditable;
            //모재특급A
            _EDIT_Hl = EDIT_YN.NonEditable;
            //모재특급B
            _EDIT_H2 = EDIT_YN.NonEditable;
            //모재A
            _EDIT_H3 = EDIT_YN.NonEditable;
            //모재B
            _EDIT_H4 = EDIT_YN.NonEditable;
            //모재C
            _EDIT_H5 = EDIT_YN.NonEditable;
            //슈레더C
            _EDIT_H6 = EDIT_YN.NonEditable;
            //모재AA
            _EDIT_H7 = EDIT_YN.NonEditable;
            //슈레더A
            _EDIT_H8 = EDIT_YN.NonEditable;
            //모재P
            _EDIT_H9 = EDIT_YN.NonEditable;
        }

        private void PurTeamRecord_TextChanged(object sender, EventArgs e)
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
    }
}