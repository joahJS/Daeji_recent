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
using System.Data.SqlClient;
using DevExpress.XtraSplashScreen;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace AccAdm
{

    public partial class DailyUnitCost : DevExpress.XtraEditors.XtraForm
    {
        public DailyUnitCost()
        {
            InitializeComponent();
        }

        private void DailyJikNap_Load(object sender, EventArgs e)
        {
            DateEditDT.EditValue = DateTime.Today;
            SetLoadFormLayout();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr1, GridViewRetr2, GridViewRetr3, BGridViewRetr4, BGridRetrView5, GridRetrView6, GridViewRetr7, GridViewRetr8, GridViewRetr9 };
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            
            string sDate = DateEditDT.EditValue?.ToString().Substring(0, 10);

            string sFromDate = sDate.Substring(0, 7) + "-01";
            string sYYMM = sDate.Replace("-", "").Substring(0, 6);
            string sMM = sDate.Substring(5, 2);

            GridBandMM1.Caption = sMM+"월 실적";
            GridBandMM2.Caption = sMM+ "월 목표";

            StringBuilder strSql = new StringBuilder();

            //스크랩
            #region 목표 및 실적
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(");
            strSql.AppendLine("     /*월 매출 목표량 야드매입목표: o_gs 폐압목표: o_weight 야드매출목표: o_gs+o_sd+o_weight 직납매출목표:o_yk */");
            strSql.AppendLine("     SELECT '1' AS NUM                                                   ");
            strSql.AppendLine("          , SUM(O_GS) AS O_GS                                            ");
            strSql.AppendLine("          , SUM(O_WEIGHT) AS O_WEIGHT                                    ");
            strSql.AppendLine("          , SUM(O_GS) + SUM(O_WEIGHT) + SUM(O_SD) AS YADMCH              ");
            strSql.AppendLine("          , SUM(O_YK) AS O_YK                                            ");
            strSql.AppendLine("       FROM SALEMAECHUL                                                  ");
            strSql.AppendLine("      WHERE YYMM = '"+ sYYMM + "'                                              ");
            strSql.AppendLine(" ), TEMP2 AS(                                                            ");
            strSql.AppendLine("     /*야드매입월계*/                                                    ");
            strSql.AppendLine("     SELECT '2' AS NUM, SUM(Z1.DANJUNG) / 1000 AS DANJUNG                ");
            strSql.AppendLine("       FROM(                                                             ");
            strSql.AppendLine("               SELECT SUM(A.DANJUNG) AS DANJUNG                          ");
            strSql.AppendLine("               FROM INLIST A                                             ");
            strSql.AppendLine("               LEFT JOIN JAJAE B                                         ");
            strSql.AppendLine("                 ON A.J_SERIAL = B.J_SERIAL                              ");
            strSql.AppendLine("               LEFT JOIN ACC_DEALER_CD C                                 ");
            strSql.AppendLine("                 ON A.J_ID1 = C.DEALER_CD                                ");
            strSql.AppendLine("               LEFT JOIN HR_EMP_BASIS D                                  ");
            strSql.AppendLine("                 ON C.CHRG_ID = D.EMP_ID                                 ");
            strSql.AppendLine("              WHERE A.KERATYPE = '매입'                                  ");
            strSql.AppendLine("                AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sDate + "'       ");
            strSql.AppendLine("                AND A.J_LOTNO <> '4'                                     ");
            strSql.AppendLine("                AND B.DAEGUBUN <> '슈레더'                               ");
            strSql.AppendLine("                AND B.Gubun1 <> '인센티브'                               ");
            strSql.AppendLine("                AND D.EMP_NM <> '박봉섭'                                 ");
            strSql.AppendLine("              UNION ALL                                                  ");
            strSql.AppendLine("             SELECT SUM(A.DANJUNG) AS DANJUNG                            ");
            strSql.AppendLine("               FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E   ");
            strSql.AppendLine("              WHERE A.J_SERIAL = B.J_SERIAL                              ");
            strSql.AppendLine("                AND A.J_ID1 = D.DEALER_CD                                ");
            strSql.AppendLine("                AND D.CHRG_ID = E.EMP_ID                                 ");
            strSql.AppendLine("                AND A.KERATYPE = '매입'                                  ");
            strSql.AppendLine("                AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sDate + "'       ");
            strSql.AppendLine("                AND A.J_LOTNO <> '4'                                     ");
            strSql.AppendLine("                AND B.DAEGUBUN <> '슈레더'                               ");
            strSql.AppendLine("                AND B.GUBUN1 = '경량S'                                   ");
            strSql.AppendLine("       )Z1                                                               ");
            strSql.AppendLine(" ), TEMP3 AS(                                                            ");
            strSql.AppendLine("     /*폐압 월계*/                                                       ");
            strSql.AppendLine("     SELECT '2' AS NUM, SUM(A.DANJUNG) / 1000 AS DANJUNG                 ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E           ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("        AND A.J_ID1 = D.DEALER_CD                                        ");
            strSql.AppendLine("        AND D.CHRG_ID = E.EMP_ID                                         ");
            strSql.AppendLine("        AND A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sDate + "'               ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("        AND B.DAEGUBUN = '슈레더'                                        ");
            strSql.AppendLine("        AND B.Gubun1 <> '인센티브'                                       ");
            strSql.AppendLine(" ), TEMP4 AS(                                                            ");
            strSql.AppendLine("     /*야드매출 월계*/                                                   ");
            strSql.AppendLine("     SELECT '2' AS NUM, SUM(A.DANJUNG)/ 1000 AS DANJUNG                  ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                            ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("        AND A.J_ID1 = '6531121044'                                       ");
            strSql.AppendLine("        AND A.KERATYPE = '매출'                                          ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sFromDate + "'AND '"+ sDate + "'                ");
            strSql.AppendLine("        AND B.GUBUN1 <> '인센티브'                                       ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine(" ), TEMP5 AS(                                                            ");
            strSql.AppendLine("     /*YK직납 매출 월계*/                                                ");
            strSql.AppendLine("     SELECT '2' AS NUM, SUM(A.DANJUNG)/ 1000 AS DANJUNG                  ");
            strSql.AppendLine("       FROM INLIST A                                                     ");
            strSql.AppendLine("       LEFT JOIN JAJAE B                                                 ");
            strSql.AppendLine("         ON A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("      WHERE A.KERATYPE = '매출'                                          ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sFromDate + "' AND '"+ sDate + "'               ");
            strSql.AppendLine("        AND A.J_LOTNO = '4'                                              ");
            strSql.AppendLine("        AND B.Gubun1 <> '인센티브'                                       ");
            strSql.AppendLine(" )                                                                       ");
            strSql.AppendLine("                                                                         ");
            strSql.AppendLine(" /* 목표 및 월계*/                                                       ");
            strSql.AppendLine(" SELECT NUM                                                              ");
            strSql.AppendLine("      , O_GS                                                             ");
            strSql.AppendLine("      , O_WEIGHT                                                         ");
            strSql.AppendLine("      , YADMCH                                                           ");
            strSql.AppendLine("      , O_YK                                                             ");
            strSql.AppendLine("   FROM TEMP1                                                            ");
            strSql.AppendLine("  UNION ALL                                                              ");
            strSql.AppendLine(" SELECT T2.NUM                                                           ");
            strSql.AppendLine("      , T2.DANJUNG AS W1                                                 ");
            strSql.AppendLine("      , T3.DANJUNG AS W2                                                 ");
            strSql.AppendLine("      , T4.DANJUNG AS W3                                                 ");
            strSql.AppendLine("      , T5.DANJUNG AS W4                                                 ");
            strSql.AppendLine("   FROM TEMP2 T2                                                         ");
            strSql.AppendLine("   LEFT JOIN TEMP3 T3                                                    ");
            strSql.AppendLine("     ON T2.NUM = T3.NUM                                                  ");
            strSql.AppendLine("   LEFT JOIN TEMP4 T4                                                    ");
            strSql.AppendLine("     ON T2.NUM = T4.NUM                                                  ");
            strSql.AppendLine("   LEFT JOIN TEMP5 T5                                                    ");
            strSql.AppendLine("     ON T2.NUM = T5.NUM                                                  ");

            DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt1 != null)
            {
                for(int i = 0; i < dt1.Rows.Count; i++)
                {
                    string sNUM = dt1.Rows[i]["NUM"]?.ToString();
	                string sO_GS = dt1.Rows[i]["O_GS"]?.ToString();
	                string sO_WEIGHT= dt1.Rows[i]["O_WEIGHT"]?.ToString();
	                string sYADMCH = dt1.Rows[i]["YADMCH"]?.ToString();
                    string sO_YK = dt1.Rows[i]["O_YK"]?.ToString();

                    double.TryParse(sO_GS, out double dO_GS);
                    double.TryParse(sO_WEIGHT, out double dO_WEIGHT);
                    double.TryParse(sYADMCH, out double dYADMCH);
                    double.TryParse(sO_YK, out double dO_YK);

                    if (sNUM.Equals("1"))
                    {
                        TxtOgs1.EditValue = dO_GS;
                        TxtOW1.EditValue = dO_WEIGHT;
                        TxtYd1.EditValue = dYADMCH;
                        TxtYk1.EditValue = dO_YK;
                    }
                    else if (sNUM.Equals("2"))
                    {
                        TxtOgs2.EditValue = dO_GS;
                        TxtOW2.EditValue = dO_WEIGHT;
                        TxtYd2.EditValue = dYADMCH;
                        TxtYk2.EditValue = dO_YK;
                    }
                }
            }
            else
            {
                TxtOgs1.EditValue = null;
                TxtOgs2.EditValue = null;
                TxtOgs3.EditValue = null;

                TxtOW1.EditValue = null;
                TxtOW2.EditValue = null;
                TxtOW3.EditValue = null;

                TxtYd1.EditValue = null;
                TxtYd2.EditValue = null;
                TxtYd3.EditValue = null;

                TxtYk1.EditValue = null;
                TxtYk2.EditValue = null;
                TxtYk3.EditValue = null;
            }
            #endregion

            #region 담당자별 실적 및 목표
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF                      ");
            strSql.AppendLine(" SET ARITHIGNORE ON                         ");
            strSql.AppendLine(" SET ARITHABORT OFF;                        ");
            strSql.AppendLine("                                            ");
            strSql.AppendLine(" WITH TEMP6 AS(                             ");
            strSql.AppendLine("     /*담당자별 실적 집계*/                 ");
            strSql.AppendLine("     SELECT T5.NUM                          ");
            strSql.AppendLine("          , T5.NAME-- 담당자                ");
            strSql.AppendLine("          , T1.DANJUNG AS SCDAY-- 당일스크랩");
            strSql.AppendLine("          , T2.DANJUNG AS PEDAY-- 당일폐압  ");
            strSql.AppendLine("          , T3.DANJUNG AS SCMON-- 월스크랩  ");
            strSql.AppendLine("          , T4.DANJUNG AS PEMON-- 월폐압    ");
            strSql.AppendLine("          , T5.SCRAP-- 스크랩 목표          ");
            strSql.AppendLine("          , T5.PEAP-- 폐압 목표             ");
            strSql.AppendLine("          , T3.DANJUNG / T5.SCRAP * 100 AS DALPER_S-- 달성률 스크랩                  ");
            strSql.AppendLine("          , T4.DANJUNG / T5.PEAP * 100 AS DALPER_P-- 달성률 폐압                     ");
            strSql.AppendLine("          , ISNULL(T3.DANJUNG, 0) - ISNULL(T1.DANJUNG, 0) AS ATSIL_S--전일실적 스크랩");
            strSql.AppendLine("          , ISNULL(T4.DANJUNG, 0) - ISNULL(T2.DANJUNG, 0) AS ATSIL_P--전일실적 폐압  ");
            strSql.AppendLine("                                                                                     ");
            strSql.AppendLine("       FROM( /*매입 월 목표량 스크랩목표: I_LIGHT + I_WEIGHT , 폐압:I_CHAPI  */      ");
            strSql.AppendLine("             SELECT CASE WHEN NAME = '손상영' THEN 1                                 ");
            strSql.AppendLine("                         WHEN NAME = '오상훈' THEN 2                                 ");
            strSql.AppendLine("                         WHEN NAME = '김명철' THEN 3                                 ");
            strSql.AppendLine("                         WHEN NAME = '이우택' THEN 4                                 ");
            strSql.AppendLine("                         WHEN NAME = '박봉섭' THEN 5                                 ");
            strSql.AppendLine("                         WHEN NAME = '기타' THEN 6 END AS NUM                        ");
            strSql.AppendLine("                  , NAME                                                             ");
            strSql.AppendLine("                  , SUM(I_LIGHT) * 1000 + SUM(I_WEIGHT) * 1000 AS SCRAP              ");
            strSql.AppendLine("                  , SUM(I_CHAPI) * 1000 AS PEAP                                      ");
            strSql.AppendLine("               FROM SALEMAEIP                                                        ");
            strSql.AppendLine("              WHERE YYMM = '"+ sYYMM + "'                                                  ");
            strSql.AppendLine("              GROUP BY NAME) T5                                                      ");
            strSql.AppendLine("       LEFT JOIN(/*담당자별 당일 스크랩 실적*/                                       ");
            strSql.AppendLine("                 SELECT CASE WHEN E.EMP_NM = '손상영' THEN 1                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '오상훈' THEN 2                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '김명철' THEN 3                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '이우택' THEN 4                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '박봉섭' THEN 5                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '기타' THEN 6 END AS NUM                ");
            strSql.AppendLine("                      , E.EMP_NM                                                     ");
            strSql.AppendLine("                      , SUM(A.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                   FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E           ");
            strSql.AppendLine("                  WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                    AND A.J_ID1 = D.DEALER_CD                                        ");
            strSql.AppendLine("                    AND D.CHRG_ID = E.EMP_ID                                         ");
            strSql.AppendLine("                    AND A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                    AND A.J_DATE = '"+ sDate + "'                                      ");
            strSql.AppendLine("                    AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                    AND B.DAEGUBUN <> '슈레더'                                       ");
            strSql.AppendLine("                    AND B.GUBUN1 <> '인센티브'                                       ");
            strSql.AppendLine("                    AND E.EMP_NM <> '김홍철'                                         ");
            strSql.AppendLine("                  GROUP BY E.EMP_NM) T1                                              ");
            strSql.AppendLine("         ON T5.NUM = T1.NUM                                                          ");
            strSql.AppendLine("       LEFT JOIN(/*담당자별 당일 폐압 실적*/                                         ");
            strSql.AppendLine("                 SELECT CASE WHEN E.EMP_NM = '손상영' THEN 1                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '오상훈' THEN 2                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '김명철' THEN 3                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '이우택' THEN 4                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '박봉섭' THEN 5                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '기타' THEN 6 END AS NUM                ");
            strSql.AppendLine("                      , E.EMP_NM                                                     ");
            strSql.AppendLine("                      , SUM(A.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                   FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E           ");
            strSql.AppendLine("                  WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                    AND A.J_ID1 = D.DEALER_CD                                        ");
            strSql.AppendLine("                    AND D.CHRG_ID = E.EMP_ID                                         ");
            strSql.AppendLine("                    AND A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                    AND A.J_DATE = '"+sDate+"'                                      ");
            strSql.AppendLine("                    AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                    AND B.DAEGUBUN = '슈레더'                                        ");
            strSql.AppendLine("                    AND B.GUBUN1 <> '인센티브'                                       ");
            strSql.AppendLine("                    AND E.EMP_NM <> '김홍철'                                         ");
            strSql.AppendLine("                  GROUP BY E.EMP_NM) T2                                              ");
            strSql.AppendLine("         ON T5.NUM = T2.NUM                                                          ");
            strSql.AppendLine("       LEFT JOIN(/*담당자별 월 스크랩실적*/                                          ");
            strSql.AppendLine("                 SELECT CASE WHEN E.EMP_NM = '손상영' THEN 1                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '오상훈' THEN 2                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '김명철' THEN 3                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '이우택' THEN 4                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '박봉섭' THEN 5                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '기타' THEN 6 END AS NUM                ");
            strSql.AppendLine("                      , E.EMP_NM                                                     ");
            strSql.AppendLine("                      , SUM(A.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                   FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E           ");
            strSql.AppendLine("                  WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                    AND A.J_ID1 = D.DEALER_CD                                        ");
            strSql.AppendLine("                    AND D.CHRG_ID = E.EMP_ID                                         ");
            strSql.AppendLine("                    AND A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                    AND A.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'               ");
            strSql.AppendLine("                    AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                    AND B.DAEGUBUN <> '슈레더'                                       ");
            strSql.AppendLine("                    AND B.GUBUN1 <> '인센티브'                                       ");
            strSql.AppendLine("                    AND E.EMP_NM <> '김홍철'                                         ");
            strSql.AppendLine("                  GROUP BY E.EMP_NM) T3                                              ");
            strSql.AppendLine("         ON T5.NUM = T3.NUM                                                          ");
            strSql.AppendLine("       LEFT JOIN(/*담당자별 월 폐압 실적*/                                           ");
            strSql.AppendLine("                 SELECT CASE WHEN E.EMP_NM = '손상영' THEN 1                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '오상훈' THEN 2                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '김명철' THEN 3                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '이우택' THEN 4                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '박봉섭' THEN 5                         ");
            strSql.AppendLine("                             WHEN E.EMP_NM = '기타' THEN 6 END AS NUM                ");
            strSql.AppendLine("                      , E.EMP_NM                                                     ");
            strSql.AppendLine("                      , SUM(A.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                   FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E           ");
            strSql.AppendLine("                  WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                    AND A.J_ID1 = D.DEALER_CD                                        ");
            strSql.AppendLine("                    AND D.CHRG_ID = E.EMP_ID                                         ");
            strSql.AppendLine("                    AND A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                    AND A.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'               ");
            strSql.AppendLine("                    AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                    AND B.DAEGUBUN = '슈레더'                                        ");
            strSql.AppendLine("                    AND B.GUBUN1 <> '인센티브'                                       ");
            strSql.AppendLine("                    AND E.EMP_NM <> '김홍철'                                         ");
            strSql.AppendLine("                  GROUP BY E.EMP_NM) T4                                              ");
            strSql.AppendLine("         ON T5.NUM = T4.NUM                                                          ");
            strSql.AppendLine("     ), TEMP13 AS(                                                                   ");
            strSql.AppendLine("         /*기타,경량S,수입*/                                                         ");
            strSql.AppendLine("         SELECT 6 AS NUM                                                             ");
            strSql.AppendLine("              , '스크랩계' AS NAME                                                   ");
            strSql.AppendLine("              , SUM(SCDAY) AS SCDAY                                                  ");
            strSql.AppendLine("              , SUM(PEDAY) AS PEDAY                                                  ");
            strSql.AppendLine("              , SUM(SCMON) AS SCMON                                                  ");
            strSql.AppendLine("              , SUM(PEMON) AS PEMON                                                  ");
            strSql.AppendLine("              , SUM(SCRAP) AS SCRAP                                                  ");
            strSql.AppendLine("              , SUM(PEAP) AS PEAP                                                    ");
            strSql.AppendLine("              , NULL AS DALPER_S                                                     ");
            strSql.AppendLine("              , NULL AS DALPER_P                                                     ");
            strSql.AppendLine("              , SUM(ATSIL_S) AS ATSIL_S                                              ");
            strSql.AppendLine("              , SUM(ATSIL_P) AS ATSIL_P                                              ");
            strSql.AppendLine("           FROM TEMP6                                                                ");
            strSql.AppendLine("          WHERE NUM NOT IN(5,6)                                                      ");
            strSql.AppendLine("          UNION ALL                                                                  ");
            strSql.AppendLine("         SELECT 7 AS NUM                                                             ");
            strSql.AppendLine("              , NAME --담당자                                                        ");
		    strSql.AppendLine("              , SCDAY-- 당일스크랩                                                   ");
		    strSql.AppendLine("              , PEDAY-- 당일폐압                                                     ");
		    strSql.AppendLine("              , SCMON-- 월스크랩                                                     ");
		    strSql.AppendLine("              , PEMON-- 월폐압                                                       ");
		    strSql.AppendLine("              , SCRAP-- 스크랩 목표                                                  ");
            strSql.AppendLine("              , PEAP --폐압 목표                                                     ");
		    strSql.AppendLine("              , DALPER_S-- 달성률 스크랩                                             ");
            strSql.AppendLine("              , DALPER_P --달성률 폐압                                               ");
		    strSql.AppendLine("              , ATSIL_S--전일실적 스크랩                                             ");
            strSql.AppendLine("              , ATSIL_P --전일실적 폐압                                              ");
            strSql.AppendLine("           FROM TEMP6                                                                ");
            strSql.AppendLine("          WHERE NUM = 6                                                              ");
            strSql.AppendLine("          UNION ALL                                                                  ");
            strSql.AppendLine("         SELECT 8 AS NUM                                                             ");
            strSql.AppendLine("              , '경량S'                                                              ");
	        strSql.AppendLine("              , A1.DANJUNG                                                           ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , A2.DANJUNG                                                           ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , ISNULL(A2.DANJUNG, 0) - ISNULL(A1.DANJUNG, 0)--전일실적 스크랩       ");
            strSql.AppendLine("              , NULL--전일실적 폐압                                                  ");
            strSql.AppendLine("           FROM(/*당일 경량S*/                                                       ");
            strSql.AppendLine("                 SELECT A.GUBUN1                                                     ");
            strSql.AppendLine("                      , SUM(B.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                   FROM JAJAE A                                                      ");
            strSql.AppendLine("                   LEFT JOIN INLIST B                                                ");
            strSql.AppendLine("                     ON A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                    AND B.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                    AND B.J_DATE = '"+sDate+"'                                      ");
            strSql.AppendLine("                    AND B.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                  WHERE A.GUBUN1 = '경량S'                                           ");
            strSql.AppendLine("                    AND A.DAEGUBUN <> '슈레더'                                       ");
            strSql.AppendLine("                  GROUP BY A.GUBUN1) A1                                              ");
            strSql.AppendLine("           LEFT JOIN(/*월 경량S*/                                                    ");
            strSql.AppendLine("                     SELECT A.GUBUN1                                                 ");
            strSql.AppendLine("                          , SUM(B.DANJUNG) AS DANJUNG                                ");
            strSql.AppendLine("                       FROM JAJAE A                                                  ");
            strSql.AppendLine("                       LEFT JOIN INLIST B                                            ");
            strSql.AppendLine("                         ON A.J_SERIAL = B.J_SERIAL                                  ");
            strSql.AppendLine("                        AND B.KERATYPE = '매입'                                      ");
            strSql.AppendLine("                        AND B.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'           ");
            strSql.AppendLine("                        AND B.J_LOTNO <> '4'                                         ");
            strSql.AppendLine("                      WHERE A.GUBUN1 = '경량S'                                       ");
            strSql.AppendLine("                        AND A.DAEGUBUN <> '슈레더'                                   ");
            strSql.AppendLine("                      GROUP BY A.GUBUN1) A2                                          ");
            strSql.AppendLine("             ON A1.GUBUN1 = A2.GUBUN1                                                ");
            strSql.AppendLine("          UNION ALL                                                                  ");
            strSql.AppendLine("         SELECT 9 AS NUM                                                             ");
            strSql.AppendLine("              , '수입'                                                               ");
	        strSql.AppendLine("              , A1.DANJUNG                                                           ");
		    strSql.AppendLine("              , A2.DANJUNG                                                           ");
		    strSql.AppendLine("              , A3.DANJUNG                                                           ");
		    strSql.AppendLine("              , A4.DANJUNG                                                           ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , NULL                                                                 ");
		    strSql.AppendLine("              , ISNULL(A3.DANJUNG, 0) - ISNULL(A1.DANJUNG, 0)                        ");
		    strSql.AppendLine("              , ISNULL(A4.DANJUNG, 0) - ISNULL(A2.DANJUNG, 0)                        ");
            strSql.AppendLine("           FROM(/*당일 수입스크랩: 수입경량,수입선반,수입압축,수입철사 */            ");
            strSql.AppendLine("                 SELECT '수입' AS GUBUN1                                             ");
            strSql.AppendLine("                      , SUM(A.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                   FROM INLIST A                                                     ");
            strSql.AppendLine("                   LEFT JOIN JAJAE B                                                 ");
            strSql.AppendLine("                     ON A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                  WHERE A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                    AND A.J_DATE = '"+sDate+"'                                      ");
            strSql.AppendLine("                    AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                    AND B.DAEGUBUN <> '슈레더'                                       ");
            strSql.AppendLine("                    AND B.GUBUN1 IN('수입경량', '수입선반', '수입압축', '수입철사')) A1");
            strSql.AppendLine("            LEFT JOIN(/*당일 수입폐압: 수입모재*/                                      ");
            strSql.AppendLine("                      SELECT '수입' AS GUBUN1                                          ");
            strSql.AppendLine("                           , SUM(A.DANJUNG) AS DANJUNG                                 ");
            strSql.AppendLine("                        FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E        ");
            strSql.AppendLine("                       WHERE A.J_SERIAL = B.J_SERIAL                                   ");
            strSql.AppendLine("                         AND A.J_ID1 = D.DEALER_CD                                     ");
            strSql.AppendLine("                         AND D.CHRG_ID = E.EMP_ID                                      ");
            strSql.AppendLine("                         AND A.KERATYPE = '매입'                                       ");
            strSql.AppendLine("                         AND A.J_DATE = '"+sDate+"'                                   ");
            strSql.AppendLine("                         AND A.J_LOTNO <> '4'                                          ");
            strSql.AppendLine("                         AND B.DAEGUBUN = '슈레더'                                     ");
            strSql.AppendLine("                         AND B.GUBUN1 IN('수입모재')) A2                               ");
            strSql.AppendLine("            ON A1.GUBUN1 = A2.GUBUN1                                                   ");
            strSql.AppendLine("           LEFT JOIN(/*월 수입스크랩: 수입경량,수입선반,수입압축,수입철사 */           ");
            strSql.AppendLine("                     SELECT '수입' AS GUBUN1                                           ");
            strSql.AppendLine("                          , SUM(A.DANJUNG) AS DANJUNG                                  ");
            strSql.AppendLine("                       FROM INLIST A                                                   ");
            strSql.AppendLine("                       LEFT JOIN JAJAE B                                               ");
            strSql.AppendLine("                         ON A.J_SERIAL = B.J_SERIAL                                    ");
            strSql.AppendLine("                      WHERE A.KERATYPE = '매입'                                        ");
            strSql.AppendLine("                        AND A.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'             ");
            strSql.AppendLine("                        AND A.J_LOTNO <> '4'                                           ");
            strSql.AppendLine("                        AND B.DAEGUBUN <> '슈레더'                                     ");
            strSql.AppendLine("                        AND B.GUBUN1 IN('수입경량','수입선반','수입압축','수입철사') ) A3");
            strSql.AppendLine("            ON A1.GUBUN1 = A3.GUBUN1                                                     ");
            strSql.AppendLine("           LEFT JOIN(/*월 수입폐압: 수입모재*/                                           ");
            strSql.AppendLine("                     SELECT '수입' AS GUBUN1                                             ");
            strSql.AppendLine("                          , SUM(A.DANJUNG) AS DANJUNG                                    ");
            strSql.AppendLine("                       FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E           ");
            strSql.AppendLine("                      WHERE A.J_SERIAL = B.J_SERIAL                                      ");
            strSql.AppendLine("                        AND A.J_ID1 = D.DEALER_CD                                        ");
            strSql.AppendLine("                        AND D.CHRG_ID = E.EMP_ID                                         ");
            strSql.AppendLine("                        AND A.KERATYPE = '매입'                                          ");
            strSql.AppendLine("                        AND A.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'               ");
            strSql.AppendLine("                        AND A.J_LOTNO <> '4'                                             ");
            strSql.AppendLine("                        AND B.DAEGUBUN = '슈레더'                                        ");
            strSql.AppendLine("                        AND B.GUBUN1 IN('수입모재')) A4                                  ");
            strSql.AppendLine("            ON A1.GUBUN1 = A4.GUBUN1                                                     ");
            strSql.AppendLine("     )                                                                                   ");
            strSql.AppendLine("                                                                                         ");
            strSql.AppendLine("     SELECT NUM                                                                          ");
            strSql.AppendLine("          , NAME                                                                         ");
            strSql.AppendLine("          , SCDAY                                                                        ");
            strSql.AppendLine("          , PEDAY                                                                        ");
            strSql.AppendLine("          , SCMON                                                                        ");
            strSql.AppendLine("          , PEMON                                                                        ");
            strSql.AppendLine("          , SCRAP                                                                        ");
            strSql.AppendLine("          , PEAP                                                                         ");
            strSql.AppendLine("          , CASE WHEN DALPER_S IS NULL OR DALPER_S = 0 THEN NULL ELSE CONCAT(CAST(ROUND(DALPER_S,0) AS VARCHAR),'%') END AS DALPER_S                                                ");
            strSql.AppendLine("          , CASE WHEN DALPER_P IS NULL OR DALPER_P = 0 THEN NULL ELSE CONCAT(CAST(ROUND(DALPER_P,0) AS VARCHAR),'%') END AS DALPER_P                                                ");
            strSql.AppendLine("          , ATSIL_S                                                                      ");
            strSql.AppendLine("          , ATSIL_P                                                                      ");
            strSql.AppendLine("       FROM(SELECT NUM                                                                   ");
            strSql.AppendLine("                  , NAME                                                                 ");
            strSql.AppendLine("                  , SCDAY                                                                ");
            strSql.AppendLine("                  , PEDAY                                                                ");
            strSql.AppendLine("                  , SCMON                                                                ");
            strSql.AppendLine("                  , PEMON                                                                ");
            strSql.AppendLine("                  , SCRAP                                                                ");
            strSql.AppendLine("                  , PEAP                                                                 ");
            strSql.AppendLine("                  , DALPER_S                                                             ");
            strSql.AppendLine("                  , DALPER_P                                                             ");
            strSql.AppendLine("                  , ATSIL_S                                                              ");
            strSql.AppendLine("                  , ATSIL_P                                                              ");
            strSql.AppendLine("               FROM TEMP6                                                                ");
            strSql.AppendLine("              WHERE NUM <> 6                                                             ");
            strSql.AppendLine("              UNION ALL                                                                  ");
            strSql.AppendLine("             SELECT NUM                                                                  ");
            strSql.AppendLine("                  , NAME                                                                 ");
            strSql.AppendLine("                  , SCDAY                                                                ");
            strSql.AppendLine("                  , PEDAY                                                                ");
            strSql.AppendLine("                  , SCMON                                                                ");
            strSql.AppendLine("                  , PEMON                                                                ");
            strSql.AppendLine("                  , SCRAP                                                                ");
            strSql.AppendLine("                  , PEAP                                                                 ");
            strSql.AppendLine("                  , DALPER_S                                                             ");
            strSql.AppendLine("                  , DALPER_P                                                             ");
            strSql.AppendLine("                  , ATSIL_S                                                              ");
            strSql.AppendLine("                  , ATSIL_P                                                              ");
            strSql.AppendLine("               FROM TEMP13                                                               ");
            strSql.AppendLine("              UNION ALL                                                                  ");
            strSql.AppendLine("             SELECT 10 AS NUM                                                            ");
            strSql.AppendLine("                  , '합 계'                                                              ");
            strSql.AppendLine("                  , SUM(SCDAY)-- 당일스크랩                                              ");
            strSql.AppendLine("                  , SUM(PEDAY)-- 당일폐압                                                ");
            strSql.AppendLine("                  , SUM(SCMON)-- 월스크랩                                                ");
            strSql.AppendLine("                  , SUM(PEMON)-- 월폐압                                                  ");
            strSql.AppendLine("                  , SUM(SCRAP)-- 스크랩 목표                                             ");
            strSql.AppendLine("                  , SUM(PEAP)-- 폐압 목표                                                ");
            strSql.AppendLine("                  , NULL-- 달성률 스크랩                                                 ");
            strSql.AppendLine("                  , NULL-- 달성률 폐압                                                   ");
            strSql.AppendLine("                  , SUM(CASE WHEN NUM != 6 THEN ATSIL_S END)--전일실적 스크랩            ");
            strSql.AppendLine("                  , SUM(CASE WHEN NUM != 6 THEN ATSIL_P END)--전일실적 폐압              ");
            strSql.AppendLine("               FROM TEMP13) Z1                                                           ");
            strSql.AppendLine("      ORDER BY Z1.NUM                                                                    ");

            DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr1.DataSource = dt2;
            #endregion

            #region 스크랩 매입
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" SELECT B3.JUNPYOID--전표번호 ");
            strSql.AppendLine("      , B3.J_DATE             ");
            strSql.AppendLine("      , A1.J_BNUM--차량번호   ");
	        strSql.AppendLine("      , B2.DEALER_NM--업체명  ");
            strSql.AppendLine("      , A1.GUBUN1--등급       ");
            strSql.AppendLine("      , B3.DANJUNG--인수량    ");
            strSql.AppendLine("      , ISNULL(C1.CHAGAM, ISNULL(A2.CHAGAM, CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END)) AS FIRST_CHAGAM --검수감량");
            strSql.AppendLine("      , B3.HALIN-- 조정후감량                                              ");
            strSql.AppendLine("      , B3.MIDANGA-- TABLE단가                                             ");
            strSql.AppendLine("      , B3.DANGA--지급단가                                                 ");
            strSql.AppendLine("      , B2.J_REGION--지역                                                  ");
            strSql.AppendLine("      , B3.CKONGKEP / B3.DANJUNG AS CKONGKEP --운반비(운반비단가)          ");
            strSql.AppendLine("      , B3.DANGA + (B3.CKONGKEP / B3.DANJUNG) AS MIPDAN --매입단가         ");
            strSql.AppendLine("      , B3.MIDANGA - (B3.DANGA + (B3.CKONGKEP / B3.DANJUNG)) AS CHA --차액 ");
            strSql.AppendLine("      , A1.GUMSUBIGO--검수판정                                             ");
            strSql.AppendLine("      , B3.CHRG_NM--담당자                                                 ");
            strSql.AppendLine("      , E1.DAN AS CHRDAN--담당단가                                         ");
            strSql.AppendLine("      , '' AS DTEXT --단가보완사유                                         ");
            strSql.AppendLine("      , B3.IKONGKEP--지급금액                                              ");
            strSql.AppendLine("      , ISNULL(B3.IKONGKEP, 0) + ISNULL(B3.CKONGKEP, 0) AS MAIPK --매입금액");
            strSql.AppendLine("   FROM MESURING A1                                                        ");
            strSql.AppendLine("   LEFT JOIN MESURING_SEQ A2                                               ");
            strSql.AppendLine("     ON A1.JUNPYOID = A2.JUNPYOID                                          ");
            strSql.AppendLine("    AND A2.CHG_SEQ = (SELECT MIN(X1.CHG_SEQ) FROM MESURING_SEQ X1 WHERE X1.JUNPYOID = A1.JUNPYOID )"); 
            strSql.AppendLine("   LEFT JOIN INLIST B3                                                                             ");
            strSql.AppendLine("     ON A1.IPCHULGO_MAIPID = B3.J_ID                                                               ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1                                                                       ");
            strSql.AppendLine("     ON A1.GUMSU_SERIAL = B1.EMP_ID                                                                ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B2                                                                      ");
            strSql.AppendLine("     ON A1.MAIPCHERID = B2.DEALER_CD                                                               ");
            strSql.AppendLine("   LEFT JOIN MESURE_ISPT_INFO C1                                                                   ");
            strSql.AppendLine("     ON A1.JUNPYOID = C1.JUNPYOID                                                                  ");
            strSql.AppendLine("   LEFT JOIN(SELECT X1.USRCD, X2.EMP_NM FROM ZUSRLST X1                                            ");
            strSql.AppendLine("                   LEFT JOIN HR_EMP_BASIS X2 ON X1.INSANO = X2.EMP_ID) C2                          ");
            strSql.AppendLine("    ON C1.ENT_ID = C2.USRCD                                                                        ");
            strSql.AppendLine("   LEFT JOIN JAJAE D1                                                                              ");
            strSql.AppendLine("     ON A1.J_SERIAL = D1.J_SERIAL                                                                  ");
            strSql.AppendLine("   LEFT JOIN(SELECT E.EMP_NM                                                                       ");
            strSql.AppendLine("                   , (SUM(A.CKONGKEP) +SUM(A.IKONGKEP))/ SUM(A.DANJUNG) AS DAN                     ");
            strSql.AppendLine("                FROM INLIST A, JAJAE B, MESURING C, ACC_DEALER_CD D, HR_EMP_BASIS E                ");
            strSql.AppendLine("               WHERE A.J_SERIAL = B.J_SERIAL                                                       ");
            strSql.AppendLine("                 AND A.J_ID = C.IPCHULGO_MAIPID                                                    ");
            strSql.AppendLine("                 AND A.J_ID1 = D.DEALER_CD                                                         ");
            strSql.AppendLine("                 AND D.CHRG_ID = E.EMP_ID                                                          ");
            strSql.AppendLine("                 AND A.KERATYPE = '매입'                                                           ");
            strSql.AppendLine("                 AND A.J_DATE = '"+sDate+"'                                                       ");
            strSql.AppendLine("                 AND A.J_LOTNO <> '4'                                                              ");
            strSql.AppendLine("                 AND B.DAEGUBUN <> '슈레더'                                                        ");
            strSql.AppendLine("               GROUP BY E.EMP_NM) E1                                                               ");
            strSql.AppendLine("     ON B3.CHRG_NM = E1.EMP_NM                                                                     ");
            strSql.AppendLine("  WHERE B3.KERATYPE = '매입'                                                                       ");
            strSql.AppendLine("    AND B3.J_DATE = '"+sDate+"'                                                                   ");
            strSql.AppendLine("    AND B3.J_LOTNO <> 4                                                                            ");
            strSql.AppendLine("    AND D1.DAEGUBUN <> '슈레더'                                                                    ");
            strSql.AppendLine("    AND D1.GUBUN1 <> '인센티브'                                                                    ");
            strSql.AppendLine("  ORDER BY B3.CHRG_NM, B2.DEALER_NM ASC                                                            ");

            DataTable dt3 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr2.DataSource = dt3;
            #endregion

            #region 어음할인율
            strSql.Clear();
            strSql.AppendLine(" SELECT DATAVALUE AS AHALIN                                                                                                          ");
            strSql.AppendLine("   FROM MEAIPSILJUK                                                                                                                 ");
            strSql.AppendLine("  WHERE J_DATE = (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE DAEGUBUN = '기초자료' AND GUBUN = '어음할인' AND J_DATE <= '" + sDate + "')");
            strSql.AppendLine("    AND DAEGUBUN = '기초자료'                                                                                                       ");
            strSql.AppendLine("    AND GUBUN = '어음할인'                                                                                                          ");

            DataTable dtAHALIN = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            double dAHALIN = 0;
            if (dtAHALIN != null)
            {
                double.TryParse(dtAHALIN.Rows[0]["AHALIN"]?.ToString(), out dAHALIN);
                TxtAHALIN.EditValue = dAHALIN;
            }
            else
            {
                TxtAHALIN.EditValue = null;
            }
            #endregion

            #region 집계
            //어음할인율 dAHALIN
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" DECLARE @SUMILMECH1 NUMERIC                                ");
            strSql.AppendLine("       , @SUMILMECH2 NUMERIC                                ");
            strSql.AppendLine("       , @SUMWILMECH1 NUMERIC                               ");
            strSql.AppendLine("       , @SUMWILMECH2 NUMERIC                               ");
            strSql.AppendLine("       , @ILMECH1 NUMERIC                                   ");
            strSql.AppendLine("       , @ILMECH2 NUMERIC                                   ");
            strSql.AppendLine("       , @WLMECH1 NUMERIC                                   ");
            strSql.AppendLine("       , @WLMECH2 NUMERIC;                                  ");
            strSql.AppendLine("                                                            ");
            strSql.AppendLine("  /*마진률 계산*/                                           ");
            strSql.AppendLine("  /*YK 일매출 전체*/                                        ");
            strSql.AppendLine("  SELECT @SUMILMECH1 = SUM(A.DANJUNG)                       ");
            strSql.AppendLine("       , @SUMILMECH2 = SUM(A.KONGKEP) - SUM(A.CKONGKEP)     ");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                  ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                            ");
            strSql.AppendLine("     AND A.KERATYPE = '매출'                                ");
            strSql.AppendLine("     AND A.J_DATE = '"+sDate+"'                            ");
            strSql.AppendLine("     AND A.J_ID1 = '6531121044'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                             ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                   ");
            strSql.AppendLine("                                                            ");
            strSql.AppendLine("  /*YK 월매출 전체*/                                        ");
            strSql.AppendLine("  SELECT @SUMWILMECH1 = SUM(A.DANJUNG)                      ");
            strSql.AppendLine("       , @SUMWILMECH2 = SUM(A.KONGKEP) - SUM(A.CKONGKEP)    ");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                  ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                            ");
            strSql.AppendLine("     AND A.J_ID1 = '6531121044'                             ");
            strSql.AppendLine("     AND A.KERATYPE = '매출'                                ");
            strSql.AppendLine("     AND A.J_DATE >= '"+sFromDate+"'                           ");
            strSql.AppendLine("     AND A.J_DATE <= '"+sDate+"'                           ");
            strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                             ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                   ");
            strSql.AppendLine("  /*YK 일매출 G/S*/                                         ");
            strSql.AppendLine("  SELECT @ILMECH1 = SUM(A.DANJUNG)                          ");
            strSql.AppendLine("       , @ILMECH2 = SUM(A.KONGKEP) - SUM(A.CKONGKEP)        ");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                  ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                            ");
            strSql.AppendLine("     AND A.KERATYPE = '매출'                                ");
            strSql.AppendLine("     AND A.J_DATE = '"+sDate+"'                            ");
            strSql.AppendLine("     AND A.J_ID1 = '6531121044'                             ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                   ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '슈레더'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철A'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철B'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 AH'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 AL'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량A'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 AL'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 A-B'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량-ABL'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 AH'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량_ABH'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량B'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                             ");
            strSql.AppendLine("  /*YK 월매출 G/S*/                                         ");
            strSql.AppendLine("  SELECT @WLMECH1 = SUM(A.DANJUNG)                          ");
            strSql.AppendLine("       , @WLMECH2 = SUM(A.KONGKEP) - SUM(A.CKONGKEP)        ");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                  ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                            ");
            strSql.AppendLine("     AND A.KERATYPE = '매출'                                ");
            strSql.AppendLine("     AND A.J_DATE >= '"+sFromDate+"'                           ");
            strSql.AppendLine("     AND A.J_DATE <= '"+sDate+"'                           ");
            strSql.AppendLine("     AND A.J_ID1 = '6531121044'                             ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                   ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '슈레더'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철A'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철B'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 AH'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 AL'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량A'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 AL'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 A-B'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량-ABL'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 AH'                              ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량_ABH'                             ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량B'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                             ");
            strSql.AppendLine("                                                            ");
            strSql.AppendLine("  /*스크랩 일계(수입,저가 제외)  지급단가: iKongKep/danjung 소숫점2자리 */");
            strSql.AppendLine("  SELECT '스크랩 일계(수입및저가제외)' AS GB --구분                        ");
            strSql.AppendLine("       , SUM(B3.DANJUNG) AS DANJUNG --인수량                               ");
            strSql.AppendLine("       , SUM(ISNULL(C1.CHAGAM, ISNULL(A2.CHAGAM, CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END))) AS FIRST_CHAGAM --검수감량");
            strSql.AppendLine("       , SUM(B3.HALIN) AS HALIN --조정후감량                                                        ");
            strSql.AppendLine("       , SUM(B3.IKONGKEP) / SUM(B3.DANJUNG) AS GIDAN --지급단가                                     ");
            strSql.AppendLine("       , (SUM(B3.IKONGKEP) + SUM(B3.CKONGKEP)) / SUM(B3.DANJUNG) AS AVGDAN --평균단가                 ");
            strSql.AppendLine("       , NULL AS MGIN--마진율                                                                       ");
            strSql.AppendLine("       , (SUM(B3.IKONGKEP) + SUM(B3.CKONGKEP)) / SUM(B3.DANJUNG) AS MIPDAN --매입단가                 ");
            strSql.AppendLine("       , NULL AS MGDAN-- 마진단가                                                                   ");
            strSql.AppendLine("       , SUM(B3.IKONGKEP) AS IKONGKEP --지급금액                                                    ");
            strSql.AppendLine("       , SUM(B3.IKONGKEP) + SUM(B3.CKONGKEP) AS MKONGKEP --매입금액                                 ");
            strSql.AppendLine("    FROM MESURING A1                                                                                ");
            strSql.AppendLine("    LEFT JOIN MESURING_SEQ A2 ON A1.JUNPYOID = A2.JUNPYOID                                          ");
            strSql.AppendLine("     AND A2.CHG_SEQ = (SELECT MIN(X1.CHG_SEQ) FROM MESURING_SEQ X1 WHERE X1.JUNPYOID = A1.JUNPYOID )"); 
            strSql.AppendLine("    LEFT JOIN INLIST B3 ON A1.IPCHULGO_MAIPID = B3.J_ID                                             ");
            strSql.AppendLine("    LEFT JOIN MESURE_ISPT_INFO C1 ON A1.JUNPYOID = C1.JUNPYOID                                      ");
            strSql.AppendLine("    LEFT JOIN(SELECT X1.USRCD, X2.EMP_NM                                                            ");
            strSql.AppendLine("                  FROM ZUSRLST X1                                                                   ");
            strSql.AppendLine("                  LEFT JOIN HR_EMP_BASIS X2                                                         ");
            strSql.AppendLine("                    ON X1.INSANO = X2.EMP_ID) C2                                                    ");
            strSql.AppendLine("     ON C1.ENT_ID = C2.USRCD                                                                        ");
            strSql.AppendLine("    LEFT JOIN JAJAE D1 ON A1.J_SERIAL = D1.J_SERIAL                                                 ");
            strSql.AppendLine("   WHERE B3.KERATYPE = '매입'                                                                       ");
            strSql.AppendLine("     AND B3.J_DATE = '"+sDate+"'                                                                   ");
            strSql.AppendLine("     AND B3.J_LOTNO <> '4'                                                                          ");
            strSql.AppendLine("     AND D1.DAEGUBUN <> '슈레더'                                                                    ");
            strSql.AppendLine("     AND D1.GUBUN1 <> '인센티브'                                                                    ");
            strSql.AppendLine("     AND D1.DAEGUBUN <> '수입경량'                                                                  ");
            strSql.AppendLine("     AND D1.DAEGUBUN <> '수입철사'                                                                  ");
            strSql.AppendLine("     AND D1.DAEGUBUN <> '수입압축'                                                                  ");
            strSql.AppendLine("   UNION ALL                                                                                        ");
            strSql.AppendLine("  /*스크랩 일계 전체*/                                                                              ");
            strSql.AppendLine("  SELECT '스크랩 일계(전체)' AS GB --구분                                                           ");
            strSql.AppendLine("       , SUM(A.DANJUNG) AS DANJUNG --인수량                                                         ");
            strSql.AppendLine("       , NULL AS FIRST_CHAGAM--검수감량                                                             ");
            strSql.AppendLine("       , SUM(A.HALIN) AS HALIN --조정후감량                                                         ");
            strSql.AppendLine("       , SUM(A.IKONGKEP) / SUM(A.DANJUNG) AS GIDAN --지급단가                                       ");
            strSql.AppendLine("       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --평균단가                    ");
            strSql.AppendLine("       , (-1) * (1 - ((@SUMILMECH2 / @SUMILMECH1) - "+ dAHALIN + ") / ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)))   AS MGIN --마진율");
            strSql.AppendLine("       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --매입단가                                ");
            strSql.AppendLine("       ,  ((@ILMECH2 / @ILMECH1) - "+ dAHALIN + ") - (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MGDAN --마진단가");
            strSql.AppendLine("       , SUM(A.IKONGKEP) AS IKONGKEP --지급금액                                                                 ");
            strSql.AppendLine("       , SUM(A.IKONGKEP) + SUM(A.CKONGKEP) AS MKONGKEP --매입금액                                               ");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                                                                      ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                                                                                ");
            strSql.AppendLine("     AND A.KERATYPE = '매입'                                                                                    ");
            strSql.AppendLine("     AND A.J_DATE = '"+sDate+"'                                                                                ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '슈레더'                                                                                 ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '인센티브'                                                                               ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                                                                       ");
            strSql.AppendLine("   UNION ALL                                                                                                    ");
            strSql.AppendLine("  /*스크랩 월계*/                                                                                               ");
            strSql.AppendLine("  SELECT '스크랩 월계' AS GB --구분                                                                             ");
            strSql.AppendLine("       , SUM(A.DANJUNG) AS DANJUNG --인수량                                                                     ");
            strSql.AppendLine("       , NULL AS FIRST_CHAGAM--검수감량                                                                         ");
            strSql.AppendLine("       , SUM(A.HALIN) AS HALIN --조정후감량                                                                     ");
            strSql.AppendLine("       , SUM(A.IKONGKEP) / SUM(A.DANJUNG) AS GIDAN --지급단가                                                   ");
            strSql.AppendLine("       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --평균단가                                ");
            strSql.AppendLine("       , (-1) * (1 - ((@SUMILMECH2 / @SUMILMECH1) - "+ dAHALIN + ") / ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)))   AS MGIN --마진율");
            strSql.AppendLine("       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --매입단가                                                        ");
            strSql.AppendLine("       ,  ((@SUMWILMECH2 / @SUMWILMECH1) - "+ dAHALIN + ") - (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MGDAN --마진단가                ");
            strSql.AppendLine("       , SUM(A.IKONGKEP) AS IKONGKEP --지급금액                                                                                         ");
            strSql.AppendLine("       , SUM(A.IKONGKEP) + SUM(A.CKONGKEP) AS MKONGKEP --매입금액                                                                       ");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                                                                                              ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                                                                                                        ");
            strSql.AppendLine("     AND A.KERATYPE = '매입'                                                                                                            ");
            strSql.AppendLine("     AND A.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'                                                                                 ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                                                                                               ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '슈레더'                                                                                                         ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '인센티브'                                                                                                       ");
            strSql.AppendLine("   UNION ALL                                                                                                                            ");
            strSql.AppendLine("  /*G/S모재 입고수량*/                                                                                                                  ");
            strSql.AppendLine("  SELECT 'G/S 모재 입고수량' AS GB --구분                                                                                               ");
	        strSql.AppendLine("       , SUM(A.DANJUNG) AS DANJUNG --인수량                                                                                             ");
	        strSql.AppendLine("       , NULL AS FIRST_CHAGAM--검수감량                                                                                                 ");
	        strSql.AppendLine("       , SUM(A.HALIN) AS HALIN --조정후감량                                                                                             ");
	        strSql.AppendLine("       , SUM(A.IKONGKEP) / SUM(A.DANJUNG) AS GIDAN --지급단가                                                                           ");
	        strSql.AppendLine("       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --평균단가                                                        ");
	        strSql.AppendLine("       , (-1) * (1 - ((@SUMILMECH2 / @SUMILMECH1) - "+ dAHALIN + ") / ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)))   AS MGIN --마진율");
	        strSql.AppendLine("       , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --매입단가                                                        ");
	        strSql.AppendLine("       ,  ((@ILMECH2 / @ILMECH1) - "+ dAHALIN + ") - (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MGDAN --마진단가                        ");
            strSql.AppendLine("       , SUM(A.IKONGKEP) AS IKONGKEP --지급금액                                                                                         ");
	        strSql.AppendLine("       , SUM(A.IKONGKEP) + SUM(A.CKONGKEP) AS MKONGKEP --매입금액");
            strSql.AppendLine("    FROM INLIST A, JAJAE B                                       ");
            strSql.AppendLine("   WHERE A.J_SERIAL = B.J_SERIAL                                 ");
            strSql.AppendLine("     AND A.KERATYPE = '매입'                                     ");
            strSql.AppendLine("     AND A.J_DATE = '"+sDate+"'                                 ");
            strSql.AppendLine("     AND A.J_LOTNO <> '4'                                        ");
            strSql.AppendLine("     AND B.DAEGUBUN <> '슈레더'                                  ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철A'                                     ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철B'                                     ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 BL'                                   ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 AH'                                   ");
            strSql.AppendLine("     AND B.GUBUN1 <> '생철 AL'                                   ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량A'                                     ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 AL'                                   ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 A-B'                                  ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량-ABL'                                  ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량 AH'                                   ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량_ABH'                                  ");
            strSql.AppendLine("     AND B.GUBUN1 <> '중량B'                                     ");
            strSql.AppendLine("     AND B.GUBUN1 <> '기계작업철'                                ");
            strSql.AppendLine("     AND B.GUBUN1 <> '인센티브'                                  ");
            strSql.AppendLine("   UNION ALL                                                     ");
            strSql.AppendLine("   /*G/S 모재 월 입고수량*/                                      ");
            strSql.AppendLine("   SELECT 'G/S 모재 입고수량(누계)' AS GB --구분                 ");
            strSql.AppendLine("        , SUM(A.DANJUNG) AS DANJUNG --인수량                     ");
            strSql.AppendLine("        , NULL AS FIRST_CHAGAM--검수감량                         ");
            strSql.AppendLine("        , SUM(A.HALIN) AS HALIN --조정후감량                     ");
            strSql.AppendLine("        , SUM(A.IKONGKEP) / SUM(A.DANJUNG) AS GIDAN --지급단가   ");
            strSql.AppendLine("        , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --평균단가");
            strSql.AppendLine("        , (-1) * (1 - ((@SUMILMECH2 / @SUMILMECH1) - "+ dAHALIN + ") / ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)))   AS MGIN --마진율");
            strSql.AppendLine("        , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MDAN --매입단가                                ");
            strSql.AppendLine("        ,  ((@WLMECH2 / @WLMECH1) - "+ dAHALIN + ") - (SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG) AS MGDAN --마진단가");
            strSql.AppendLine("        , SUM(A.IKONGKEP) AS IKONGKEP --지급금액                  ");
            strSql.AppendLine("        , SUM(A.IKONGKEP) + SUM(A.CKONGKEP) AS MKONGKEP --매입금액");
            strSql.AppendLine("     FROM INLIST A, JAJAE B                                       ");
            strSql.AppendLine("    WHERE A.J_SERIAL = B.J_SERIAL                                 ");
            strSql.AppendLine("      AND A.KERATYPE = '매입'                                     ");
            strSql.AppendLine("      AND A.J_DATE BETWEEN '"+sFromDate+"' AND '"+sDate+"'          ");
            strSql.AppendLine("      AND A.J_LOTNO <> '4'                                        ");
            strSql.AppendLine("      AND B.DAEGUBUN <> '슈레더'                                  ");
            strSql.AppendLine("      AND B.GUBUN1 <> '생철A'                                     ");
            strSql.AppendLine("      AND B.GUBUN1 <> '생철B'                                     ");
            strSql.AppendLine("      AND B.GUBUN1 <> '생철 BL'                                   ");
            strSql.AppendLine("      AND B.GUBUN1 <> '생철 AH'                                   ");
            strSql.AppendLine("      AND B.GUBUN1 <> '생철 AL'                                   ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량A'                                     ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량 AL'                                   ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량 A-B'                                  ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량-ABL'                                  ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량 AH'                                   ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량_ABH'                                  ");
            strSql.AppendLine("      AND B.GUBUN1 <> '중량B'                                     ");
            strSql.AppendLine("      AND B.GUBUN1 <> '기계작업철'                                ");
            strSql.AppendLine("      AND B.GUBUN1 <> '인센티브'                                  ");

            DataTable dt4 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr3.DataSource = dt4;
            #endregion

            #region YK 단가
            //어음할인율 dAHALIN
            //전체
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                  ");
            strSql.AppendLine("     SELECT 6531121044 AS J_ID   ");
            strSql.AppendLine(" )                               ");
            strSql.AppendLine(" SELECT A1.J_ID                  ");
            strSql.AppendLine("      , A2.DANJUNG AS DANJUNG1   ");
            strSql.AppendLine("      , A2.KONGKEP AS KONGKEP1   ");
            strSql.AppendLine("      , A2.DANGA AS DANGA1       ");
            strSql.AppendLine("      , A3.DANJUNG AS DANJUNG2   ");
            strSql.AppendLine("      , A3.KONGKEP AS KONGKEP2   ");
            strSql.AppendLine("      , A3.DANGA AS DANGA2       ");
            strSql.AppendLine("   FROM TEMP1 A1                 ");
            strSql.AppendLine("   LEFT JOIN(                    ");
            strSql.AppendLine("             /*YK 일매출 전체*/  ");
            strSql.AppendLine("             SELECT A.J_ID1      ");
            strSql.AppendLine("                  , SUM(A.DANJUNG) AS DANJUNG                   ");
            strSql.AppendLine("                  , SUM(A.KONGKEP) -SUM(A.CKONGKEP) AS KONGKEP  ");
            strSql.AppendLine("                  , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) -" + dAHALIN + " AS DANGA");
            strSql.AppendLine("               FROM INLIST A                  ");
            strSql.AppendLine("               LEFT JOIN JAJAE B              ");
            strSql.AppendLine("                 ON A.J_SERIAL = B.J_SERIAL   ");
            strSql.AppendLine("              WHERE A.KERATYPE = '매출'       ");
            strSql.AppendLine("                AND A.J_DATE = '"+ sDate + "'   ");
            strSql.AppendLine("                AND A.J_ID1 = '6531121044'    ");
            strSql.AppendLine("                AND B.GUBUN1 <> '인센티브'    ");
            strSql.AppendLine("                AND A.J_LOTNO <> '4'          ");
            strSql.AppendLine("              GROUP BY A.J_ID1 ) A2           ");
            strSql.AppendLine("     ON A1.J_ID = A2.J_ID1                    ");
            strSql.AppendLine("   LEFT JOIN(/*YK 월매출 전체*/               ");
            strSql.AppendLine("             SELECT A.J_ID1                   ");
            strSql.AppendLine("                  , SUM(A.DANJUNG) AS DANJUNG ");
            strSql.AppendLine("                  , SUM(A.KONGKEP) -SUM(A.CKONGKEP) AS KONGKEP");
            strSql.AppendLine("                  , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) -" + dAHALIN + " AS DANGA");
            strSql.AppendLine("               FROM INLIST A, JAJAE B                              ");
            strSql.AppendLine("              WHERE A.J_SERIAL = B.J_SERIAL                        ");
            strSql.AppendLine("                AND A.J_ID1 = '6531121044'                         ");
            strSql.AppendLine("                AND A.KERATYPE = '매출'                            ");
            strSql.AppendLine("                AND A.J_DATE BETWEEN '"+ sFromDate + "'AND '"+ sDate + "'  ");
            strSql.AppendLine("                AND B.GUBUN1 <> '인센티브'                         ");
            strSql.AppendLine("                AND A.J_LOTNO <> '4'                               ");
            strSql.AppendLine("              GROUP BY A.J_ID1 ) A3                                ");
            strSql.AppendLine("     ON A1.J_ID = A3.J_ID1                                         ");

            DataTable dt5 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr4.DataSource = dt5;


            strSql.Clear();
            strSql.AppendLine(" WITH TEMP1 AS(                 ");
            strSql.AppendLine("     SELECT 6531121044 AS J_ID  ");
            strSql.AppendLine(" )                              ");
            strSql.AppendLine(" SELECT A1.J_ID                 ");
            strSql.AppendLine("      , A2.DANJUNG AS DANJUNG1  ");
            strSql.AppendLine("      , A2.KONGKEP AS KONGKEP1  ");
            strSql.AppendLine("      , A2.DANGA AS DANGA1      ");
            strSql.AppendLine("      , A3.DANJUNG AS DANJUNG2  ");
            strSql.AppendLine("      , A3.KONGKEP AS KONGKEP2  ");
            strSql.AppendLine("      , A3.DANGA AS DANGA2      ");
            strSql.AppendLine("   FROM TEMP1 A1                ");
            strSql.AppendLine("   LEFT JOIN(                   ");
            strSql.AppendLine("             /*yk 일매출 G/S */ ");
            strSql.AppendLine("             SELECT A.J_ID1     ");
            strSql.AppendLine("                  , SUM(A.DANJUNG) AS DANJUNG");
            strSql.AppendLine("                  , SUM(A.KONGKEP) -SUM(A.CKONGKEP) AS KONGKEP");
            strSql.AppendLine("                  , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) -" + dAHALIN + " AS DANGA");
            strSql.AppendLine("               FROM INLIST A, JAJAE B        ");
            strSql.AppendLine("              WHERE A.J_SERIAL = B.J_SERIAL  ");
            strSql.AppendLine("                AND A.KERATYPE = '매출'      ");
            strSql.AppendLine("                AND A.J_DATE = '"+ sDate + "'  ");
            strSql.AppendLine("                AND A.J_ID1 = '6531121044'   ");
            strSql.AppendLine("                AND A.J_LOTNO <> '4'         ");
            strSql.AppendLine("                AND B.DAEGUBUN <> '슈레더'   ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철A'      ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철B'      ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철 AH'    ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철 AL'    ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량A'      ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량 AL'    ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량 A-B'   ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량-ABL'   ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량 AH'    ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량_ABH'   ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량B'      ");
            strSql.AppendLine("                AND B.GUBUN1 <> '인센티브'   ");
            strSql.AppendLine("              GROUP BY A.J_ID1 ) A2          ");
            strSql.AppendLine("     ON A1.J_ID = A2.J_ID1                   ");
            strSql.AppendLine("   LEFT JOIN(/*YK 월 매출 G/S*/              ");
            strSql.AppendLine("             SELECT A.J_ID1                  ");
            strSql.AppendLine("                  , SUM(A.DANJUNG) AS DANJUNG");
            strSql.AppendLine("                  , SUM(A.KONGKEP) -SUM(A.CKONGKEP) AS KONGKEP");
            strSql.AppendLine("                  , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) -" + dAHALIN + " AS DANGA");
            strSql.AppendLine("               FROM INLIST A, JAJAE B                             ");
            strSql.AppendLine("              WHERE A.J_SERIAL = B.J_SERIAL                       ");
            strSql.AppendLine("                AND A.KERATYPE = '매출'                           ");
            strSql.AppendLine("                AND A.J_DATE BETWEEN '"+ sFromDate + "'AND '"+ sDate + "' ");
            strSql.AppendLine("                AND A.J_ID1 = '6531121044'                        ");
            strSql.AppendLine("                AND A.J_LOTNO <> '4'                              ");
            strSql.AppendLine("                AND B.DAEGUBUN <> '슈레더'                        ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철A'                           ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철B'                           ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철 AH'                         ");
            strSql.AppendLine("                AND B.GUBUN1 <> '생철 AL'                         ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량A'                           ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량 AL'                         ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량 A-B'                        ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량-ABL'                        ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량 AH'                         ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량_ABH'                        ");
            strSql.AppendLine("                AND B.GUBUN1 <> '중량B'                           ");
            strSql.AppendLine("                AND B.GUBUN1 <> '인센티브'                        ");
            strSql.AppendLine("              GROUP BY A.J_ID1 ) A3                               ");
            strSql.AppendLine("     ON A1.J_ID = A3.J_ID1                                        ");

            DataTable dt6 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr5.DataSource = dt6;
            #endregion

            //폐압
            #region 일일 야드 매입 단가 내역서
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                          ");
            strSql.AppendLine("     /*일일야드 매입 단가 내역서(폐압)*/ ");
            strSql.AppendLine("     SELECT A1.J_BNUM--차번              ");
            strSql.AppendLine("          , B2.DEALER_NM--거래처         ");
            strSql.AppendLine("          , A1.GUBUN1--등급              ");
            strSql.AppendLine("          , B3.DANJUNG--인수량           ");
            strSql.AppendLine("          , ISNULL(C1.CHAGAM, ISNULL(A2.CHAGAM, CASE WHEN A1.KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END)) AS FIRST_CHAGAM--검수감량");
            strSql.AppendLine("          , B3.HALIN--조정감량                                         ");
            strSql.AppendLine("          , B3.MIDANGA--테이블단가                                     ");
            strSql.AppendLine("          , B3.DANGA--지급단가                                         ");
            strSql.AppendLine("          , B2.J_REGION--지역                                          ");
            strSql.AppendLine("          , B3.CKONGKEP / B3.DANJUNG AS CKONGKEP--운반비(운반비단가)   ");
            strSql.AppendLine("          , NULL AS CKONGKEP2--운반단가                                ");
            strSql.AppendLine("          , (B3.IKONGKEP + B3.CKONGKEP) / B3.DANJUNG AS MDAN--매입단가 ");
            strSql.AppendLine("          , B3.MIDANGA - ((B3.IKONGKEP + B3.CKONGKEP) / B3.DANJUNG) AS CHA--차액");
            strSql.AppendLine("          , A1.GUMSUBIGO--검수판정                                              ");
            strSql.AppendLine("          , B3.CHRG_NM--담당자                                                  ");
            strSql.AppendLine("          , B3.IKONGKEP-- 지급액                                                ");
            strSql.AppendLine("          , B3.IKONGKEP + B3.CKONGKEP AS MAIPKEP-- 매입금액                     ");
            strSql.AppendLine("       FROM MESURING A1                                                         ");
            strSql.AppendLine("       LEFT JOIN MESURING_SEQ A2                                                ");
            strSql.AppendLine("         ON A1.JUNPYOID = A2.JUNPYOID                                           ");
            strSql.AppendLine("        AND A2.CHG_SEQ = (SELECT MIN(X1.CHG_SEQ) FROM MESURING_SEQ X1 WHERE X1.JUNPYOID = A1.JUNPYOID AND X1.CHAGAM > 0)");
            strSql.AppendLine("       LEFT JOIN INLIST B3                              ");
            strSql.AppendLine("         ON A1.IPCHULGO_MAIPID = B3.J_ID                ");
            strSql.AppendLine("       LEFT JOIN HR_EMP_BASIS B1                        ");
            strSql.AppendLine("         ON A1.GUMSU_SERIAL = B1.EMP_ID                 ");
            strSql.AppendLine("       LEFT JOIN ACC_DEALER_CD B2                       ");
            strSql.AppendLine("         ON A1.MAIPCHERID = B2.DEALER_CD                ");
            strSql.AppendLine("       LEFT JOIN MESURE_ISPT_INFO C1                    ");
            strSql.AppendLine("         ON A1.JUNPYOID = C1.JUNPYOID                   ");
            strSql.AppendLine("       LEFT JOIN(SELECT X1.USRCD, X2.EMP_NM             ");
            strSql.AppendLine("                     FROM ZUSRLST X1                    ");
            strSql.AppendLine("                     LEFT JOIN HR_EMP_BASIS X2          ");
            strSql.AppendLine("                       ON X1.INSANO = X2.EMP_ID) C2     ");
            strSql.AppendLine("        ON C1.ENT_ID = C2.USRCD                         ");
            strSql.AppendLine("       LEFT JOIN JAJAE D1 ON A1.J_SERIAL = D1.J_SERIAL  ");
            strSql.AppendLine("      WHERE B3.KERATYPE = '매입'                        ");
            strSql.AppendLine("        AND B3.J_DATE = '"+sDate+"'                    ");
            strSql.AppendLine("        AND B3.J_LOTNO <> '4'                           ");
            strSql.AppendLine("        AND D1.DAEGUBUN = '슈레더'                      ");
            strSql.AppendLine("        AND D1.GUBUN1 <> '인센티브'                     ");
            strSql.AppendLine(" )                                                      ");
            strSql.AppendLine(" SELECT 1 AS NUM                                        ");
            strSql.AppendLine("      , J_BNUM --차번                                   ");
            strSql.AppendLine("      , DEALER_NM--거래처                               ");
            strSql.AppendLine("      , GUBUN1--등급                                    ");
            strSql.AppendLine("      , DANJUNG--인수량                                 ");
            strSql.AppendLine("      , FIRST_CHAGAM--검수감량                          ");
            strSql.AppendLine("      , HALIN--조정감량                                 ");
            strSql.AppendLine("      , CAST(MIDANGA AS VARCHAR) AS MIDANGA --테이블단가");
            strSql.AppendLine("      , DANGA--지급단가                                 ");
            strSql.AppendLine("      , J_REGION--지역                                  ");
            strSql.AppendLine("      , CKONGKEP--운반비(운반비단가)                    ");
            strSql.AppendLine("      , CKONGKEP2--운반단가                             ");
            strSql.AppendLine("      , MDAN--매입단가                                  ");
            strSql.AppendLine("      , CHA--차액                                       ");
            strSql.AppendLine("      , GUMSUBIGO--검수판정                             ");
            strSql.AppendLine("      , CHRG_NM--담당자                                 ");
            strSql.AppendLine("      , IKONGKEP-- 지급액                               ");
            strSql.AppendLine("      , MAIPKEP-- 매입금액                              ");
            strSql.AppendLine("  FROM TEMP1                                            ");
            strSql.AppendLine(" UNION ALL                                              ");
            strSql.AppendLine(" SELECT 2 AS NUM                                        ");
            strSql.AppendLine("     , NULL                                             ");
            strSql.AppendLine("     , '합  계'                           ");
            strSql.AppendLine("      , NULL                                            ");
            strSql.AppendLine("      , SUM(DANJUNG)                                    ");
            strSql.AppendLine("      , SUM(FIRST_CHAGAM)                               ");
            strSql.AppendLine("      , SUM(HALIN)                                      ");
            strSql.AppendLine("      , '지급단가'                                      ");
            strSql.AppendLine("      , SUM(IKONGKEP) / SUM(DANJUNG)                    ");
            strSql.AppendLine("      , NULL                                            ");
            strSql.AppendLine("      , NULL                                            ");
            strSql.AppendLine("      , '매입단가'                                      ");
            strSql.AppendLine("      , SUM(MAIPKEP) / SUM(DANJUNG)                     ");
            strSql.AppendLine("      , NULL                                            ");
            strSql.AppendLine("      , NULL                                            ");
            strSql.AppendLine("      , NULL                                            ");
            strSql.AppendLine("      , SUM(IKONGKEP)                                   ");
            strSql.AppendLine("      , SUM(MAIPKEP)                                    ");
            strSql.AppendLine("   FROM TEMP1                                           ");

            DataTable dt7 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr6.DataSource = dt7;

            #endregion

            //직납
            #region 직납 매입
            strSql.Clear();
            /*YK직납 매입*/
            strSql.AppendLine("SET ANSI_WARNINGS OFF");
            strSql.AppendLine("SET ARITHIGNORE ON   ");
            strSql.AppendLine("SET ARITHABORT OFF;  ");

            strSql.AppendLine(" WITH TEMP1 AS(                               ");
            strSql.AppendLine("     SELECT Z2.ROWNUM                         ");
            strSql.AppendLine("          , Z1.JUNPYOID                       ");
            strSql.AppendLine("          , Z1.J_BNUM--차량번호               ");
            strSql.AppendLine("          , Z2.PURC_DEALER_NM--거래처         ");
            strSql.AppendLine("          , Z1.GUBUN1--등급                   ");
            strSql.AppendLine("          , Z1.SALEUNITPRICE--어음(지급)단가  ");
            strSql.AppendLine("          , Z2.J_REGION--상차지역             ");
            strSql.AppendLine("          , '' as JERDAN --절단비감가         ");
            strSql.AppendLine("          , Z2.PURCHUNITPRICE--매입단가       ");
            strSql.AppendLine("          , Z1.CKONGKEP AS CARRYCOST--운반비  ");
            strSql.AppendLine("          , (Z1.OWEIGHT * Z2.PURCHUNITPRICE + Z1.CKONGKEP) / Z1.OWEIGHT AS ARRIVEUNITPRICE--도착도단가");
            strSql.AppendLine("          , Z1.OWEIGHT--실중량 ");
            strSql.AppendLine("          , Z2.HALIN--감량     ");
            strSql.AppendLine("          , ISNULL(Z1.SALEUNITPRICE, 0) - ((ISNULL(Z1.OWEIGHT, 0) * ISNULL(Z2.PURCHUNITPRICE, 0)) + ISNULL(Z1.CKONGKEP, 0)) / Z1.OWEIGHT AS DIFFUNITPRICE--차액");
            strSql.AppendLine("          , '' as AVGDAN --평균단가                       ");
            strSql.AppendLine("          , Z1.SALEPRICE--매출액                          ");
            strSql.AppendLine("          , Z2.PURC_AMT--매입금액                         ");
            strSql.AppendLine("          , Z2.J_BOOKING-- 매출처명인듯                       ");
            strSql.AppendLine("      FROM (SELECT A2.JUNPYOID                            ");
            strSql.AppendLine("                 , A1.J_DATE                              ");
            strSql.AppendLine("                 , C1.GUBUN1                              ");
            strSql.AppendLine("                 , A2.J_BNUM                              ");
            strSql.AppendLine("                 , B1.DEALER_NM AS SALE_DEALER_NM         ");
            strSql.AppendLine("                 , A1.DANJUNG AS OWEIGHT                  ");
            strSql.AppendLine("                 , A1.DANGA AS SALEUNITPRICE              ");
            strSql.AppendLine("                 , A1.KONGKEP AS SALEPRICE                ");
            strSql.AppendLine("                 , CASE WHEN A1.CKONGKEP = 0 THEN NULL    ");
            strSql.AppendLine("                        ELSE A1.CKONGKEP END AS CKONGKEP  ");
            strSql.AppendLine("              FROM INLIST A1                              ");
            strSql.AppendLine("              LEFT JOIN MESURING A2                       ");
            strSql.AppendLine("                ON A1.J_ID = A2.IPCHULGO_MACHULID         ");
            strSql.AppendLine("              LEFT JOIN(SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID");
            strSql.AppendLine("                          FROM ACC_DEALER_CD X1                        ");
            strSql.AppendLine("                          LEFT JOIN HR_EMP_BASIS X2                    ");
            strSql.AppendLine("                              ON X1.CHRG_ID = X2.EMP_ID) B1            ");
            strSql.AppendLine("                ON A1.J_ID1 = B1.DEALER_CD                             ");
            strSql.AppendLine("              LEFT JOIN JAJAE C1                                       ");
            strSql.AppendLine("                ON A1.J_SERIAL = C1.J_SERIAL                           ");
            strSql.AppendLine("             WHERE A2.KERATYPE = '직송'                                ");
            strSql.AppendLine("               AND A1.KERATYPE = '매출'                                ");
            strSql.AppendLine("               AND A1.J_DATE BETWEEN '"+sDate+"' AND '"+sDate+"')Z1  ");
            strSql.AppendLine("      LEFT JOIN(SELECT ROW_NUMBER() OVER(ORDER BY A1.J_DATE, REPLACE(B1.DEALER_NM, '(주)', ''), A1.J_SERIAL) AS ROWNUM");
            strSql.AppendLine("                     , A2.JUNPYOID                                                                                    ");
            strSql.AppendLine("                     , A1.J_DATE                                                                                      ");
            strSql.AppendLine("                     , B1.DEALER_NM AS PURC_DEALER_NM                                                                 ");
            strSql.AppendLine("                     , A1.DANGA AS PURCHUNITPRICE                                                                     ");
            strSql.AppendLine("                     , A1.IKONGKEP AS PURC_AMT                                                                        ");
            strSql.AppendLine("                     , B1.J_REGION                                                                                    ");
            strSql.AppendLine("                     , A1.HALIN                                                                                       ");
            strSql.AppendLine("                     , A1.J_BOOKING                                                                                   ");
            strSql.AppendLine("                  FROM INLIST A1                                                                                      ");
            strSql.AppendLine("                  LEFT JOIN MESURING A2                                                                               ");
            strSql.AppendLine("                    ON A1.J_ID = A2.IPCHULGO_MAIPID                                                                   ");
            strSql.AppendLine("                  LEFT JOIN(SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID, X1.J_REGION                     ");
            strSql.AppendLine("                              FROM ACC_DEALER_CD X1                                                                   ");
            strSql.AppendLine("                              LEFT JOIN HR_EMP_BASIS X2                                                               ");
            strSql.AppendLine("                                  ON X1.CHRG_ID = X2.EMP_ID) B1                                                       ");
            strSql.AppendLine("                    ON A1.J_ID1 = B1.DEALER_CD                                                                        ");
            strSql.AppendLine("                  LEFT JOIN JAJAE C1                                                                                  ");
            strSql.AppendLine("                    ON A1.J_SERIAL = C1.J_SERIAL                                                                      ");
            strSql.AppendLine("                 WHERE A2.KERATYPE = '직송'                                                                           ");
            strSql.AppendLine("                   AND A1.KERATYPE = '매입'                                                                           ");
            strSql.AppendLine("                   AND A1.J_DATE BETWEEN '"+sDate+"' AND '"+sDate+"')Z2                                             ");
            strSql.AppendLine("        ON Z1.JUNPYOID = Z2.JUNPYOID                                                                                  ");
            strSql.AppendLine(" )                                                                                                                    ");
            strSql.AppendLine("                                                                                                                      ");
            strSql.AppendLine(" SELECT ROWNUM                                                                                                        ");
            strSql.AppendLine("      , JUNPYOID                                                                                                      ");
            strSql.AppendLine("      , J_BNUM --차량번호                                                                                             ");
            strSql.AppendLine("      , PURC_DEALER_NM--거래처                                                                                        ");
            strSql.AppendLine("      , GUBUN1--등급                                                                                                  ");
            strSql.AppendLine("      , SALEUNITPRICE--어음(지급)단가                                                                                 ");
            strSql.AppendLine("      , J_REGION--상차지역                                                                                            ");
            strSql.AppendLine("      , JERDAN-- 절단비감가                                                                                           ");
            strSql.AppendLine("      , PURCHUNITPRICE--매입단가                                                                                      ");
            strSql.AppendLine("      , CARRYCOST--운반비                                                                                             ");
            strSql.AppendLine("      , ARRIVEUNITPRICE--도착도단가                                                                                   ");
            strSql.AppendLine("      , OWEIGHT--실중량                                                                                               ");
            strSql.AppendLine("      , HALIN--감량                                                                                                   ");
            strSql.AppendLine("      , DIFFUNITPRICE--차액                                                                                           ");
            strSql.AppendLine("      , AVGDAN--평균단가                                                                                              ");
            strSql.AppendLine("      , SALEPRICE--매출액                                                                                             ");
            strSql.AppendLine("      , PURC_AMT--매입금액                                                                                            ");
            strSql.AppendLine("   FROM TEMP1                                                                                                         ");
            strSql.AppendLine("  UNION ALL                                                                                                           ");
            strSql.AppendLine(" SELECT 999999997 AS NUM                                                                                              ");
            strSql.AppendLine("      , NULL                                                                                                          ");
            strSql.AppendLine("      , NULL --차량번호                                                                                               ");
            strSql.AppendLine("      , 'YK 직납 매입 소계'--거래처                                                                                   ");
            strSql.AppendLine("      , NULL--등급                                                                                                    ");
            strSql.AppendLine("      , NULL--어음(지급)단가                                                                                          ");
            strSql.AppendLine("      , NULL--상차지역                                                                                                ");
            strSql.AppendLine("      , NULL-- 절단비감가                                                                                             ");
            strSql.AppendLine("      , NULL--매입단가                                                                                                ");
            strSql.AppendLine("      , NULL--운반비                                                                                                  ");
            strSql.AppendLine("      , NULL--도착도단가                                                                                              ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING = '와이케이스틸 주식회사' THEN OWEIGHT END)--실중량                                   ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING = '와이케이스틸 주식회사' THEN HALIN END)--감량                                       ");
            strSql.AppendLine("      , SUM(DIFFUNITPRICE)--차액                                                                                      ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING = '와이케이스틸 주식회사' THEN PURC_AMT END) / SUM(CASE WHEN J_BOOKING = '와이케이스틸 주식회사' THEN OWEIGHT END)--평균단가");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING = '와이케이스틸 주식회사' THEN SALEPRICE END)--매출액    ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING = '와이케이스틸 주식회사' THEN PURC_AMT END)--매입금액   ");
            strSql.AppendLine("  FROM TEMP1                                                                             ");
            strSql.AppendLine(" UNION ALL                                                                               ");
            strSql.AppendLine(" SELECT 999999998 AS NUM                                                                 ");
            strSql.AppendLine("      , NULL                                                                             ");
            strSql.AppendLine("      , NULL --차량번호                                                                  ");
            strSql.AppendLine("      , 'YK 직납 매입 소계(기타)'--거래처                                                ");
            strSql.AppendLine("      , NULL--등급                                                                       ");
            strSql.AppendLine("      , NULL--어음(지급)단가                                                             ");
            strSql.AppendLine("      , NULL--상차지역                                                                   ");
            strSql.AppendLine("      , NULL-- 절단비감가                                                                ");
            strSql.AppendLine("      , NULL--매입단가                                                                   ");
            strSql.AppendLine("      , NULL--운반비                                                                     ");
            strSql.AppendLine("      , NULL--도착도단가                                                                 ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING != '와이케이스틸 주식회사' THEN OWEIGHT END)--실중량     ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING != '와이케이스틸 주식회사' THEN HALIN END)--감량         ");
            strSql.AppendLine("      , NULL--차액                                                                       ");
            strSql.AppendLine("      , NULL--평균단가                                                                   ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING != '와이케이스틸 주식회사' THEN SALEPRICE END)--매출액   ");
            strSql.AppendLine("      , SUM(CASE WHEN J_BOOKING != '와이케이스틸 주식회사' THEN PURC_AMT END)--매입금액  ");
            strSql.AppendLine("   FROM TEMP1                                                                            ");
            strSql.AppendLine("  UNION ALL                                                                              ");
            strSql.AppendLine(" SELECT 999999999 AS NUM                                                                 ");
            strSql.AppendLine("      , NULL                                                                             ");
            strSql.AppendLine("      , NULL --차량번호                                                                  ");
            strSql.AppendLine("      , 'YK 직납 매입 일계'--거래처                                                      ");
            strSql.AppendLine("      , NULL--등급                                                                       ");
            strSql.AppendLine("      , NULL--어음(지급)단가                                                             ");
            strSql.AppendLine("      , NULL--상차지역                                                                   ");
            strSql.AppendLine("      , NULL-- 절단비감가                                                                ");
            strSql.AppendLine("      , NULL--매입단가                                                                   ");
            strSql.AppendLine("      , NULL--운반비                                                                     ");
            strSql.AppendLine("      , NULL--도착도단가                                                                 ");
            strSql.AppendLine("      , SUM(OWEIGHT)--실중량                                                             ");
            strSql.AppendLine("      , SUM(HALIN)--감량                                                                 ");
            strSql.AppendLine("      , NULL--차액                                                                       ");
            strSql.AppendLine("      , NULL--평균단가                                                                   ");
            strSql.AppendLine("      , SUM(SALEPRICE)--매출액                                                           ");
            strSql.AppendLine("      , SUM(PURC_AMT)--매입금액                                                          ");
            strSql.AppendLine("   FROM TEMP1                                                                            ");

            DataTable dt8 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr7.DataSource = dt8;
            #endregion

            #region 직납 YK 매출
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" SELECT A.JUNPYOID                                   ");
            strSql.AppendLine("      , A.J_DATE");
            strSql.AppendLine("      , C.J_BNUM--차량번호                           ");
            strSql.AppendLine("      , B.GUBUN1--등급                               ");
            strSql.AppendLine("      , A.KONGKEP / A.DANJUNG AS GIDAN--지급단가     ");
            strSql.AppendLine("      , A.DANJUNG + A.HALIN AS DANJUNG--중량(감량전) ");
            strSql.AppendLine("      , A.HALIN--감량                                ");
            strSql.AppendLine("      , A.KONGKEP--매출금액                          ");
            strSql.AppendLine("      , A.CKONGKEP--운반비                           ");
            strSql.AppendLine("      , A.KONGKEP - A.CKONGKEP AS SUNKEP--순매출액   ");
            strSql.AppendLine("      , (A.KONGKEP - A.CKONGKEP) / (A.DANJUNG + A.HALIN) AS SUNDAN --순매출단가");
            strSql.AppendLine("      , D.GAMSU--검수                                                          ");
            strSql.AppendLine("      , D.JANGBE--장비                                                         ");
            strSql.AppendLine("      , D.NWEIGHT--납품후공차                                                  ");
            strSql.AppendLine("      , D.CWEIGHT--공차                                                        ");
            strSql.AppendLine("      , D.CHAWEIGHT--차이중량                                                  ");
            strSql.AppendLine("      , A.DANJUNG AS GMDANJUNG--중량(감량후)                                   ");
            strSql.AppendLine("   FROM INLIST A                                                               ");
            strSql.AppendLine("   LEFT JOIN JAJAE B                                                           ");
            strSql.AppendLine("     ON A.J_SERIAL = B.J_SERIAL                                                ");
            strSql.AppendLine("   LEFT JOIN MESURING C                                                        ");
            strSql.AppendLine("     ON A.J_ID = C.IPCHULGO_MACHULID                                           ");
            strSql.AppendLine("   LEFT JOIN JKNAP_W D                                                         ");
            strSql.AppendLine("     ON A.JUNPYOID = D.JUNPYOID                                                ");
            strSql.AppendLine("  WHERE A.KERATYPE = '매출'                                                    ");
            strSql.AppendLine("    AND A.J_DATE = '"+ sDate + "'                                                ");
            strSql.AppendLine("    AND A.J_ID1 = '6531121044'                                                 ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브'                                                 ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                       ");

            DataTable dt9 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr8.DataSource = dt9;
            #endregion

            #region ASR
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" SELECT A2.J_BNUM                   ");
            strSql.AppendLine("      , B1.DEALER_NM                ");
            strSql.AppendLine("      , A3.GUBUN1                   ");
            strSql.AppendLine("      , A1.DANGA                    ");
            strSql.AppendLine("      , A1.WEIGHT                   ");
            strSql.AppendLine("      , NULL AS GAMW                ");
            strSql.AppendLine("      , A1.PRICE                    ");
            strSql.AppendLine("      , NULL AS BIGO                ");
            strSql.AppendLine("   FROM J_MAGAM A1                  ");
            strSql.AppendLine("   LEFT JOIN MESURING A2            ");
            strSql.AppendLine("     ON A1.M_ID = A2.JUNPYOID       ");
            strSql.AppendLine("   LEFT JOIN JAJAE A3               ");
            strSql.AppendLine("     ON A1.J_SERIAL = A3.J_SERIAL   ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B1       ");
            strSql.AppendLine("     ON A2.J_ASSIGNID = B1.DEALER_CD");
            strSql.AppendLine("  WHERE A1.J_SERIAL = 2025163--ASR  ");
            strSql.AppendLine("   AND J_DATE = '"+ sDate + "'        ");

            DataTable dt10 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr9.DataSource = dt10;
            #endregion

            //매출일계 매입+매출 갯수
            double dMaip = 0;
            double dMachul = 0;

            if(dt8 != null)
            {
                dMaip = dt8.Rows.Count - 3;
            }

            if(dt9 != null)
            {
                dMachul = dt9.Rows.Count;
            }

            TxtCarCnt.EditValue = (dMaip + dMachul).ToString() + " 대";

            SplashScreenManager.CloseForm();
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));

            FileInfo_1 fileInfo = new FileInfo_1("6");

            string[] sPath = fileInfo.CheckFileInfo();

            if (sPath != null)
            {
                SetWeeklyReportForm(sPath[0], sPath[1]);
            }

            SplashScreenManager.CloseForm();
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

                string sDate = DateEditDT.EditValue?.ToString().Substring(0, 10);

                ExcelApp = new Excel.Application();

                //메세지 없애기
                ExcelApp.DisplayAlerts = false;

                wb = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                //스크랩
                ws = wb.Worksheets.get_Item(1);

                //시작 줄
                int iStartRow = 0;

                #region 통계
                DataTable dt1 = (DataTable)GridRetr1.DataSource;

                if (dt1 != null)
                {
                    //일자부분 세팅
                    ws.Range["A2"].Value = sDate;

                    //목표 통계
                    //야드매입
                    string sTxtOgs1 = TxtOgs1.EditValue?.ToString();
                    string sTxtOgs2 = TxtOgs2.EditValue?.ToString();
                    string sTxtOgs3 = TxtOgs3.EditValue?.ToString();

                    ws.Range["B3"].Value = sTxtOgs1;
                    ws.Range["B4"].Value = sTxtOgs2;
                    ws.Range["B5"].Value = sTxtOgs3;

                    //폐압
                    string sTxtOW1 = TxtOW1.EditValue.ToString();
                    string sTxtOW2 = TxtOW2.EditValue.ToString();
                    string sTxtOW3 = TxtOW3.EditValue.ToString();

                    ws.Range["E3"].Value = sTxtOW1;
                    ws.Range["E4"].Value = sTxtOW2;
                    ws.Range["E5"].Value = sTxtOW3;

                    //야드매출                 
                    string sTxtYd1 = TxtYd1.EditValue.ToString();
                    string sTxtYd2 = TxtYd2.EditValue.ToString();
                    string sTxtYd3 = TxtYd3.EditValue.ToString();

                    ws.Range["H3"].Value = sTxtYd1;
                    ws.Range["H4"].Value = sTxtYd2;
                    ws.Range["H5"].Value = sTxtYd3;

                    //직납매출                
                    string sTxtYk1 = TxtYk1.EditValue.ToString();
                    string sTxtYk2 = TxtYk2.EditValue.ToString();
                    string sTxtYk3 = TxtYk3.EditValue.ToString();

                    ws.Range["K3"].Value = sTxtYk1;
                    ws.Range["K4"].Value = sTxtYk2;
                    ws.Range["K5"].Value = sTxtYk3;


                    //담당자별 집계
                    iStartRow = 7;

                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        string sNAME = dt1.Rows[i]["NAME"]?.ToString();
                        string sSCDAY = dt1.Rows[i]["SCDAY"]?.ToString();
                        string sPEDAY = dt1.Rows[i]["PEDAY"]?.ToString();
                        string sSCMON = dt1.Rows[i]["SCMON"]?.ToString();
                        string sPEMON = dt1.Rows[i]["PEMON"]?.ToString();
                        string sSCRAP = dt1.Rows[i]["SCRAP"]?.ToString();
                        string sPEAP = dt1.Rows[i]["PEAP"]?.ToString();
                        string sDALPER_S = dt1.Rows[i]["DALPER_S"]?.ToString();
                        string sDALPER_P = dt1.Rows[i]["DALPER_P"]?.ToString();
                        string sATSIL_S = dt1.Rows[i]["ATSIL_S"]?.ToString();
                        string sATSIL_P = dt1.Rows[i]["ATSIL_P"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["A" + iApplyRowIdx].Value = sNAME;
                        ws.Range["B" + iApplyRowIdx].Value = sSCDAY;
                        ws.Range["C" + iApplyRowIdx].Value = sPEDAY;
                        ws.Range["D" + iApplyRowIdx].Value = sSCMON;
                        ws.Range["E" + iApplyRowIdx].Value = sPEMON;
                        ws.Range["F" + iApplyRowIdx].Value = sSCRAP;
                        ws.Range["G" + iApplyRowIdx].Value = sPEAP;
                        ws.Range["H" + iApplyRowIdx].Value = sDALPER_S;
                        ws.Range["I" + iApplyRowIdx].Value = sDALPER_P;
                        ws.Range["J" + iApplyRowIdx].Value = sATSIL_S;
                        ws.Range["K" + iApplyRowIdx].Value = sATSIL_P;
                    }
                }
                #endregion

                #region 스크랩 매입
                DataTable dt2 = (DataTable)GridRetr2.DataSource;

                iStartRow = 19;
                //매입
                if (dt2 != null)
                {
                    int iApplyRowIdx = 0;
                    for (int i=0; i<dt2.Rows.Count; i++)
                    {
                        string sJ_DATE       = dt2.Rows[i]["J_DATE"]?.ToString();
                        string sJ_BNUM       = dt2.Rows[i]["J_BNUM"]?.ToString();
                        string sDEALER_NM    = dt2.Rows[i]["DEALER_NM"]?.ToString();
                        string sGUBUN1       = dt2.Rows[i]["GUBUN1"]?.ToString();
                        string sDANJUNG      = dt2.Rows[i]["DANJUNG"]?.ToString();
                        string sFIRST_CHAGAM = dt2.Rows[i]["FIRST_CHAGAM"]?.ToString();
                        string sHALIN        = dt2.Rows[i]["HALIN"]?.ToString();
                        string sMIDANGA      = dt2.Rows[i]["MIDANGA"]?.ToString();
                        string sDANGA        = dt2.Rows[i]["DANGA"]?.ToString();
                        string sJ_REGION     = dt2.Rows[i]["J_REGION"]?.ToString();
                        string sCKONGKEP     = dt2.Rows[i]["CKONGKEP"]?.ToString();
                        string sMIPDAN       = dt2.Rows[i]["MIPDAN"]?.ToString();
                        string sCHA          = dt2.Rows[i]["CHA"]?.ToString();
                        string sGUMSUBIGO    = dt2.Rows[i]["GUMSUBIGO"]?.ToString();
                        string sCHRG_NM      = dt2.Rows[i]["CHRG_NM"]?.ToString();
                        string sCHRDAN       = dt2.Rows[i]["CHRDAN"]?.ToString();
                        string sDTEXT        = dt2.Rows[i]["DTEXT"]?.ToString();
                        string sIKONGKEP     = dt2.Rows[i]["IKONGKEP"]?.ToString();
                        string sMAIPK = dt2.Rows[i]["MAIPK"]?.ToString();

                        iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["A" + iApplyRowIdx].Value = i + 1;

                        //차번 텍스트형식으로 변경
                        Excel.Range range1 = ws.Range["B" + iApplyRowIdx];
                        range1.NumberFormat = "@";

                        ws.Range["B" + iApplyRowIdx].Value = sJ_BNUM;
                        //셀병합
                        Excel.Range range2 = ws.Range["C" + iApplyRowIdx, "D" + iApplyRowIdx];
                        range2.Merge();

                        ws.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                        ws.Range["E" + iApplyRowIdx].Value = sGUBUN1;
                        ws.Range["F" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["G" + iApplyRowIdx].Value = sFIRST_CHAGAM;
                        ws.Range["H" + iApplyRowIdx].Value = sHALIN;
                        ws.Range["I" + iApplyRowIdx].Value = sMIDANGA;
                        ws.Range["J" + iApplyRowIdx].Value = sDANGA;
                        ws.Range["K" + iApplyRowIdx].Value = sJ_REGION;
                        ws.Range["L" + iApplyRowIdx].Value = sCKONGKEP;
                        ws.Range["M" + iApplyRowIdx].Value = sMIPDAN;
                        ws.Range["N" + iApplyRowIdx].Value = sCHA;
                        ws.Range["O" + iApplyRowIdx].Value = sGUMSUBIGO;
                        ws.Range["P" + iApplyRowIdx].Value = sCHRG_NM;
                        ws.Range["Q" + iApplyRowIdx].Value = sCHRDAN;
                        ws.Range["R" + iApplyRowIdx].Value = sDTEXT;
                        ws.Range["T" + iApplyRowIdx].Value = sIKONGKEP;
                        ws.Range["U" + iApplyRowIdx].Value = sMAIPK;

                        //셀 모든 테두리
                        Excel.Range range3 = ws.Range["A" + iApplyRowIdx+":"+"U"+ iApplyRowIdx];
                        range3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                    iStartRow = iApplyRowIdx;

                    if (iStartRow < 19)
                    {
                        iStartRow = 19;
                    }
                }

                //스크랩 집계
                DataTable dt3 = (DataTable)GridRetr3.DataSource;

                //어음할인율
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" IF EXISTS(SELECT AHALIN FROM JKNAP_W WHERE JUNPYOID = 0 AND J_DATE = '"+sDate+"') ");
                strSql.AppendLine("    BEGIN                                                                          ");
                strSql.AppendLine("          SELECT AHALIN                                                            ");
                strSql.AppendLine("            FROM JKNAP_W                                                           ");
                strSql.AppendLine("           WHERE JUNPYOID = 0                                                      ");
                strSql.AppendLine("             AND J_DATE = '"+sDate+"'                                              ");
                strSql.AppendLine("      END                                                                          ");
                strSql.AppendLine(" ELSE                                                                              ");
                strSql.AppendLine("     BEGIN                                                                         ");
                strSql.AppendLine("           SELECT AHALIN                                                           ");
                strSql.AppendLine("             FROM JKNAP_W                                                          ");
                strSql.AppendLine("            WHERE JUNPYOID = 0                                                     ");
                strSql.AppendLine("              AND J_DATE = (SELECT MAX(J_DATE) FROM JKNAP_W WHERE JUNPYOID = 0)    ");
                strSql.AppendLine("       END                                                                         ");

                DataTable dtAHALIN = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt3 != null)
                {
                    int iApplyRowIdx = 0;
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        string sGB           = dt3.Rows[i]["GB"]?.ToString();
                        string sDANJUNG      = dt3.Rows[i]["DANJUNG"]?.ToString();
                        string sFIRST_CHAGAM = dt3.Rows[i]["FIRST_CHAGAM"]?.ToString();
                        string sHALIN        = dt3.Rows[i]["HALIN"]?.ToString();
                        string sGIDAN        = dt3.Rows[i]["GIDAN"]?.ToString();
                        string sAVGDAN = dt3.Rows[i]["AVGDAN"]?.ToString();
                        string sMGIN = dt3.Rows[i]["MGIN"]?.ToString();
                        string sMIPDAN = dt3.Rows[i]["MIPDAN"]?.ToString();
                        string sMGDAN = dt3.Rows[i]["MGDAN"]?.ToString();
                        string sIKONGKEP = dt3.Rows[i]["IKONGKEP"]?.ToString();
                        string sMKONGKEP = dt3.Rows[i]["MKONGKEP"]?.ToString();

                        iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["C" + iApplyRowIdx].Value = sGB;
                        ws.Range["F" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["G" + iApplyRowIdx].Value = sFIRST_CHAGAM;
                        ws.Range["H" + iApplyRowIdx].Value = sHALIN;
                        ws.Range["I" + iApplyRowIdx].Value = "지급단가";
                        ws.Range["J" + iApplyRowIdx].Font.Color = Color.FromArgb(255, 0, 0);
                        ws.Range["J" + iApplyRowIdx].NumberFormat = "#,##0.00;-#,##0.00";
                        ws.Range["J" + iApplyRowIdx].Value = sGIDAN;

                        //셀병합
                        Excel.Range range1 = ws.Range["K" + iApplyRowIdx, "L" + iApplyRowIdx];
                        range1.Merge();

                        ws.Range["M" + iApplyRowIdx].NumberFormat = "#,##0.00;-#,##0.00";
                        ws.Range["M" + iApplyRowIdx].Value = sAVGDAN;
                        ws.Range["N" + iApplyRowIdx].NumberFormat = "0%";
                        ws.Range["N" + iApplyRowIdx].Value = sMGIN;
                        
                        ws.Range["P" + iApplyRowIdx].Value = "매입단가";
                        ws.Range["Q" + iApplyRowIdx].Font.Color = Color.FromArgb(255, 0, 0);
                        ws.Range["Q" + iApplyRowIdx].Value = sMIPDAN;
                        ws.Range["T" + iApplyRowIdx].Value = sIKONGKEP;
                        ws.Range["U" + iApplyRowIdx].Value = sMKONGKEP;

                        //행별 설정
                        Color color = Color.White;

                        ws.Range["S" + iApplyRowIdx].NumberFormat = "0.00";

                        if (i == 0)
                        {
                            ws.Range["O" + iApplyRowIdx].Value = "매입스크랩평균단가(수입및저가제외)";
                            ws.Range["R" + iApplyRowIdx].Font.Color = Color.FromArgb(255, 0, 0);
                            ws.Range["R" + iApplyRowIdx].Value = "어음할인율";
                            ws.Range["S" + iApplyRowIdx].Font.Color = Color.FromArgb(255, 0, 0);
                            ws.Range["S" + iApplyRowIdx].Value = dtAHALIN.Rows[0]["AHALIN"]?.ToString();

                            color = Color.FromArgb(216, 228, 188);
                        }
                        else if (i == 1)
                        {
                            ws.Range["O" + iApplyRowIdx].Value = "스크랩 마진율(일계)";
                            ws.Range["K" + iApplyRowIdx].Value = "스크랩 평균단가";
                            ws.Range["R" + iApplyRowIdx].Value = "마진단가";
                            ws.Range["S" + iApplyRowIdx].Value = sMGDAN;

                            color = Color.FromArgb(183, 222, 232);
                        }
                        else if (i == 2)
                        {
                            ws.Range["O" + iApplyRowIdx].Value = "스크랩 마진율(월계)";
                            ws.Range["K" + iApplyRowIdx].Value = "스크랩 평균단가";
                            ws.Range["R" + iApplyRowIdx].Value = "마진단가";
                            ws.Range["S" + iApplyRowIdx].Value = sMGDAN;

                            color = Color.FromArgb(196, 215, 155);
                        }
                        else if (i == 3)
                        {
                            ws.Range["O" + iApplyRowIdx].Value = "G/S마진율(일계)";
                            ws.Range["K" + iApplyRowIdx].Value = "G/S 야드평균단가";
                            ws.Range["R" + iApplyRowIdx].Value = "마진단가";
                            ws.Range["S" + iApplyRowIdx].Value = sMGDAN;

                            color = Color.FromArgb(228, 223, 236);
                        }
                        else if (i == 4)
                        {
                            ws.Range["O" + iApplyRowIdx].Value = "G/S마진율(월계)";
                            ws.Range["K" + iApplyRowIdx].Value = "G/S 야드평균단가";
                            ws.Range["R" + iApplyRowIdx].Value = "마진단가";
                            ws.Range["S" + iApplyRowIdx].Value = sMGDAN;

                            color = Color.FromArgb(177, 160, 199);
                        }

                        //배경색 설정
                        Excel.Range range2 = ws.Range["A" + iApplyRowIdx+":"+"U" + iApplyRowIdx];
                        range2.Interior.Color = color;

                        range2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                    iStartRow = iApplyRowIdx;
                }

                //YK단가
                DataTable dt4 = (DataTable)GridRetr4.DataSource;

                if(dt4 != null)
                {
                    int iApplyRowIdx = 0;
                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {
                        string sDANJUNG1 = dt4.Rows[i]["DANJUNG1"]?.ToString();
                        string sKONGKEP1 = dt4.Rows[i]["KONGKEP1"]?.ToString();
                        string sDANGA1   = dt4.Rows[i]["DANGA1"]?.ToString();
                        string sDANJUNG2 = dt4.Rows[i]["DANJUNG2"]?.ToString();
                        string sKONGKEP2 = dt4.Rows[i]["KONGKEP2"]?.ToString();
                        string sDANGA2 = dt4.Rows[i]["DANGA2"]?.ToString();

                        iApplyRowIdx = iStartRow + (i + 1);

                        //셀병합
                        Excel.Range range1 = ws.Range["P" + iApplyRowIdx+":"+"R" + iApplyRowIdx];
                        range1.Merge();
                        range1.Value = "YK 전체 평균단가(일계)";
                        range1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range1.Interior.Color = Color.FromArgb(183, 222, 232);

                        Excel.Range range2 = ws.Range["S" + iApplyRowIdx+":"+ "U" + iApplyRowIdx];
                        range2.Merge();
                        range2.Value = "YK 전체 평균단가(월계)";
                        range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range2.Interior.Color = Color.FromArgb(183, 222, 232);

                        iApplyRowIdx = iApplyRowIdx + 1;

                        ws.Range["P" + iApplyRowIdx+":"+"Q"+ iApplyRowIdx].NumberFormat = "#,##0;-#,##0";
                        ws.Range["R" + iApplyRowIdx, "U" + iApplyRowIdx].NumberFormat = "#,##0.00;-#,##0.00";
                        ws.Range["P" + iApplyRowIdx + ":" + "R" + iApplyRowIdx].Interior.Color = Color.FromArgb(196, 215, 155);
                        ws.Range["P" + iApplyRowIdx].Value = sDANJUNG1;
                        ws.Range["Q" + iApplyRowIdx].Value = sKONGKEP1;
                        ws.Range["R" + iApplyRowIdx].Value = sDANGA1;
                        
                        ws.Range["S" + iApplyRowIdx + ":" + "T" + iApplyRowIdx].NumberFormat = "#,##0;-#,##0";
                        ws.Range["S" + iApplyRowIdx + ":" + "U" + iApplyRowIdx].Interior.Color = Color.FromArgb(196, 215, 155);
                        ws.Range["S" + iApplyRowIdx].Value = sDANJUNG2;
                        ws.Range["T" + iApplyRowIdx].Value = sKONGKEP2;
                        ws.Range["U" + iApplyRowIdx].Value = sDANGA2;

                        Excel.Range range3 = ws.Range["P" + iApplyRowIdx + ":" + "U" + iApplyRowIdx];
                        range3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    iStartRow = iApplyRowIdx;
                }

                //YK G/S 평균 단가
                DataTable dt5 = (DataTable)GridRetr5.DataSource;

                if (dt5 != null)
                {
                    int iApplyRowIdx = 0;
                    for (int i = 0; i < dt5.Rows.Count; i++)
                    {
                        string sDANJUNG1 = dt5.Rows[i]["DANJUNG1"]?.ToString();
                        string sKONGKEP1 = dt5.Rows[i]["KONGKEP1"]?.ToString();
                        string sDANGA1 = dt5.Rows[i]["DANGA1"]?.ToString();
                        string sDANJUNG2 = dt5.Rows[i]["DANJUNG2"]?.ToString();
                        string sKONGKEP2 = dt5.Rows[i]["KONGKEP2"]?.ToString();
                        string sDANGA2 = dt5.Rows[i]["DANGA2"]?.ToString();

                        iApplyRowIdx = iStartRow + (i + 1);

                        //셀병합
                        Excel.Range range1 = ws.Range["P" + iApplyRowIdx+":"+ "R" + iApplyRowIdx];
                        range1.Merge();
                        range1.Value = "YK G/S단가(일계)";
                        range1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range1.Interior.Color = Color.FromArgb(228, 233, 236);

                        Excel.Range range2 = ws.Range["S" + iApplyRowIdx+":"+ "U" + iApplyRowIdx];
                        range2.Merge();
                        range2.Value = "YK G/S단가(월계)";
                        range2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range2.Interior.Color = Color.FromArgb(228, 233, 236);

                        iApplyRowIdx = iApplyRowIdx + 1;

                        ws.Range["P" + iApplyRowIdx + ":" + "Q" + iApplyRowIdx].NumberFormat = "#,##0;-#,##0";
                        ws.Range["R" + iApplyRowIdx].NumberFormat = "#,##0.00;-#,##0.00";
                        ws.Range["P" + iApplyRowIdx + ":" + "R" + iApplyRowIdx].Interior.Color = Color.FromArgb(204, 192, 218);

                        ws.Range["P" + iApplyRowIdx].Value = sDANJUNG1;
                        ws.Range["Q" + iApplyRowIdx].Value = sKONGKEP1;
                        ws.Range["R" + iApplyRowIdx].Value = sDANGA1;

                        ws.Range["S" + iApplyRowIdx + ":" + "T" + iApplyRowIdx].NumberFormat = "#,##0;-#,##0";
                        ws.Range["U" + iApplyRowIdx].NumberFormat = "#,##0.00;-#,##0.00";
                        ws.Range["S" + iApplyRowIdx + ":" + "U" + iApplyRowIdx].Interior.Color = Color.FromArgb(204, 192, 218);
                        ws.Range["S" + iApplyRowIdx].Value = sDANJUNG2;
                        ws.Range["T" + iApplyRowIdx].Value = sKONGKEP2;
                        ws.Range["U" + iApplyRowIdx].Value = sDANGA2;

                        Excel.Range range3 = ws.Range["P" + iApplyRowIdx + ":" + "U" + iApplyRowIdx];
                        range3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    iStartRow = iApplyRowIdx;
                }
                #endregion


                //폐압
                Excel.Worksheet ws2 = wb.Worksheets.get_Item(2);

                //일자부분 세팅
                ws2.Range["A3"].Value = DateTime.Parse(sDate).ToString("D");

                //시작 줄
                iStartRow = 0;

                #region 폐압
                DataTable dt6 = (DataTable)GridRetr6.DataSource;

                if (dt6 != null)
                {
                    iStartRow = 4;

                    for (int i = 0; i < dt6.Rows.Count; i++)
                    {
                        string sJ_BNUM       = dt6.Rows[i]["J_BNUM"]?.ToString();
                        string sDEALER_NM    = dt6.Rows[i]["DEALER_NM"]?.ToString();
                        string sGUBUN1       = dt6.Rows[i]["GUBUN1"]?.ToString();
                        string sDANJUNG      = dt6.Rows[i]["DANJUNG"]?.ToString();
                        string sFIRST_CHAGAM = dt6.Rows[i]["FIRST_CHAGAM"]?.ToString();
                        string sHALIN        = dt6.Rows[i]["HALIN"]?.ToString();
                        string sMIDANGA      = dt6.Rows[i]["MIDANGA"]?.ToString();
                        string sDANGA        = dt6.Rows[i]["DANGA"]?.ToString();
                        string sJ_REGION     = dt6.Rows[i]["J_REGION"]?.ToString();
                        string sCKONGKEP     = dt6.Rows[i]["CKONGKEP"]?.ToString();
                        string sCKONGKEP2    = dt6.Rows[i]["CKONGKEP2"]?.ToString();
                        string sMDAN         = dt6.Rows[i]["MDAN"]?.ToString();
                        string sCHA          = dt6.Rows[i]["CHA"]?.ToString();
                        string sGUMSUBIGO    = dt6.Rows[i]["GUMSUBIGO"]?.ToString();
                        string sCHRG_NM      = dt6.Rows[i]["CHRG_NM"]?.ToString();
                        string sIKONGKEP     = dt6.Rows[i]["IKONGKEP"]?.ToString();
                        string sMAIPKEP = dt6.Rows[i]["MAIPKEP"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        if(i != dt6.Rows.Count - 1)
                        {
                            ws2.Range["A" + iApplyRowIdx].Value = i + 1;
                        }

                        //차번 텍스트형식으로 변경
                        Excel.Range range1 = ws2.Range["B" + iApplyRowIdx];
                        range1.NumberFormat = "@";

                        ws2.Range["B" + iApplyRowIdx].Value = sJ_BNUM;
                        ws2.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                        ws2.Range["D" + iApplyRowIdx].Value = sGUBUN1;
                        ws2.Range["E" + iApplyRowIdx].Value = sDANJUNG;
                        ws2.Range["F" + iApplyRowIdx].Value = sFIRST_CHAGAM;
                        ws2.Range["G" + iApplyRowIdx].Value = sHALIN;
                        ws2.Range["H" + iApplyRowIdx].Value = sMIDANGA;
                        ws2.Range["I" + iApplyRowIdx].Value = sDANGA;
                        ws2.Range["J" + iApplyRowIdx].Value = sJ_REGION;
                        ws2.Range["K" + iApplyRowIdx].Value = sCKONGKEP;
                        ws2.Range["L" + iApplyRowIdx].Value = sCKONGKEP2;
                        ws2.Range["M" + iApplyRowIdx].Value = sMDAN;
                        ws2.Range["N" + iApplyRowIdx].Value = sCHA;
                        ws2.Range["O" + iApplyRowIdx].Value = sGUMSUBIGO;
                        ws2.Range["P" + iApplyRowIdx].Value = sCHRG_NM;
                        ws2.Range["Q" + iApplyRowIdx].Value = sIKONGKEP;
                        ws2.Range["R" + iApplyRowIdx].Value = sMAIPKEP;

                        //셀 border
                        ws2.Range["A" + iApplyRowIdx + ":" + "R" + iApplyRowIdx].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        if (i == dt6.Rows.Count - 1)
                        {
                            //합계 행 디자인 변경
                            ws2.Range["A" + iApplyRowIdx + ":" + "R" + iApplyRowIdx].Interior.Color = Color.FromArgb(183, 222, 232);
                            ws2.Range["C" + iApplyRowIdx].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            ws2.Range["I" + iApplyRowIdx].NumberFormat = "0.00";
                            ws2.Range["M" + iApplyRowIdx].NumberFormat = "0.00";

                            ws2.Range["M" + iApplyRowIdx].Font.Color = Color.Red;
                        }
                    }
                }
                #endregion


                //직납
                Excel.Worksheet ws3 = wb.Worksheets.get_Item(3);

                //일자부분 세팅
                ws3.Range["A2"].Value = DateTime.Parse(sDate).ToString("D");

                //시작 줄
                iStartRow = 3;

                #region 매입
                DataTable dt7 = (DataTable)GridRetr7.DataSource;

                if(dt7 != null)
                {
                    int iApplyRowIdx = 0;
                    for (int i = 0; i < dt7.Rows.Count; i++)
                    {                                                         
                        string sROWNUM          = dt7.Rows[i]["ROWNUM"]?.ToString();
                        string sJ_BNUM          = dt7.Rows[i]["J_BNUM"]?.ToString();
                        string sPURC_DEALER_NM  = dt7.Rows[i]["PURC_DEALER_NM"]?.ToString();
                        string sGUBUN1          = dt7.Rows[i]["GUBUN1"]?.ToString();
                        string sSALEUNITPRICE   = dt7.Rows[i]["SALEUNITPRICE"]?.ToString();
                        string sJ_REGION        = dt7.Rows[i]["J_REGION"]?.ToString();
                        string sJERDAN          = dt7.Rows[i]["JERDAN"]?.ToString();
                        string sPURCHUNITPRICE  = dt7.Rows[i]["PURCHUNITPRICE"]?.ToString();
                        string sCARRYCOST       = dt7.Rows[i]["CARRYCOST"]?.ToString();
                        string sARRIVEUNITPRICE = dt7.Rows[i]["ARRIVEUNITPRICE"]?.ToString();
                        string sOWEIGHT         = dt7.Rows[i]["OWEIGHT"]?.ToString();
                        string sHALIN           = dt7.Rows[i]["HALIN"]?.ToString();
                        string sDIFFUNITPRICE   = dt7.Rows[i]["DIFFUNITPRICE"]?.ToString();
                        string sAVGDAN          = dt7.Rows[i]["AVGDAN"]?.ToString();
                        string sSALEPRICE       = dt7.Rows[i]["SALEPRICE"]?.ToString();
                        string sPURC_AMT = dt7.Rows[i]["PURC_AMT"]?.ToString();

                        iApplyRowIdx = iStartRow + (i + 1);

                        if (sROWNUM.Equals("999999997"))
                        {
                            //ws3.Range["A" + iApplyRowIdx].Value = sPURC_DEALER_NM;
                        }
                        else if (sROWNUM.Equals("999999998"))
                        {
                            //ws3.Range["A" + iApplyRowIdx].Value = sPURC_DEALER_NM;
                        }
                        else if (sROWNUM.Equals("999999999"))
                        {
                            //ws3.Range["A" + iApplyRowIdx].Value = sPURC_DEALER_NM;
                        }
                        else
                        {
                            if(i > 2)
                            {
                                Excel.Range range = ws3.Rows[iApplyRowIdx];
                                range.Insert();
                            }

                            //차번 텍스트형식으로 변경
                            Excel.Range range1 = ws3.Range["A"+ iApplyRowIdx, "B" + iApplyRowIdx];
                            range1.NumberFormat = "@";

                            ws3.Range["A" + iApplyRowIdx].Value = i+1;
                            ws3.Range["B" + iApplyRowIdx].Value = sJ_BNUM;
                            ws3.Range["C" + iApplyRowIdx].Value = sPURC_DEALER_NM;
                        }

                        ws3.Range["D" + iApplyRowIdx].Value = sGUBUN1;
                        ws3.Range["E" + iApplyRowIdx].Value = sSALEUNITPRICE;
                        ws3.Range["F" + iApplyRowIdx].Value = sJ_REGION;
                        ws3.Range["G" + iApplyRowIdx].Value = sJERDAN;
                        ws3.Range["H" + iApplyRowIdx].Value = sPURCHUNITPRICE;
                        ws3.Range["I" + iApplyRowIdx].Value = sCARRYCOST;
                        ws3.Range["J" + iApplyRowIdx].Value = sARRIVEUNITPRICE;
                        ws3.Range["K" + iApplyRowIdx].Value = sOWEIGHT;
                        ws3.Range["L" + iApplyRowIdx].Value = sHALIN;
                        ws3.Range["M" + iApplyRowIdx].NumberFormat = "#,##0;[red]-#,##0";
                        ws3.Range["M" + iApplyRowIdx].Value = sDIFFUNITPRICE;
                        ws3.Range["N" + iApplyRowIdx].Value = sAVGDAN;
                        ws3.Range["O" + iApplyRowIdx].Value = sSALEPRICE;
                        ws3.Range["P" + iApplyRowIdx].Value = sPURC_AMT;

                        ws3.Range["A" + iApplyRowIdx + ":" + "P" + iApplyRowIdx].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                    iStartRow = iApplyRowIdx;
                }
                #endregion

                #region 매출

                DataTable dt8 = (DataTable)GridRetr8.DataSource;

                if(dt8 != null)
                {
                    iStartRow += 1;

                    if(iStartRow < 10)
                    {
                        iStartRow = 10;
                    }
                    int iApplyRowIdx = 0;
                    for (int i = 0; i < dt8.Rows.Count; i++)
                    {
                        string sJ_BNUM    = dt8.Rows[i]["J_BNUM"]?.ToString();
                        string sGUBUN1    = dt8.Rows[i]["GUBUN1"]?.ToString();
                        string sGIDAN     = dt8.Rows[i]["GIDAN"]?.ToString();
                        string sDANJUNG   = dt8.Rows[i]["DANJUNG"]?.ToString();
                        string sHALIN     = dt8.Rows[i]["HALIN"]?.ToString();
                        string sKONGKEP   = dt8.Rows[i]["KONGKEP"]?.ToString();
                        string sCKONGKEP  = dt8.Rows[i]["CKONGKEP"]?.ToString();
                        string sSUNKEP    = dt8.Rows[i]["SUNKEP"]?.ToString();
                        string sSUNDAN    = dt8.Rows[i]["SUNDAN"]?.ToString();
                        string sGAMSU     = dt8.Rows[i]["GAMSU"]?.ToString();
                        string sJANGBE    = dt8.Rows[i]["JANGBE"]?.ToString();
                        string sNWEIGHT   = dt8.Rows[i]["NWEIGHT"]?.ToString();
                        string sCWEIGHT   = dt8.Rows[i]["CWEIGHT"]?.ToString();
                        string sCHAWEIGHT = dt8.Rows[i]["CHAWEIGHT"]?.ToString();
                        string sGMDANJUNG = dt8.Rows[i]["GMDANJUNG"]?.ToString();

                        iApplyRowIdx = iStartRow + (i + 1);

                        if (i > 16)
                        {
                            Excel.Range range = ws3.Rows[iApplyRowIdx];
                            range.Insert();
                        }

                        ws3.Range["A" + iApplyRowIdx].Value = i + 1;

                        //차번 텍스트형식으로 변경
                        Excel.Range range1 = ws3.Range["A" + iApplyRowIdx, "B" + iApplyRowIdx];
                        range1.NumberFormat = "@";

                        ws3.Range["B" + iApplyRowIdx].Value = sJ_BNUM;
                        ws3.Range["C" + iApplyRowIdx].Value = sGUBUN1;
                        ws3.Range["D" + iApplyRowIdx].Value = sGIDAN;
                        ws3.Range["E" + iApplyRowIdx].Value = sDANJUNG;
                        ws3.Range["F" + iApplyRowIdx].Value = sHALIN;
                        ws3.Range["G" + iApplyRowIdx].Value = sKONGKEP;
                        ws3.Range["H" + iApplyRowIdx].Value = sCKONGKEP;
                        ws3.Range["I" + iApplyRowIdx].Value = sSUNKEP;
                        ws3.Range["J" + iApplyRowIdx].Value = sSUNDAN;
                        ws3.Range["K" + iApplyRowIdx].Value = sGAMSU;
                        ws3.Range["L" + iApplyRowIdx].Value = sJANGBE;
                        ws3.Range["M" + iApplyRowIdx].Value = sNWEIGHT;
                        ws3.Range["N" + iApplyRowIdx].Value = sCWEIGHT;

                        ws3.Range["O" + iApplyRowIdx].NumberFormat = "#,##0;[red]-#,##0";
                        ws3.Range["O" + iApplyRowIdx].Value = sCHAWEIGHT;
                        ws3.Range["P" + iApplyRowIdx].Value = sGMDANJUNG;

                        ws3.Range["A" + iApplyRowIdx + ":" + "P" + iApplyRowIdx].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                    if(dt8.Rows.Count < 16)
                    {
                        iStartRow = iStartRow+17;
                    }
                    else
                    {
                        iStartRow = iApplyRowIdx;
                    }
                }

                //매출일계
                int rowNumCarCnt = iStartRow + 4;

                if(rowNumCarCnt < 31)
                {
                    rowNumCarCnt = 31;
                }
                string sCarCnt = TxtCarCnt.EditValue?.ToString();
                ws3.Range["D" + rowNumCarCnt].Value = sCarCnt;

                #endregion

                #region ASR
                iStartRow += 7;

                if(iStartRow < 34)
                {
                    iStartRow = 34;
                }

                DataTable dt9 = (DataTable)GridRetr9.DataSource;

                if(dt9 != null)
                {
                    for(int i=0;i<dt9.Rows.Count; i++)
                    {
                        string sJ_BNUM    = dt9.Rows[i]["J_BNUM"]?.ToString();
                        string sDEALER_NM = dt9.Rows[i]["DEALER_NM"]?.ToString();
                        string sGUBUN1    = dt9.Rows[i]["GUBUN1"]?.ToString();
                        string sDANGA     = dt9.Rows[i]["DANGA"]?.ToString();
                        string sWEIGHT    = dt9.Rows[i]["WEIGHT"]?.ToString();
                        string sGAMW      = dt9.Rows[i]["GAMW"]?.ToString();
                        string sPRICE     = dt9.Rows[i]["PRICE"]?.ToString();
                        string sBIGO = dt9.Rows[i]["BIGO"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        //차번 텍스트형식으로 변경
                        Excel.Range range1 = ws3.Range["A" + iApplyRowIdx, "B" + iApplyRowIdx];
                        range1.NumberFormat = "@";

                        if (i > 2)
                        {
                            Excel.Range range = ws3.Rows[iApplyRowIdx];
                            range.Insert();

                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }

                        ws3.Range["B" + iApplyRowIdx].Value = sJ_BNUM;
                        ws3.Range["C" + iApplyRowIdx].Value = sDEALER_NM;
                        ws3.Range["D" + iApplyRowIdx].Value = sGUBUN1;
                        ws3.Range["E" + iApplyRowIdx].Value = sDANGA;
                        ws3.Range["F" + iApplyRowIdx].Value = sWEIGHT;
                        ws3.Range["G" + iApplyRowIdx].Value = sGAMW;
                        ws3.Range["H" + iApplyRowIdx].Value = sPRICE;
                        ws3.Range["I" + iApplyRowIdx].Value = sBIGO;
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

        private void DailyUnitCost_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                //BtnClose.PerformClick();
            }
            else if (e.KeyCode == Keys.F3)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel.PerformClick();
            }
        }

        private void TxtOgs1_EditValueChanged(object sender, EventArgs e)
        {
            string sTxtOgs1 = TxtOgs1.EditValue?.ToString();
            string sTxtOgs2 = TxtOgs2.EditValue?.ToString();

            double.TryParse(sTxtOgs1, out double dTxtOgs1);
            double.TryParse(sTxtOgs2, out double dTxtOgs2);

            double result = dTxtOgs2 / dTxtOgs1 * 100;

            TxtOgs3.EditValue = Math.Round(result) + "%";
        }

        private void TxtOW1_EditValueChanged(object sender, EventArgs e)
        {
            string sTxtOW1 = TxtOW1.EditValue?.ToString();
            string sTxtOW2 = TxtOW2.EditValue?.ToString();

            double.TryParse(sTxtOW1, out double dTxtOW1);
            double.TryParse(sTxtOW2, out double dTxtOW2);

            double result = dTxtOW2 / dTxtOW1 * 100;

            TxtOW3.EditValue = Math.Round(result) + "%";
        }

        private void TxtYd1_EditValueChanged(object sender, EventArgs e)
        {
            string sTxtYd1 = TxtYd1.EditValue?.ToString();
            string sTxtYd2 = TxtYd2.EditValue?.ToString();

            double.TryParse(sTxtYd1, out double dTxtYd1);
            double.TryParse(sTxtYd2, out double dTxtYd2);

            double result = dTxtYd2 / dTxtYd1 * 100;

            TxtYd3.EditValue = Math.Round(result) + "%";
        }

        private void TxtYk1_EditValueChanged(object sender, EventArgs e)
        {
            string sTxtYk1 = TxtYk1.EditValue?.ToString();
            string sTxtYk2 = TxtYk2.EditValue?.ToString();

            double.TryParse(sTxtYk1, out double dTxtYk1);
            double.TryParse(sTxtYk2, out double dTxtYk2);

            double result = dTxtYk2 / dTxtYk1 * 100;

            TxtYk3.EditValue = Math.Round(result) + "%";
        }

        private void DateEditDT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewRetr1_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sName = GridViewRetr1.GetRowCellValue(e.RowHandle, "NAME")?.ToString();

            if (sName != null && sName.Equals("스크랩계"))
            {
                e.Appearance.BackColor = Color.LightGreen;
            }
            else if(sName != null && sName.Equals("합 계"))
            {
                e.Appearance.BackColor = Color.LightGreen;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr2_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void BtnAUSV_Click(object sender, EventArgs e)
        {
            string sAHALIN = TxtAHALIN.EditValue?.ToString();
            string sDate = DateEditDT.EditValue?.ToString().Substring(0, 10);

            if (string.IsNullOrEmpty(sAHALIN))
            {
                XtraMessageBox.Show("변경할 어음할인율을 입력해주세요.");
                TxtAHALIN.Focus();
                return;
            }

            StringBuilder strSql = new StringBuilder();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" IF EXISTS(SELECT J_DATE FROM MEAIPSILJUK WHERE J_DATE = '" + sDate + "' AND GUBUN = '어음할인') ");
                strSql.AppendLine("    BEGIN                                                                              ");
                strSql.AppendLine("          UPDATE MEAIPSILJUK                                                           ");
                strSql.AppendLine("             SET DATAVALUE = " + sAHALIN + "                                                         ");
                strSql.AppendLine("           WHERE J_DATE = '" + sDate + "' AND GUBUN = '어음할인'                          ");
                strSql.AppendLine("      END                                                                              ");
                strSql.AppendLine(" ELSE                                                                                  ");
                strSql.AppendLine("    BEGIN                                                                              ");
                strSql.AppendLine("          INSERT INTO MEAIPSILJUK(J_DATE                                               ");
                strSql.AppendLine("                                 , DAEGUBUN                                            ");
                strSql.AppendLine("                                 , GUBUN                                               ");
                strSql.AppendLine("                                 , DATAVALUE)                                          ");
                strSql.AppendLine("                           VALUES('" + sDate + "'                                                   ");
                strSql.AppendLine("                                 , '기초자료'                                          ");
                strSql.AppendLine("                                 , '어음할인'                                          ");
                strSql.AppendLine("                                 , " + sAHALIN + ")                                                  ");
                strSql.AppendLine("      END                                                                              ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

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

        private void GridRetrView6_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sNum = GridRetrView6.GetRowCellValue(e.RowHandle, "NUM")?.ToString();

            if(!string.IsNullOrEmpty(sNum) && sNum.Equals("2"))
            {
                e.Appearance.BackColor = Color.LightGreen;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr7_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sNum = GridViewRetr7.GetRowCellValue(e.RowHandle, "ROWNUM")?.ToString();

            if (!string.IsNullOrEmpty(sNum) && (sNum.Equals("999999997") || sNum.Equals("999999998")))
            {
                e.Appearance.BackColor = Color.LightYellow;
            }
            else if (!string.IsNullOrEmpty(sNum) && sNum.Equals("999999999"))
            {
                e.Appearance.BackColor = Color.LightGreen;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void RepoTxtNGong_KeyDown(object sender, KeyEventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;

            string sVal = textEdit.EditValue?.ToString();
            GridViewRetr8.SetFocusedRowCellValue("NWEIGHT", sVal);

            string sNGong = GridViewRetr8.GetFocusedRowCellValue("NWEIGHT")?.ToString();
            string sGong = GridViewRetr8.GetFocusedRowCellValue("CWEIGHT")?.ToString();

            double dNGong = 0;
            double dGong = 0;

            double.TryParse(sNGong, out dNGong);
            double.TryParse(sGong, out dGong);

            GridViewRetr8.SetFocusedRowCellValue("CHAWEIGHT", dNGong - dGong);
        }

        private void RepoTxtGong_KeyDown(object sender, KeyEventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;

            string sVal = textEdit.EditValue?.ToString();
            GridViewRetr8.SetFocusedRowCellValue("CWEIGHT", sVal);

            string sNGong = GridViewRetr8.GetFocusedRowCellValue("NWEIGHT")?.ToString();
            string sGong = GridViewRetr8.GetFocusedRowCellValue("CWEIGHT")?.ToString();

            double dNGong = 0;
            double dGong = 0;

            double.TryParse(sNGong, out dNGong);
            double.TryParse(sGong, out dGong);

            GridViewRetr8.SetFocusedRowCellValue("CHAWEIGHT", dNGong - dGong);
        }

        private void BtnSaveJkw_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)GridRetr8.DataSource;

            if (dt == null || dt.Rows.Count < 0)
            {
                XtraMessageBox.Show("수정한 내용이 없습니다.");
                return;
            }

            DataSet dtSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dtSave.Tables[0];

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                if (dtMerge.Rows.Count > 0)
                {
                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        string sJUNPYOID  = dtMerge.Rows[i]["JUNPYOID"]?.ToString();
                        string sJ_DATE    = dtMerge.Rows[i]["J_DATE"]?.ToString();
                        string sGAMSU     = dtMerge.Rows[i]["GAMSU"]?.ToString();
                        string sJANGBE    = dtMerge.Rows[i]["JANGBE"]?.ToString();
                        string sNWEIGHT   = dtMerge.Rows[i]["NWEIGHT"]?.ToString();
                        string sCWEIGHT   = dtMerge.Rows[i]["CWEIGHT"]?.ToString();
                        string sCHAWEIGHT = dtMerge.Rows[i]["CHAWEIGHT"]?.ToString();

                        strSql.AppendLine(" IF EXISTS(SELECT JUNPYOID FROM JKNAP_W WHERE JUNPYOID = " + sJUNPYOID + " AND J_DATE = '"+ sJ_DATE + "')");
                        strSql.AppendLine("    BEGIN                                                                            ");
                        strSql.AppendLine("          UPDATE JKNAP_W                                                             ");
                        strSql.AppendLine("             SET GAMSU = '"+ sGAMSU + "'                                                          ");
			            strSql.AppendLine("               , JANGBE = '"+ sJANGBE + "'                                                         ");
			            strSql.AppendLine("               , NWEIGHT = "+ sNWEIGHT + "                                                         ");
			            strSql.AppendLine("               , CWEIGHT = "+ sCWEIGHT + "                                                         ");
			            strSql.AppendLine("               , CHAWEIGHT = "+ sCHAWEIGHT + "                                                       ");
                        strSql.AppendLine("           WHERE JUNPYOID = "+ sJUNPYOID + " AND J_DATE = '"+ sJ_DATE + "'                              ");
                        strSql.AppendLine("      END                                                                            ");
                        strSql.AppendLine(" ELSE                                                                                ");
                        strSql.AppendLine("    BEGIN                                                                            ");
                        strSql.AppendLine("          INSERT INTO JKNAP_W(JUNPYOID                                               ");
                        strSql.AppendLine("                             , J_DATE                                                ");
                        strSql.AppendLine("                             , GAMSU                                                 ");
                        strSql.AppendLine("                             , JANGBE                                                ");
                        strSql.AppendLine("                             , NWEIGHT                                               ");
                        strSql.AppendLine("                             , CWEIGHT                                               ");
                        strSql.AppendLine("                             , CHAWEIGHT )                                           ");
                        strSql.AppendLine("                       VALUES("+ sJUNPYOID + "                                                      ");
                        strSql.AppendLine("                             , '"+ sJ_DATE + "'                                                    ");
                        strSql.AppendLine("                             , '"+ sGAMSU + "'                                                    ");
                        strSql.AppendLine("                             , '"+ sJANGBE + "'                                                    ");
                        strSql.AppendLine("                             , "+ sNWEIGHT + "                                                     ");
                        strSql.AppendLine("                             , "+ sCWEIGHT + "                                                     ");
                        strSql.AppendLine("                             , "+ sCHAWEIGHT + ")                                                    ");
                        strSql.AppendLine("      END                                                                            ");
                                                                                                                                
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    XtraMessageBox.Show("저장을 완료했습니다.");

                    BtnRetr.PerformClick();
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                MessageBox.Show(ex.Message);
            }
            finally
            {
                DBConn.dbTran = null;
            }
        }

        private void DailyUnitCost_TextChanged(object sender, EventArgs e)
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

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateEditDT.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateEditDT.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateEditDT.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateEditDT.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }
    }
}