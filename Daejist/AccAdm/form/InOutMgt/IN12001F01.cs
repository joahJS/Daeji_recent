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
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Helpers;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace AccAdm
{
    public partial class IN12001F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN12001F01()
        {
            InitializeComponent();
        }

        private void IN12001F01_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();

            DateYMD.EditValue = DateTime.Today.AddDays(-1) ;
            BtnRetr_Click(null, null);
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { BGridViewRetr1, BGridViewRetr5, BGridViewRetr3, BGridViewRetr4 };
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
            Cursor = Cursors.WaitCursor;
            string sDateYmd = DateYMD.EditValue?.ToString().Substring(0,10);
            string sDateFrom = sDateYmd.Substring(0, 7) + "-01";

            DateTime.TryParse(sDateFrom, out DateTime FromDate);
            DateTime ToDate = FromDate.AddMonths(1).AddDays(-1);

            int iCnt = 0;

            DateTime temp = FromDate;
            while (true)
            {
                if (temp > ToDate)
                    break;

                if(temp.DayOfWeek != DayOfWeek.Sunday)
                {
                    iCnt++;
                }

                temp = temp.AddDays(1);
            }
            //년월
            string sYYMM = sDateYmd.Replace("-", "").Substring(0, 6);


            //월 근무일수 조회 =====================================================================================
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT yymm, WorkDays ");
            strSql.AppendLine("  FROM MonWork ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "' ");

            DataTable dt_Wday = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt_Wday.Rows.Count >= 1) iCnt = Convert.ToInt16(dt_Wday.Rows[0][1].ToString());
            tE_WorkDay.Text = Convert.ToString( iCnt);



            string sDD = iCnt.ToString();

            

            #region 그리드 따로 분리해서 조회
            #region 매출
            //매출목표량
            strSql.Clear();
            strSql.AppendLine(" SELECT SUM(O_GS) AS O_GS        ");//g/s
            strSql.AppendLine("      , SUM(O_WEIGHT) AS O_WEIGHT");//중량
            strSql.AppendLine("      , SUM(O_YK) AS O_YK        ");
            strSql.AppendLine("      , SUM(O_SD) AS O_SD        ");//슈레더
            strSql.AppendLine("      , SUM(O_GS) + SUM(O_WEIGHT) + SUM(O_SD) AS TOT");
            strSql.AppendLine("   FROM SALEMAECHUL              ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "'    ");

            DataTable dtMCMOK = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtMCMOK != null)
            {
                double.TryParse(dtMCMOK.Rows[0]["O_SD"]?.ToString(), out double c1);
                double.TryParse(dtMCMOK.Rows[0]["O_WEIGHT"]?.ToString(), out double w1);
                double.TryParse(dtMCMOK.Rows[0]["O_GS"]?.ToString(), out double g1);
                double.TryParse(dtMCMOK.Rows[0]["TOT"]?.ToString(), out double t1);

                GridBandC.Caption = c1.ToString("#,0 T");
                GridBandW.Caption = w1.ToString("#,0 T");
                GridBandG.Caption = g1.ToString("#,0 T");
                GridBandTotMeChul.Caption = t1.ToString("#,0 T");

                double.TryParse(dtMCMOK.Rows[0]["O_YK"]?.ToString(), out double yk1);
                GridBandJK.Caption = (c1 + w1 + g1 + yk1).ToString("#,0 T");
            }
            else
            {
                GridBandC.Caption = "";
                GridBandW.Caption = "";
                GridBandG.Caption = "";
                GridBandTotMeChul.Caption = "";
                GridBandJK.Caption = "";

            }

            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH DATERANGE(DATES) AS (                                                      ");
            strSql.AppendLine("      SELECT CONVERT(DATE, '"+sDateFrom+"')--시작일자                               ");
            strSql.AppendLine("       UNION ALL                                                                 ");
            strSql.AppendLine("      SELECT DATEADD(D, 1, DATES)                                                ");
            strSql.AppendLine("        FROM DATERANGE                                                           ");
            strSql.AppendLine("       WHERE DATES < CONVERT(DATE, '"+sDateYmd+"')--종료일자                       ");
            strSql.AppendLine("  ), MOK1 AS(                                                                    ");
            strSql.AppendLine("      --매출                                                                     ");
            strSql.AppendLine("      SELECT SUM(O_GS) AS O_GS --G / S                                           ");
            strSql.AppendLine("           , SUM(O_WEIGHT) AS O_WEIGHT --중량                                    ");
            strSql.AppendLine("           , SUM(O_YK) AS O_YK                                                   ");
            strSql.AppendLine("           , SUM(O_SD) AS O_SD --슈레더                                          ");
            strSql.AppendLine("        FROM SALEMAECHUL                                                         ");
            strSql.AppendLine("       WHERE YYMM = '"+sYYMM+"'                                                     ");
            strSql.AppendLine("  ), TEMP1 AS(                                                                   ");
            strSql.AppendLine("      --매출 슈레더 현대:6541754044 YK:6531121044 대한: 6541233044                               ");
            strSql.AppendLine("      SELECT A1.DATES                                                            ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                 ");
            strSql.AppendLine("                  ELSE(SELECT ISNULL(O_SD / "+ sDD + ", 0) FROM MOK1) END AS MOKDAY      ");
            strSql.AppendLine("           , X1.YK                                                               ");
            strSql.AppendLine("           , X1.GITA                                                             ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 AND X1.GITA IS NULL THEN NULL                          ");
            strSql.AppendLine("                  ELSE ISNULL(X1.YK, 0) + ISNULL(X1.GITA, 0)                               ");
            strSql.AppendLine("                            END CHA                                   ");
            strSql.AppendLine("        FROM DATERANGE A1                                                                                     ");
            strSql.AppendLine("        LEFT JOIN(SELECT Z1.J_DATE                                                                            ");
            strSql.AppendLine("                         , SUM(CASE WHEN Z1.J_ID1 = 6541754044 THEN Z1.DANJUNG END) AS YK                     ");
            strSql.AppendLine("                         , SUM(CASE WHEN Z1.J_ID1 NOT IN(6541754044) THEN Z1.DANJUNG END) AS GITA ");
            strSql.AppendLine("                      FROM(SELECT A.J_DATE                                                                    ");
            strSql.AppendLine("                                 , A.J_ID1                                                                    ");
            strSql.AppendLine("                                 , SUM(A.DANJUNG) / 1000 AS DANJUNG                                           ");
            strSql.AppendLine("                              FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                         ");
            strSql.AppendLine("                             WHERE A.J_SERIAL = B.J_SERIAL                                                    ");
            strSql.AppendLine("                               AND A.J_ID1 = D.DEALER_CD                                                      ");
            strSql.AppendLine("                               AND D.CHRG_ID = E.EMP_ID                                                       ");
            strSql.AppendLine("                               AND A.KERATYPE = '매출'                                                        ");
            strSql.AppendLine("                               AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                             ");
            strSql.AppendLine("                               AND A.J_LOTNO <> '4'                                                           ");
            strSql.AppendLine("                               AND B.DAEGUBUN = '슈레더'                                                      ");
            strSql.AppendLine("                             GROUP BY A.J_DATE, A.J_ID1)Z1                                                    ");
            strSql.AppendLine("                     GROUP BY Z1.J_DATE) X1                                                                   ");
            strSql.AppendLine("          ON A1.DATES = X1.J_DATE                                                                             ");
            strSql.AppendLine("  ), TEMP2 AS(                                                                                                ");
            strSql.AppendLine("      --매출 중량                                                                                             ");
            strSql.AppendLine("      SELECT A1.DATES                                                                                         ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                              ");
            strSql.AppendLine("                      ELSE(SELECT ISNULL(O_WEIGHT / "+sDD+", 0) FROM MOK1) END AS MOKDAY                           ");
            strSql.AppendLine("           , X1.YK                                                                                            ");
            strSql.AppendLine("           , X1.GITA                                                                                          ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 AND X1.GITA IS NULL THEN NULL                          ");
            strSql.AppendLine("                  ELSE ISNULL(X1.YK, 0) + ISNULL(X1.GITA, 0)                               ");
            strSql.AppendLine("                            END CHA                               ");
            strSql.AppendLine("        FROM DATERANGE A1                                                                                     ");
            strSql.AppendLine("        LEFT JOIN(SELECT Z1.J_DATE                                                                            ");
            strSql.AppendLine("                           , SUM(CASE WHEN Z1.J_ID1 = 6541754044 THEN Z1.DANJUNG END) AS YK                   ");
            strSql.AppendLine("                           , SUM(CASE WHEN Z1.J_ID1 NOT IN(6541754044) THEN Z1.DANJUNG END) AS GITA");
            strSql.AppendLine("                        FROM(SELECT A.J_DATE                                                                   ");
            strSql.AppendLine("                                   , A.J_ID1                                                                   ");
            strSql.AppendLine("                                   , SUM(A.DANJUNG) / 1000 AS DANJUNG                                          ");
            strSql.AppendLine("                                FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                        ");
            strSql.AppendLine("                               WHERE A.J_SERIAL = B.J_SERIAL                                                   ");
            strSql.AppendLine("                                 AND A.J_ID1 = D.DEALER_CD                                                     ");
            strSql.AppendLine("                                 AND D.CHRG_ID = E.EMP_ID                                                      ");
            strSql.AppendLine("                                 AND A.KERATYPE = '매출'                                                       ");
            strSql.AppendLine("                                 AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                            ");
            strSql.AppendLine("                                 AND A.J_LOTNO <> '4'                                                          ");
            strSql.AppendLine("                                 AND B.DAEGUBUN = '고철A'                                                      ");
            strSql.AppendLine("                               GROUP BY A.J_DATE, A.J_ID1)Z1                                                   ");
            strSql.AppendLine("                       GROUP BY Z1.J_DATE) X1                                                                  ");
            strSql.AppendLine("          ON A1.DATES = X1.J_DATE                                                                              ");
            strSql.AppendLine("  ), TEMP3 AS(                                                                                                 ");
            strSql.AppendLine("      --매출 G / S                                                                                             ");
            strSql.AppendLine("      SELECT A1.DATES                                                                                          ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            strSql.AppendLine("                          ELSE(SELECT ISNULL(O_GS / "+sDD+", 0) FROM MOK1) END AS MOKDAY                            ");
            strSql.AppendLine("           , X1.YK                                                                                             ");
            strSql.AppendLine("           , X1.GITA                                                                                           ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 AND X1.GITA IS NULL THEN NULL                           ");
            strSql.AppendLine("                  ELSE ISNULL(X1.YK, 0)  + ISNULL(X1.GITA, 0)                                ");
            strSql.AppendLine("                            END CHA                                    ");
            strSql.AppendLine("        FROM DATERANGE A1                                                                                      ");
            strSql.AppendLine("        LEFT JOIN(SELECT Z1.J_DATE                                                                             ");
            strSql.AppendLine("                         , SUM(CASE WHEN Z1.J_ID1 = 6541754044 THEN Z1.DANJUNG END) AS YK                      ");
            strSql.AppendLine("                         , SUM(CASE WHEN Z1.J_ID1 NOT IN(6541754044) THEN Z1.DANJUNG END) AS GITA  ");
            strSql.AppendLine("                      FROM(SELECT A.J_DATE                                                                     ");
            strSql.AppendLine("                                 , A.J_ID1                                                                     ");
            strSql.AppendLine("                                 , SUM(A.DANJUNG) / 1000 AS DANJUNG                                            ");
            strSql.AppendLine("                              FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                          ");
            strSql.AppendLine("                             WHERE A.J_SERIAL = B.J_SERIAL                                                     ");
            strSql.AppendLine("                               AND A.J_ID1 = D.DEALER_CD                                                       ");
            strSql.AppendLine("                               AND D.CHRG_ID = E.EMP_ID                                                        ");
            strSql.AppendLine("                               AND A.KERATYPE = '매출'                                                         ");
            strSql.AppendLine("                               AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                              ");
            strSql.AppendLine("                               AND A.J_LOTNO <> '4'                                                            ");
            strSql.AppendLine("                               AND B.DAEGUBUN = '고철B'                                                        ");
            strSql.AppendLine("                             GROUP BY A.J_DATE, A.J_ID1)Z1                                                     ");
            strSql.AppendLine("                     GROUP BY Z1.J_DATE) X1                                                                    ");
            strSql.AppendLine("          ON A1.DATES = X1.J_DATE                                                                              ");
            strSql.AppendLine("  ), TEMP4 AS(                                                                                                 ");
            strSql.AppendLine("     --ASR                                                                                                     ");
            strSql.AppendLine("     SELECT J_DATE                                                                                             ");
            strSql.AppendLine("          , SUM(OWEIGHT) / 1000 AS WEIGHT                                                                      ");
            strSql.AppendLine("       FROM MESURING                                                                                           ");
            strSql.AppendLine("      WHERE J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                                       ");
            strSql.AppendLine("        AND GUBUN1 = 'ASR'                                                                                     ");
            strSql.AppendLine("        AND J_CHECK = '1'                                                                                      ");
            strSql.AppendLine("      GROUP BY J_DATE                                                                                          ");
            strSql.AppendLine("  )                                                                                                            ");
            strSql.AppendLine("                                                                                                               ");
            strSql.AppendLine(" SELECT T1.DATES                                                                                               ");
            strSql.AppendLine("       , T1.MOKDAY AS MOKDAY1                                                                                  ");
            strSql.AppendLine("       , T1.YK AS YK1                                                                                          ");
            strSql.AppendLine("       , T1.GITA AS GITA1                                                                                      ");
            strSql.AppendLine("       , T1.CHA AS CHA1                                                                                        ");
            strSql.AppendLine("       , T2.MOKDAY AS MOKDAY2                                                                                  ");
            strSql.AppendLine("       , T2.YK AS YK2                                                                                          ");
            strSql.AppendLine("       , T2.GITA AS GITA2                                                                                      ");
            strSql.AppendLine("       , T2.CHA AS CHA2                                                                                        ");
            strSql.AppendLine("       , T3.MOKDAY AS MOKDAY3                                                                                  ");
            strSql.AppendLine("       , T3.YK AS YK3                                                                                          ");
            strSql.AppendLine("       , T3.GITA AS GITA3                                                                                      ");
            strSql.AppendLine("       , T3.CHA AS CHA3                                                                                        ");
            strSql.AppendLine("       , ISNULL(T1.MOKDAY, 0) + ISNULL(T2.MOKDAY, 0) + ISNULL(T3.MOKDAY, 0) AS TMOK                            ");
            strSql.AppendLine("       , ISNULL(T1.YK, 0)+ISNULL(T2.YK, 0) + ISNULL(T3.YK, 0) AS TYK                                     ");
            strSql.AppendLine("       , ISNULL(T1.GITA, 0)+ISNULL(T2.GITA, 0) + ISNULL(T3.GITA, 0) AS TGITA                   ");
            strSql.AppendLine("       , (ISNULL(T1.YK, 0) + ISNULL(T2.YK, 0) + ISNULL(T3.YK, 0)                          ");
            strSql.AppendLine("          + ISNULL(T1.GITA, 0) + ISNULL(T2.GITA, 0) + ISNULL(T3.GITA, 0))                ");
            strSql.AppendLine("           AS TCHA                        ");
            strSql.AppendLine("       , T4.WEIGHT AS ASR                                                                                ");
            strSql.AppendLine("    FROM TEMP1 T1                                                                                        ");
            strSql.AppendLine("    LEFT JOIN TEMP2 T2                                                                                         ");
            strSql.AppendLine("      ON T1.DATES = T2.DATES                                                                                   ");
            strSql.AppendLine("    LEFT JOIN TEMP3 T3                                                                                         ");
            strSql.AppendLine("      ON T1.DATES = T3.DATES                                                                                   ");
            strSql.AppendLine("    LEFT JOIN TEMP4 T4                                                                                         ");
            strSql.AppendLine("      ON T1.DATES = T4.J_DATE                                                                                  ");

            DataTable dtRetr1 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl1.DataSource = dtRetr1;
            #endregion

            #region 재고이동
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH DATERANGE(DATES) AS(                                                         ");
            strSql.AppendLine("      SELECT CONVERT(DATE, '"+sDateFrom+"')--시작일자                                 ");
            strSql.AppendLine("       UNION ALL                                                                   ");
            strSql.AppendLine("      SELECT DATEADD(D, 1, DATES)                                                  ");
            strSql.AppendLine("        FROM DATERANGE                                                             ");
            strSql.AppendLine("       WHERE DATES < CONVERT(DATE, '"+sDateYmd+"')--종료일자                         ");
            strSql.AppendLine("  ), TEMP5 AS(                                                                     ");
            strSql.AppendLine("      --재고이동 CP(슈 - 고철B)                                                    ");
            strSql.AppendLine("      SELECT A1.BASDT                                                              ");
            strSql.AppendLine("           , SUM(A1.WEIGHT) / 1000 AS WEIGHT                                       ");
            strSql.AppendLine("           , CASE WHEN COUNT(*) = 0 THEN NULL ELSE COUNT(*) END AS CNT             ");
            strSql.AppendLine("        FROM J_MAGAM A1                                                            ");
            strSql.AppendLine("       WHERE A1.BASDT BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                        ");
            strSql.AppendLine("         AND A1.J_SERIAL = 5059072                                                 ");
            strSql.AppendLine("       GROUP BY A1.BASDT                                                           ");
            strSql.AppendLine("  ), TEMP6 AS(                                                                     ");
            strSql.AppendLine("        --재고이동 DUST(이동량: 슈레더S:4045122, 납품대수: 슈레더B:2020163 로 변경)");
            strSql.AppendLine("        SELECT J_DATE                                                              ");
            strSql.AppendLine("             , SUM(CASE WHEN J_SERIAL = 4045122 THEN DANJUNG END)/ 1000 AS DANJUNG ");
            strSql.AppendLine("             , CASE WHEN COUNT(CASE WHEN J_SERIAL = 2020163 THEN 1 END) = 0 THEN NULL ELSE COUNT(CASE WHEN J_SERIAL = 2020163 THEN 1 END) END AS CNT");
            strSql.AppendLine("          FROM INLIST                                                                                                                               ");
            strSql.AppendLine("         WHERE J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                                                                         ");
            strSql.AppendLine("         GROUP BY J_DATE                                                                                                                            ");
            strSql.AppendLine("  )                                                                                                                                                 ");
            strSql.AppendLine("                                                                                                                                                    ");
            strSql.AppendLine("  SELECT A1.DATES                                                                                                                                   ");
            strSql.AppendLine("          , T1.WEIGHT AS W1                                                                                                                         ");
            strSql.AppendLine("          , T1.CNT AS CNT1                                                                                                                          ");
            strSql.AppendLine("          , T2.DANJUNG AS W2                                                                                                                        ");
            strSql.AppendLine("          , T2.CNT AS CNT2                                                                                                                          ");
            strSql.AppendLine("          , (SELECT SUM(WEIGHT) FROM TEMP5) AS TOTCP                                                                                                ");
            strSql.AppendLine("          , (SELECT SUM(DANJUNG) FROM TEMP6) AS TOTDUST                                                                                             ");
            strSql.AppendLine("       FROM DATERANGE A1                                                                                                                            ");
            strSql.AppendLine("       LEFT JOIN TEMP5 T1                                                                                                                           ");
            strSql.AppendLine("         ON A1.DATES = T1.BASDT                                                                                                                     ");
            strSql.AppendLine("       LEFT JOIN TEMP6 T2                                                                                                                           ");
            strSql.AppendLine("         ON A1.DATES = T2.J_DATE                                                                                                                    ");
            strSql.AppendLine("       ORDER BY A1.DATES");

            DataTable dtRetr3 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl3.DataSource = dtRetr3;

            if (dtRetr3 != null)
            {
                double dTotCP = 0;
                double dTotDust = 0;

                double.TryParse(dtRetr3.Rows[0]["TOTCP"]?.ToString(), out dTotCP);
                double.TryParse(dtRetr3.Rows[0]["TOTDUST"]?.ToString(), out dTotDust);

                GridBandC2.Caption = dTotCP.ToString("#,0 T");
                GridBandD.Caption = dTotDust.ToString("#,0 T");
            }
            else
            {
                GridBandC2.Caption = "";
                GridBandD.Caption = "";
            }
            #endregion

            #region 매입
            //매입목표량
            strSql.Clear();
            strSql.AppendLine(" SELECT SUM(I_LIGHT) + SUM(I_YK) AS I_LIGHT  ");//g/s 모재 + 경량C
            strSql.AppendLine("      , SUM(I_WEIGHT) AS I_WEIGHT");//중량
            strSql.AppendLine("      , SUM(I_YK) AS I_YK        ");//경량C
            strSql.AppendLine("      , SUM(I_CHAPI) AS I_CHAPI  ");//폐압
            strSql.AppendLine("      , SUM(I_LIGHT) + SUM(I_YK) + SUM(I_WEIGHT) + SUM(I_CHAPI) AS TOT");
            strSql.AppendLine("    FROM SALEMAEIP               ");
            strSql.AppendLine("   WHERE YYMM = '" + sYYMM + "'   ");

            DataTable dtMIPMOK = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtMIPMOK != null)
            {
                double.TryParse(dtMIPMOK.Rows[0]["I_CHAPI"]?.ToString(), out double p1);
                double.TryParse(dtMIPMOK.Rows[0]["I_WEIGHT"]?.ToString(), out double w2);
                double.TryParse(dtMIPMOK.Rows[0]["I_LIGHT"]?.ToString(), out double g2);
                double.TryParse(dtMIPMOK.Rows[0]["TOT"]?.ToString(), out double t1);

                GridBandPe.Caption = p1.ToString("#,0 T");
                GridBandWe.Caption = w2.ToString("#,0 T");
                GridBandGs.Caption = g2.ToString("#,0 T");
                GridBandTotMeip.Caption = t1.ToString("#,0 T");
            }
            else
            {
                GridBandPe.Caption = "";
                GridBandWe.Caption = "";
                GridBandGs.Caption = "";
                GridBandTotMeip.Caption = "";
            }

            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH DATERANGE(DATES) AS(                                                     ");
            strSql.AppendLine("      SELECT CONVERT(DATE, '"+ sDateFrom + "')--시작일자                             ");
            strSql.AppendLine("       UNION ALL                                                               ");
            strSql.AppendLine("      SELECT DATEADD(D, 1, DATES)                                              ");
            strSql.AppendLine("        FROM DATERANGE                                                         ");
            strSql.AppendLine("       WHERE DATES < CONVERT(DATE, '"+ sDateYmd + "')--종료일자                     ");
            strSql.AppendLine("  ), MOK2 AS(                                                                  ");
            strSql.AppendLine("      --매입 목표량                                                            ");
            strSql.AppendLine("      SELECT SUM(I_LIGHT) AS I_LIGHT --G / S                                   ");
            strSql.AppendLine("           , SUM(I_WEIGHT) AS I_WEIGHT --중량                                  ");
            strSql.AppendLine("           , SUM(I_YK) AS I_YK                                                 ");
            strSql.AppendLine("           , SUM(I_CHAPI) AS I_CHAPI --폐압                                    ");
            strSql.AppendLine("        FROM SALEMAEIP                                                         ");
            strSql.AppendLine("       WHERE YYMM = '"+sYYMM+"'                                                   ");
            strSql.AppendLine("  ), TEMP7 AS(                                                                 ");
            strSql.AppendLine("                                                                               ");
            strSql.AppendLine("                                                                               ");
            strSql.AppendLine("      --매입 폐압                                                              ");
            strSql.AppendLine("      SELECT A1.DATES                                                          ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL               ");
            strSql.AppendLine("                  ELSE(SELECT ISNULL(I_CHAPI / "+sDD+", 0) FROM MOK2) END AS MOKDAY ");
            strSql.AppendLine("           , A2.DANJUNG                                                        ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL               ");
            strSql.AppendLine("                  ELSE ISNULL(A2.DANJUNG,0) -(SELECT ISNULL(I_CHAPI / "+sDD+", 0) FROM MOK2) END AS CHA ");
            strSql.AppendLine("           FROM DATERANGE A1                                                                       ");
            strSql.AppendLine("        LEFT JOIN(SELECT A.J_DATE                                                                  ");
            strSql.AppendLine("                        , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                          ");
            strSql.AppendLine("                     FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                       ");
            strSql.AppendLine("                    WHERE A.J_SERIAL = B.J_SERIAL                                                  ");
            strSql.AppendLine("                      AND A.J_ID1 = D.DEALER_CD                                                    ");
            strSql.AppendLine("                      AND D.CHRG_ID = E.EMP_ID                                                     ");
            strSql.AppendLine("                      AND A.KERATYPE = '매입'                                                      ");
            strSql.AppendLine("                      AND A.J_DATE BETWEEN '"+ sDateFrom + "' AND '"+ sDateYmd + "'                           ");
            strSql.AppendLine("                      AND A.J_LOTNO <> '4'                                                         ");
            strSql.AppendLine("                      AND B.DAEGUBUN = '슈레더'                                                    ");
            strSql.AppendLine("                    GROUP BY A.J_DATE) A2                                                          ");
            strSql.AppendLine("          ON A1.DATES = A2.J_DATE                                                                  ");
            strSql.AppendLine("  ), TEMP8 AS(                                                                                     ");
            strSql.AppendLine("                                                                                                   ");
            strSql.AppendLine("                                                                                                   ");
            strSql.AppendLine("      --매입 중량(고철A)                                                                           ");
            strSql.AppendLine("      SELECT A1.DATES                                                                              ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                   ");
            strSql.AppendLine("                  ELSE(SELECT ISNULL(I_WEIGHT / "+sDD+", 0) FROM MOK2) END AS MOKDAY                    ");
            strSql.AppendLine("           , A2.DANJUNG                                                                            ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                   ");
            strSql.AppendLine("                  ELSE ISNULL(A2.DANJUNG,0) -(SELECT ISNULL(I_WEIGHT / "+sDD+", 0) FROM MOK2) END AS CHA");
            strSql.AppendLine("           FROM DATERANGE A1                                                                       ");
            strSql.AppendLine("        LEFT JOIN(SELECT A.J_DATE                                                                  ");
            strSql.AppendLine("                        , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                          ");
            strSql.AppendLine("                     FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                       ");
            strSql.AppendLine("                    WHERE A.J_SERIAL = B.J_SERIAL                                                  ");
            strSql.AppendLine("                      AND A.J_ID1 = D.DEALER_CD                                                    ");
            strSql.AppendLine("                      AND D.CHRG_ID = E.EMP_ID                                                     ");
            strSql.AppendLine("                      AND A.KERATYPE = '매입'                                                      ");
            strSql.AppendLine("                      AND A.J_DATE BETWEEN '"+ sDateFrom + "' AND '"+ sDateYmd + "'                           ");
            strSql.AppendLine("                      AND A.J_LOTNO <> '4'                                                         ");
            strSql.AppendLine("                      AND B.DAEGUBUN = '고철A'                                                     ");
            strSql.AppendLine("                    GROUP BY A.J_DATE ) A2                                                         ");
            strSql.AppendLine("          ON A1.DATES = A2.J_DATE                                                                  ");
            strSql.AppendLine("  ), TEMP9 AS(                                                                                     ");
            strSql.AppendLine("                                                                                                   ");
            strSql.AppendLine("                                                                                                   ");
            strSql.AppendLine("      --매입 G / S모재(고철B)                                                                      ");
            strSql.AppendLine("      SELECT A1.DATES                                                                              ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                   ");
            strSql.AppendLine("                  ELSE(SELECT ISNULL(I_LIGHT / "+sDD+", 0) FROM MOK2) END AS MOKDAY                     ");
            strSql.AppendLine("           , A2.DANJUNG                                                                            ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                   ");
            strSql.AppendLine("                  ELSE ISNULL(A2.DANJUNG,0) -(SELECT ISNULL(I_LIGHT / "+sDD+", 0) FROM MOK2) END AS CHA ");
            strSql.AppendLine("           FROM DATERANGE A1                                                                       ");
            strSql.AppendLine("        LEFT JOIN(SELECT A.J_DATE                                                                  ");
            strSql.AppendLine("                        , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                          ");
            strSql.AppendLine("                     FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                       ");
            strSql.AppendLine("                    WHERE A.J_SERIAL = B.J_SERIAL                                                  ");
            strSql.AppendLine("                      AND A.J_ID1 = D.DEALER_CD                                                    ");
            strSql.AppendLine("                      AND D.CHRG_ID = E.EMP_ID                                                     ");
            strSql.AppendLine("                      AND A.KERATYPE = '매입'                                                      ");
            strSql.AppendLine("                      AND A.J_DATE BETWEEN '"+ sDateFrom + "' AND '"+ sDateYmd + "'                           ");
            strSql.AppendLine("                      AND A.J_LOTNO <> '4'                                                         ");
            strSql.AppendLine("                      AND B.DAEGUBUN = '고철B'                                                     ");
            strSql.AppendLine("                    GROUP BY A.J_DATE) A2                                                          ");
            strSql.AppendLine("          ON A1.DATES = A2.J_DATE                                                                  ");
            strSql.AppendLine("  )                                                                                                ");
            strSql.AppendLine("                                                                                                   ");
            strSql.AppendLine("  SELECT T1.DATES                                                                                  ");
            strSql.AppendLine("       , T1.MOKDAY AS MOKDAY1                                                                      ");
            strSql.AppendLine("       , T1.DANJUNG AS W1                                                                          ");
            strSql.AppendLine("       , T1.CHA AS CHA1                                                                            ");
            strSql.AppendLine("       , T2.MOKDAY AS MOKDAY2                                                                      ");
            strSql.AppendLine("       , T2.DANJUNG AS W2                                                                          ");
            strSql.AppendLine("       , T2.CHA AS CHA2                                                                            ");
            strSql.AppendLine("       , T3.MOKDAY AS MOKDAY3                                                                      ");
            strSql.AppendLine("       , T3.DANJUNG AS W3                                                                          ");
            strSql.AppendLine("       , T3.CHA AS CHA3                                                                            ");
            strSql.AppendLine("       , ISNULL(T1.MOKDAY, 0) + ISNULL(T2.MOKDAY, 0) + ISNULL(T3.MOKDAY, 0) AS TMOK                ");
            strSql.AppendLine("       , ISNULL(T1.DANJUNG, 0)+ISNULL(T2.DANJUNG, 0) + ISNULL(T3.DANJUNG, 0) AS TW          ");
            strSql.AppendLine("       , (ISNULL(T1.DANJUNG, 0) + ISNULL(T2.DANJUNG, 0) + ISNULL(T3.DANJUNG, 0))       ");
            strSql.AppendLine("          -(ISNULL(T1.MOKDAY, 0) + ISNULL(T2.MOKDAY, 0) + ISNULL(T3.MOKDAY, 0)) AS TCHA            ");
            strSql.AppendLine("    FROM TEMP7 T1                                                                                  ");
            strSql.AppendLine("    LEFT JOIN TEMP8 T2                                                                             ");
            strSql.AppendLine("      ON T1.DATES = T2.DATES                                                                       ");
            strSql.AppendLine("    LEFT JOIN TEMP9 T3                                                                             ");
            strSql.AppendLine("      ON T1.DATES = T3.DATES                                                                       ");

            DataTable dtRetr2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl5.DataSource = dtRetr2;
            #endregion
            #endregion

            #region 매입출재고 분리형으로 변경(요청)  2022-06-20 이전
            //   //매출 목표량
            //   strSql.Clear();
            //   strSql.AppendLine(" SELECT SUM(O_GS) AS O_GS        ");//g/s
            //   strSql.AppendLine("      , SUM(O_WEIGHT) AS O_WEIGHT");//중량
            //   strSql.AppendLine("      , SUM(O_YK) AS O_YK        ");
            //   strSql.AppendLine("      , SUM(O_SD) AS O_SD        ");//슈레더
            //   strSql.AppendLine("   FROM SALEMAECHUL              ");
            //   strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "'    ");

            //   DataTable dtMCMOK = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //   if (dtMCMOK != null)
            //   {
            //       double.TryParse(dtMCMOK.Rows[0]["O_SD"]?.ToString(), out double c1);
            //       double.TryParse(dtMCMOK.Rows[0]["O_WEIGHT"]?.ToString(), out double w1);
            //       double.TryParse(dtMCMOK.Rows[0]["O_GS"]?.ToString(), out double g1);

            //       GridBandMSu.Caption = c1.ToString("#,0 T");
            //       GridBandMW1.Caption = w1.ToString("#,0 T");
            //       GridBandMG1.Caption = g1.ToString("#,0 T");

            //       double.TryParse(dtMCMOK.Rows[0]["O_YK"]?.ToString(), out double yk1);
            //       GridBandJK.Caption = (c1 + w1 + g1 + yk1).ToString("#,0 T");
            //   }
            //   else
            //   {
            //       GridBandMSu.Caption = "";
            //       GridBandMW1.Caption = "";
            //       GridBandMG1.Caption = "";
            //       GridBandJK.Caption = "";
            //   }

            //   //매입목표량
            //   strSql.Clear();
            //   strSql.AppendLine(" SELECT SUM(I_LIGHT) AS I_LIGHT  ");//g/s 모재
            //   strSql.AppendLine("      , SUM(I_WEIGHT) AS I_WEIGHT");//중량
            //   strSql.AppendLine("      , SUM(I_YK) AS I_YK        ");
            //   strSql.AppendLine("      , SUM(I_CHAPI) AS I_CHAPI  ");//폐압
            //   strSql.AppendLine("    FROM SALEMAEIP               ");
            //   strSql.AppendLine("   WHERE YYMM = '" + sYYMM + "'   ");

            //   DataTable dtMIPMOK = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //   if (dtMIPMOK != null)
            //   {
            //       double.TryParse(dtMIPMOK.Rows[0]["I_CHAPI"]?.ToString(), out double p1);
            //       double.TryParse(dtMIPMOK.Rows[0]["I_WEIGHT"]?.ToString(), out double w2);
            //       double.TryParse(dtMIPMOK.Rows[0]["I_LIGHT"]?.ToString(), out double g2);

            //       GridBandMP.Caption = p1.ToString("#,0 T");
            //       GridBandMW2.Caption = w2.ToString("#,0 T");
            //       GridBandMG2.Caption = g2.ToString("#,0 T");
            //   }
            //   else
            //   {
            //       GridBandMP.Caption = "";
            //       GridBandMW2.Caption = "";
            //       GridBandMG2.Caption = "";
            //   }

            //   strSql.Clear();
            //   strSql.AppendLine(" WITH DATERANGE(DATES) AS                                                               ");
            //   strSql.AppendLine(" (                                                                                      ");
            //   strSql.AppendLine("     SELECT CONVERT(DATE, '"+sDateFrom+"')--시작일자                                       ");
            //   strSql.AppendLine("      UNION ALL                                                                         ");
            //   strSql.AppendLine("     SELECT DATEADD(D, 1, DATES)                                                        ");
            //   strSql.AppendLine("       FROM DATERANGE                                                                   ");
            //   strSql.AppendLine("      WHERE DATES < CONVERT(DATE, '"+sDateYmd+"')--종료일자                               ");
            //   strSql.AppendLine(" ), MOK1 AS(                                                                            ");
            //   strSql.AppendLine("     --매출                                                                             ");
            //   strSql.AppendLine("     SELECT SUM(O_GS) AS O_GS --G / S                                                   ");
            //   strSql.AppendLine("          , SUM(O_WEIGHT) AS O_WEIGHT --중량                                            ");
            //   strSql.AppendLine("          , SUM(O_YK) AS O_YK                                                           ");
            //   strSql.AppendLine("          , SUM(O_SD) AS O_SD --슈레더                                                  ");
            //   strSql.AppendLine("       FROM SALEMAECHUL                                                                 ");
            //   strSql.AppendLine("      WHERE YYMM = '"+sYYMM+"'                                                             ");
            //   strSql.AppendLine(" ), TEMP1 AS(                                                                           ");
            //   strSql.AppendLine("     --매출 슈레더 YK:6531121044 대한: 6541233044                                       ");
            //   strSql.AppendLine("     SELECT A1.DATES                                                                    ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                         ");
            //   strSql.AppendLine("                 ELSE(SELECT ISNULL(O_SD / "+sDD+", 0) FROM MOK1) END AS MOKDAY              ");
            //   strSql.AppendLine("          , X1.YK                                                                       ");
            //   strSql.AppendLine("          , X1.DH                                                                       ");
            //   strSql.AppendLine("          , X1.GITA                                                                     ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 AND X1.GITA IS NULL THEN NULL     ");
            //   strSql.AppendLine("                 ELSE ISNULL(X1.YK, 0) + ISNULL(X1.DH, 0) + ISNULL(X1.GITA, 0)          ");
            //   strSql.AppendLine("                         - (SELECT ISNULL(O_SD / "+sDD+", 0) FROM MOK1) END CHA              ");
            //   strSql.AppendLine("       FROM DATERANGE A1                                                                ");
            //   strSql.AppendLine("       LEFT JOIN(SELECT Z1.J_DATE                                                       ");
            //   strSql.AppendLine("                        , SUM(CASE WHEN Z1.J_ID1 = 6531121044 THEN Z1.DANJUNG END) AS YK");
            //   strSql.AppendLine("                        , SUM(CASE WHEN Z1.J_ID1 = 6541233044 THEN Z1.DANJUNG END) AS DH");
            //   strSql.AppendLine("                        , SUM(CASE WHEN Z1.J_ID1 NOT IN(6531121044, 6541233044) THEN Z1.DANJUNG END) AS GITA");
            //   strSql.AppendLine("                     FROM(SELECT A.J_DATE                                            ");
            //   strSql.AppendLine("                                , A.J_ID1                                            ");
            //   strSql.AppendLine("                                , SUM(A.DANJUNG) / 1000 AS DANJUNG                   ");
            //   strSql.AppendLine("                             FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E ");
            //   strSql.AppendLine("                            WHERE A.J_SERIAL = B.J_SERIAL                            ");
            //   strSql.AppendLine("                              AND A.J_ID1 = D.DEALER_CD                              ");
            //   strSql.AppendLine("                              AND D.CHRG_ID = E.EMP_ID                               ");
            //   strSql.AppendLine("                              AND A.KERATYPE = '매출'                                ");
            //   strSql.AppendLine("                              AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'     ");
            //   strSql.AppendLine("                              AND A.J_LOTNO <> '4'                                   ");
            //   strSql.AppendLine("                              AND B.DAEGUBUN = '슈레더'                              ");
            //   strSql.AppendLine("                            GROUP BY A.J_DATE, A.J_ID1)Z1                            ");
            //   strSql.AppendLine("                    GROUP BY Z1.J_DATE) X1                                           ");
            //   strSql.AppendLine("         ON A1.DATES = X1.J_DATE                                                     ");
            //   strSql.AppendLine(" ), TEMP2 AS(                                                                        ");
            //   strSql.AppendLine("     --매출 중량                                                                     ");
            //   strSql.AppendLine("     SELECT A1.DATES                                                                 ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                      ");
            //   strSql.AppendLine("                     ELSE(SELECT ISNULL(O_WEIGHT / "+sDD+", 0) FROM MOK1) END AS MOKDAY   ");
            //   strSql.AppendLine("          , X1.YK                                                                    ");
            //   strSql.AppendLine("          , X1.DH                                                                    ");
            //   strSql.AppendLine("          , X1.GITA                                                                  ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 AND X1.GITA IS NULL THEN NULL  ");
            //   strSql.AppendLine("                 ELSE ISNULL(X1.YK, 0) + ISNULL(X1.DH, 0) + ISNULL(X1.GITA, 0)       ");
            //   strSql.AppendLine("                         - (SELECT ISNULL(O_WEIGHT / "+sDD+", 0) FROM MOK1) END CHA       ");
            //   strSql.AppendLine("       FROM DATERANGE A1                                                             ");
            //   strSql.AppendLine("       LEFT JOIN(SELECT Z1.J_DATE                                                    ");
            //   strSql.AppendLine("                          , SUM(CASE WHEN Z1.J_ID1 = 6531121044 THEN Z1.DANJUNG END) AS YK");
            //   strSql.AppendLine("                          , SUM(CASE WHEN Z1.J_ID1 = 6541233044 THEN Z1.DANJUNG END) AS DH");
            //   strSql.AppendLine("                          , SUM(CASE WHEN Z1.J_ID1 NOT IN(6531121044, 6541233044) THEN Z1.DANJUNG END) AS GITA");
            //   strSql.AppendLine("                       FROM(SELECT A.J_DATE                                                                   ");
            //   strSql.AppendLine("                                  , A.J_ID1                                                                   ");
            //   strSql.AppendLine("                                  , SUM(A.DANJUNG) / 1000 AS DANJUNG                                          ");
            //   strSql.AppendLine("                               FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                        ");
            //   strSql.AppendLine("                              WHERE A.J_SERIAL = B.J_SERIAL                                                   ");
            //   strSql.AppendLine("                                AND A.J_ID1 = D.DEALER_CD                                                     ");
            //   strSql.AppendLine("                                AND D.CHRG_ID = E.EMP_ID                                                      ");
            //   strSql.AppendLine("                                AND A.KERATYPE = '매출'                                                       ");
            //   strSql.AppendLine("                                AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                            ");
            //   strSql.AppendLine("                                AND A.J_LOTNO <> '4'                                                          ");
            //   strSql.AppendLine("                                AND B.DAEGUBUN = '고철A'                                                      ");
            //   strSql.AppendLine("                              GROUP BY A.J_DATE, A.J_ID1)Z1                                                   ");
            //   strSql.AppendLine("                      GROUP BY Z1.J_DATE) X1                                                                  ");
            //   strSql.AppendLine("         ON A1.DATES = X1.J_DATE                                                                              ");
            //   strSql.AppendLine(" ), TEMP3 AS(                                                                                                 ");
            //   strSql.AppendLine("     --매출 G / S                                                                                             ");
            //   strSql.AppendLine("     SELECT A1.DATES                                                                                          ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                         ELSE(SELECT ISNULL(O_GS / "+sDD+", 0) FROM MOK1) END AS MOKDAY                            ");
            //   strSql.AppendLine("          , X1.YK                                                                                             ");
            //   strSql.AppendLine("          , X1.DH                                                                                             ");
            //   strSql.AppendLine("          , X1.GITA                                                                                           ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 AND X1.GITA IS NULL THEN NULL                           ");
            //   strSql.AppendLine("                 ELSE ISNULL(X1.YK, 0) + ISNULL(X1.DH, 0) + ISNULL(X1.GITA, 0)                                ");
            //   strSql.AppendLine("                         - (SELECT ISNULL(O_GS / "+sDD+", 0) FROM MOK1) END CHA                                    ");
            //   strSql.AppendLine("       FROM DATERANGE A1                                                                                      ");
            //   strSql.AppendLine("       LEFT JOIN(SELECT Z1.J_DATE                                                                             ");
            //   strSql.AppendLine("                        , SUM(CASE WHEN Z1.J_ID1 = 6531121044 THEN Z1.DANJUNG END) AS YK                      ");
            //   strSql.AppendLine("                        , SUM(CASE WHEN Z1.J_ID1 = 6541233044 THEN Z1.DANJUNG END) AS DH                      ");
            //   strSql.AppendLine("                        , SUM(CASE WHEN Z1.J_ID1 NOT IN(6531121044, 6541233044) THEN Z1.DANJUNG END) AS GITA  ");
            //   strSql.AppendLine("                     FROM(SELECT A.J_DATE                                                                     ");
            //   strSql.AppendLine("                                , A.J_ID1                                                                     ");
            //   strSql.AppendLine("                                , SUM(A.DANJUNG) / 1000 AS DANJUNG                                            ");
            //   strSql.AppendLine("                             FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                          ");
            //   strSql.AppendLine("                            WHERE A.J_SERIAL = B.J_SERIAL                                                     ");
            //   strSql.AppendLine("                              AND A.J_ID1 = D.DEALER_CD                                                       ");
            //   strSql.AppendLine("                              AND D.CHRG_ID = E.EMP_ID                                                        ");
            //   strSql.AppendLine("                              AND A.KERATYPE = '매출'                                                         ");
            //   strSql.AppendLine("                              AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                              ");
            //   strSql.AppendLine("                              AND A.J_LOTNO <> '4'                                                            ");
            //   strSql.AppendLine("                              AND B.DAEGUBUN = '고철B'                                                        ");
            //   strSql.AppendLine("                            GROUP BY A.J_DATE, A.J_ID1)Z1                                                     ");
            //   strSql.AppendLine("                    GROUP BY Z1.J_DATE) X1                                                                    ");
            //   strSql.AppendLine("         ON A1.DATES = X1.J_DATE                                                                              ");
            //   strSql.AppendLine(" ), TEMP4 AS(                                                                                                 ");
            //   strSql.AppendLine("    --ASR                                                                                                     ");
            //   strSql.AppendLine("    SELECT J_DATE                                                                                             ");
            //   strSql.AppendLine("         , SUM(OWEIGHT) / 1000 AS WEIGHT                                                                      ");
            //   strSql.AppendLine("      FROM MESURING                                                                                           ");
            //   strSql.AppendLine("     WHERE J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                                       ");
            //   strSql.AppendLine("       AND GUBUN1 = 'ASR'                                                                                     ");
            //   strSql.AppendLine("       AND J_CHECK = '1'                                                                                      ");
            //   strSql.AppendLine("     GROUP BY J_DATE                                                                                          ");
            //   strSql.AppendLine(" ), MACHTB AS(                                                                                                ");
            //   strSql.AppendLine("     SELECT T1.DATES                                                                                          ");
            //   strSql.AppendLine("      , T1.MOKDAY AS MOKDAY1                                                                                  ");
            //   strSql.AppendLine("      , T1.YK AS YK1                                                                                          ");
            //   strSql.AppendLine("      , T1.DH AS DH1                                                                                          ");
            //   strSql.AppendLine("      , T1.GITA AS GITA1                                                                                      ");
            //   strSql.AppendLine("      , T1.CHA AS CHA1                                                                                        ");
            //   strSql.AppendLine("      , T2.MOKDAY AS MOKDAY2                                                                                  ");
            //   strSql.AppendLine("      , T2.YK AS YK2                                                                                          ");
            //   strSql.AppendLine("      , T2.DH AS DH2                                                                                          ");
            //   strSql.AppendLine("      , T2.GITA AS GITA2                                                                                      ");
            //   strSql.AppendLine("      , T2.CHA AS CHA2                                                                                        ");
            //   strSql.AppendLine("      , T3.MOKDAY AS MOKDAY3                                                                                  ");
            //   strSql.AppendLine("      , T3.YK AS YK3                                                                                          ");
            //   strSql.AppendLine("      , T3.DH AS DH3                                                                                          ");
            //   strSql.AppendLine("      , T3.GITA AS GITA3                                                                                      ");
            //   strSql.AppendLine("      , T3.CHA AS CHA3                                                                                        ");
            //   strSql.AppendLine("      , T4.WEIGHT AS ASR                                                                                      ");
            //   strSql.AppendLine("   FROM TEMP1 T1                                                                                              ");
            //   strSql.AppendLine("   LEFT JOIN TEMP2 T2                                                                                         ");
            //   strSql.AppendLine("     ON T1.DATES = T2.DATES                                                                                   ");
            //   strSql.AppendLine("   LEFT JOIN TEMP3 T3                                                                                         ");
            //   strSql.AppendLine("     ON T1.DATES = T3.DATES                                                                                   ");
            //   strSql.AppendLine("   LEFT JOIN TEMP4 T4                                                                                         ");
            //   strSql.AppendLine("     ON T1.DATES = T4.J_DATE                                                                                  ");
            //   strSql.AppendLine(" ), TEMP5 AS(                                                                                                 ");
            //   strSql.AppendLine("     --재고이동 CP(슈 - 고철B)                                                                                ");
            //   strSql.AppendLine("     SELECT A1.BASDT                                                                                          ");
            //   strSql.AppendLine("          , SUM(A1.WEIGHT) / 1000 AS WEIGHT                                                                   ");
            //   strSql.AppendLine("          , CASE WHEN COUNT(*) = 0 THEN NULL ELSE COUNT(*) END AS CNT                                                                                   ");
            //   strSql.AppendLine("       FROM J_MAGAM A1                                                                                        ");
            //   strSql.AppendLine("      WHERE A1.BASDT BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                                    ");
            //   strSql.AppendLine("        AND A1.J_SERIAL = 5059072                                                                             ");
            //   strSql.AppendLine("      GROUP BY A1.BASDT                                                                                       ");
            //   strSql.AppendLine(" ), TEMP6 AS(                                                                                                 ");
            //   //strSql.AppendLine("     --재고이동 DUST(슈레더S)                                                                                 ");
            //   //strSql.AppendLine("     SELECT J_DATE                                                                                            ");
            //   //strSql.AppendLine("          , SUM(DANJUNG)/ 1000 AS DANJUNG                                                                     ");
            //   //strSql.AppendLine("          , COUNT(*) AS CNT                                                                                   ");
            //   //strSql.AppendLine("       FROM INLIST                                                                                            ");
            //   //strSql.AppendLine("      WHERE J_SERIAL = 4045122                                                                                ");
            //   //strSql.AppendLine("        AND J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                                      ");
            //   //strSql.AppendLine("      GROUP BY J_DATE                                                                                         ");
            //   //요청에 의해 2022-02-03일 변경
            //   strSql.AppendLine("       --재고이동 DUST(이동량: 슈레더S:4045122, 납품대수: 슈레더B:2020163 로 변경)");
            //   strSql.AppendLine("       SELECT J_DATE                                                              ");
            //   strSql.AppendLine("            , SUM(CASE WHEN J_SERIAL = 4045122 THEN DANJUNG END)/ 1000 AS DANJUNG ");
            //   strSql.AppendLine("            , CASE WHEN COUNT(CASE WHEN J_SERIAL = 2020163 THEN 1 END) = 0 THEN NULL ELSE COUNT(CASE WHEN J_SERIAL = 2020163 THEN 1 END) END AS CNT               ");
            //   strSql.AppendLine("         FROM INLIST                                                              ");
            //   strSql.AppendLine("        WHERE J_DATE BETWEEN '" + sDateFrom + "' AND '" + sDateYmd + "'                        ");
            //   strSql.AppendLine("        GROUP BY J_DATE                                                           ");
            //   strSql.AppendLine(" ), JEGOMV AS(                                                                                                ");
            //   strSql.AppendLine("      SELECT A1.DATES                                                                                         ");
            //   strSql.AppendLine("         , T1.WEIGHT AS W1                                                                                    ");
            //   strSql.AppendLine("         , T1.CNT AS CNT1                                                                                     ");
            //   strSql.AppendLine("         , T2.DANJUNG AS W2                                                                                   ");
            //   strSql.AppendLine("         , T2.CNT AS CNT2                                                                                     ");
            //   strSql.AppendLine("         , (SELECT SUM(WEIGHT) FROM TEMP5) AS TOTCP                                                           ");
            //   strSql.AppendLine("         , (SELECT SUM(DANJUNG) FROM TEMP6) AS TOTDUST                                                        ");
            //   strSql.AppendLine("      FROM DATERANGE A1                                                                                       ");
            //   strSql.AppendLine("      LEFT JOIN TEMP5 T1                                                                                      ");
            //   strSql.AppendLine("        ON A1.DATES = T1.BASDT                                                                                ");
            //   strSql.AppendLine("      LEFT JOIN TEMP6 T2                                                                                      ");
            //   strSql.AppendLine("        ON A1.DATES = T2.J_DATE                                                                               ");
            //   strSql.AppendLine(" ), MOK2 AS(                                                                                                  ");
            //   strSql.AppendLine("                                                                                                              ");
            //   strSql.AppendLine("     --매입 목표량                                                                                            ");
            //   strSql.AppendLine("     SELECT SUM(I_LIGHT) AS I_LIGHT --G / S                                                                   ");
            //   strSql.AppendLine("          , SUM(I_WEIGHT) AS I_WEIGHT --중량                                                                  ");
            //   strSql.AppendLine("          , SUM(I_YK) AS I_YK                                                                                 ");
            //   strSql.AppendLine("          , SUM(I_CHAPI) AS I_CHAPI --폐압                                                                    ");
            //   strSql.AppendLine("       FROM SALEMAEIP                                                                                         ");
            //   strSql.AppendLine("      WHERE YYMM = '"+sYYMM+"'                                                                                   ");
            //   strSql.AppendLine(" ), TEMP7 AS(                                                                                                 ");
            //   strSql.AppendLine("                                                                                                              ");
            //   strSql.AppendLine("     --매입 폐압                                                                                              ");
            //   strSql.AppendLine("     SELECT A1.DATES                                                                                          ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                 ELSE(SELECT ISNULL(I_CHAPI / "+sDD+", 0) FROM MOK2) END AS MOKDAY                                 ");
            //   strSql.AppendLine("          , A2.DANJUNG                                                                                        ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                 ELSE ISNULL(A2.DANJUNG,0) -(SELECT ISNULL(I_CHAPI / "+sDD+", 0) FROM MOK2) END AS CHA             ");
            //   strSql.AppendLine("          FROM DATERANGE A1                                                                                   ");
            //   strSql.AppendLine("       LEFT JOIN(SELECT A.J_DATE                                                                              ");
            //   strSql.AppendLine("                       , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                                      ");
            //   strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                                   ");
            //   strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                                                              ");
            //   strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                                                                ");
            //   strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                                                                 ");
            //   strSql.AppendLine("                     AND A.KERATYPE = '매입'                                                                  ");
            //   strSql.AppendLine("                     AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                       ");
            //   strSql.AppendLine("                     AND A.J_LOTNO <> '4'                                                                     ");
            //   strSql.AppendLine("                     AND B.DAEGUBUN = '슈레더'                                                                ");
            //   strSql.AppendLine("                   GROUP BY A.J_DATE) A2                                                                      ");
            //   strSql.AppendLine("         ON A1.DATES = A2.J_DATE                                                                              ");
            //   strSql.AppendLine(" ), TEMP8 AS(                                                                                                 ");
            //   strSql.AppendLine("                                                                                                              ");
            //   strSql.AppendLine("     --매입 중량(고철A)                                                                                       ");
            //   strSql.AppendLine("     SELECT A1.DATES                                                                                          ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                 ELSE(SELECT ISNULL(I_WEIGHT / "+sDD+", 0) FROM MOK2) END AS MOKDAY                                ");
            //   strSql.AppendLine("          , A2.DANJUNG                                                                                        ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                 ELSE ISNULL(A2.DANJUNG,0) -(SELECT ISNULL(I_WEIGHT / "+sDD+", 0) FROM MOK2) END AS CHA            ");
            //   strSql.AppendLine("          FROM DATERANGE A1                                                                                   ");
            //   strSql.AppendLine("       LEFT JOIN(SELECT A.J_DATE                                                                              ");
            //   strSql.AppendLine("                       , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                                      ");
            //   strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                                   ");
            //   strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                                                              ");
            //   strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                                                                ");
            //   strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                                                                 ");
            //   strSql.AppendLine("                     AND A.KERATYPE = '매입'                                                                  ");
            //   strSql.AppendLine("                     AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                       ");
            //   strSql.AppendLine("                     AND A.J_LOTNO <> '4'                                                                     ");
            //   strSql.AppendLine("                     AND B.DAEGUBUN = '고철A'                                                                 ");
            //   strSql.AppendLine("                   GROUP BY A.J_DATE ) A2                                                                     ");
            //   strSql.AppendLine("         ON A1.DATES = A2.J_DATE                                                                              ");
            //   strSql.AppendLine(" ), TEMP9 AS(                                                                                                 ");
            //   strSql.AppendLine("                                                                                                              ");
            //   strSql.AppendLine("     --매입 G / S모재(고철B)                                                                                  ");
            //   strSql.AppendLine("     SELECT A1.DATES                                                                                          ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                 ELSE(SELECT ISNULL(I_LIGHT / "+sDD+", 0) FROM MOK2) END AS MOKDAY                                 ");
            //   strSql.AppendLine("          , A2.DANJUNG                                                                                        ");
            //   strSql.AppendLine("          , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL                                               ");
            //   strSql.AppendLine("                 ELSE ISNULL(A2.DANJUNG,0) -(SELECT ISNULL(I_LIGHT / "+sDD+", 0) FROM MOK2) END AS CHA             ");
            //   strSql.AppendLine("          FROM DATERANGE A1                                                                                   ");
            //   strSql.AppendLine("       LEFT JOIN(SELECT A.J_DATE                                                                              ");
            //   strSql.AppendLine("                       , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                                      ");
            //   strSql.AppendLine("                    FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E                                   ");
            //   strSql.AppendLine("                   WHERE A.J_SERIAL = B.J_SERIAL                                                              ");
            //   strSql.AppendLine("                     AND A.J_ID1 = D.DEALER_CD                                                                ");
            //   strSql.AppendLine("                     AND D.CHRG_ID = E.EMP_ID                                                                 ");
            //   strSql.AppendLine("                     AND A.KERATYPE = '매입'                                                                  ");
            //   strSql.AppendLine("                     AND A.J_DATE BETWEEN '"+sDateFrom+"' AND '"+sDateYmd+"'                                       ");
            //   strSql.AppendLine("                     AND A.J_LOTNO <> '4'                                                                     ");
            //   strSql.AppendLine("                     AND B.DAEGUBUN = '고철B'                                                                 ");
            //   strSql.AppendLine("                   GROUP BY A.J_DATE) A2                                                                      ");
            //   strSql.AppendLine("         ON A1.DATES = A2.J_DATE                                                                              ");
            //   strSql.AppendLine(" ), MAIPTB AS(                                                                                                ");
            //   strSql.AppendLine("      SELECT T1.DATES                                                                                         ");
            //   strSql.AppendLine("          , T1.MOKDAY AS MOKDAY1                                                                              ");
            //   strSql.AppendLine("          , T1.DANJUNG AS W1                                                                                  ");
            //   strSql.AppendLine("          , T1.CHA AS CHA1                                                                                    ");
            //   strSql.AppendLine("          , T2.MOKDAY AS MOKDAY2                                                                              ");
            //   strSql.AppendLine("          , T2.DANJUNG AS W2                                                                                  ");
            //   strSql.AppendLine("          , T2.CHA AS CHA2                                                                                    ");
            //   strSql.AppendLine("          , T3.MOKDAY AS MOKDAY3                                                                              ");
            //   strSql.AppendLine("          , T3.DANJUNG AS W3                                                                                  ");
            //   strSql.AppendLine("          , T3.CHA AS CHA3                                                                                    ");
            //   strSql.AppendLine("       FROM TEMP7 T1                                                                                          ");
            //   strSql.AppendLine("       LEFT JOIN TEMP8 T2                                                                                     ");
            //   strSql.AppendLine("         ON T1.DATES = T2.DATES                                                                               ");
            //   strSql.AppendLine("       LEFT JOIN TEMP9 T3                                                                                     ");
            //   strSql.AppendLine("         ON T1.DATES = T3.DATES                                                                               ");
            //   strSql.AppendLine(" )                                                                                                            ");
            //   strSql.AppendLine("                                                                                                              ");
            //   strSql.AppendLine(" SELECT A1.DATES                                                                                              ");
            //   strSql.AppendLine("      , A1.MOKDAY1 AS M1                                                                                      ");
            //   strSql.AppendLine("      , A1.YK1                                                                                                ");
            //   strSql.AppendLine("      , A1.DH1                                                                                                ");
            //   strSql.AppendLine("      , A1.GITA1                                                                                              ");
            //   strSql.AppendLine("      , A1.CHA1 AS CHA1                                                                                       ");
            //   strSql.AppendLine("      , A1.MOKDAY2 AS M2                                                                                      ");
            //   strSql.AppendLine("      , A1.YK2                                                                                                ");
            //   strSql.AppendLine("      , A1.DH2                                                                                                ");
            //   strSql.AppendLine("      , A1.GITA2                                                                                              ");
            //   strSql.AppendLine("      , A1.CHA2 AS CHA2                                                                                       ");
            //   strSql.AppendLine("      , A1.MOKDAY3 AS M3                                                                                      ");
            //   strSql.AppendLine("      , A1.YK3                                                                                                ");
            //   strSql.AppendLine("      , A1.DH3                                                                                                ");
            //   strSql.AppendLine("      , A1.GITA3                                                                                              ");
            //   strSql.AppendLine("      , A1.CHA3 AS CHA3                                                                                       ");
            //   strSql.AppendLine("      , A1.ASR                                                                                                ");
            //   strSql.AppendLine("      , A2.MOKDAY1 AS M4                                                                                      ");
            //strSql.AppendLine("      , A2.W1                                                                                                 ");
            //strSql.AppendLine("      , A2.CHA1 AS CHA4                                                                                       ");
            //strSql.AppendLine("      , A2.MOKDAY2 AS M5                                                                                      ");
            //strSql.AppendLine("      , A2.W2                                                                                                 ");
            //strSql.AppendLine("      , A2.CHA2 AS CHA5                                                                                       ");
            //strSql.AppendLine("      , A2.MOKDAY3 AS M6                                                                                      ");
            //strSql.AppendLine("      , A2.W3                                                                                                 ");
            //strSql.AppendLine("      , A2.CHA3 AS CHA6                                                                                       ");
            //strSql.AppendLine("      , A3.W1 AS W4                                                                                           ");
            //strSql.AppendLine("      , A3.CNT1                                                                                               ");
            //strSql.AppendLine("      , A3.W2 AS W5                                                                                           ");
            //strSql.AppendLine("      , A3.CNT2                                                                                               ");
            //strSql.AppendLine("      , A3.TOTCP                                                                                              ");
            //strSql.AppendLine("      , A3.TOTDUST                                                                                            ");
            //   strSql.AppendLine("   FROM MACHTB A1                                                                                             ");
            //   strSql.AppendLine("   LEFT JOIN MAIPTB A2                                                                                        ");
            //   strSql.AppendLine("     ON A1.DATES = A2.DATES                                                                                   ");
            //   strSql.AppendLine("   LEFT JOIN JEGOMV A3                                                                                        ");
            //   strSql.AppendLine("     ON A1.DATES = A3.DATES                                                                                   ");

            //   DataTable dtMIOJ = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //   gridControl6.DataSource = dtMIOJ;

            //   if (dtMIOJ != null)
            //   {
            //       double dTotCP = 0;
            //       double dTotDust = 0;

            //       double.TryParse(dtMIOJ.Rows[0]["TOTCP"]?.ToString(), out dTotCP);
            //       double.TryParse(dtMIOJ.Rows[0]["TOTDUST"]?.ToString(), out dTotDust);

            //       GridBandWMV1.Caption = dTotCP.ToString("#,0 T");
            //       GridBandWMV2.Caption = dTotDust.ToString("#,0 T");
            //   }
            //   else
            //   {
            //       GridBandWMV1.Caption = "";
            //       GridBandWMV2.Caption = "";
            //   }
            #endregion

            #region 직송
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH DATERANGE(DATES) AS                                ");
            strSql.AppendLine(" (                                                       ");
            strSql.AppendLine("     SELECT CONVERT(DATE, '"+ sDateFrom + "')--시작일자        ");
            strSql.AppendLine("      UNION ALL                                          ");
            strSql.AppendLine("     SELECT DATEADD(D, 1, DATES)                         ");
            strSql.AppendLine("       FROM DATERANGE                                    ");
            strSql.AppendLine("      WHERE DATES < CONVERT(DATE, '"+ sDateYmd + "')--종료일자");
            strSql.AppendLine(" ), MOK1 AS(                                             ");
            strSql.AppendLine("     SELECT SUM(O_GS) AS O_GS --G / S                    ");
	        strSql.AppendLine("          , SUM(O_WEIGHT) AS O_WEIGHT --중량             ");
	        strSql.AppendLine("          , SUM(O_YK) AS O_YK                            ");
            strSql.AppendLine("          , SUM(O_SD) AS O_SD --슈레더                   ");
            strSql.AppendLine("       FROM SALEMAECHUL                                  ");
            strSql.AppendLine("      WHERE YYMM = '"+ sYYMM + "'                              ");
            strSql.AppendLine(" ), TEMP1 AS(                                            ");
            strSql.AppendLine("      SELECT A1.DATES                                    ");
            strSql.AppendLine("           , CASE WHEN DATEPART(WEEKDAY, A1.DATES) = 1 THEN NULL");
            strSql.AppendLine("                     ELSE (SELECT (ISNULL(O_GS,0)+ISNULL(O_WEIGHT, 0) + ISNULL(O_YK, 0) + ISNULL(O_SD, 0))/ "+ sDD + " FROM MOK1) END AS MOKDAY");
	        strSql.AppendLine("           , X1.DH                 ");
	        strSql.AppendLine("           , X1.GITA               ");
            strSql.AppendLine("        FROM DATERANGE A1          ");
            strSql.AppendLine("        LEFT JOIN(SELECT Z1.J_DATE ");
            strSql.AppendLine("                       , SUM(CASE WHEN Z1.J_ID1 = 6541754044 THEN Z1.DANJUNG END) AS DH  ");
            strSql.AppendLine("                       , SUM(CASE WHEN Z1.J_ID1 NOT IN(6541754044) THEN Z1.DANJUNG END) AS GITA");
            strSql.AppendLine("                    FROM(SELECT A.J_DATE                                            ");
            strSql.AppendLine("                              , A.J_ID1                                            ");
            strSql.AppendLine("                              , SUM(A.DANJUNG) / 1000 as DANJUNG                   ");
            strSql.AppendLine("                           FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E ");
            strSql.AppendLine("                          WHERE A.J_SERIAL = B.J_SERIAL                            ");
            strSql.AppendLine("                            AND A.J_ID1 = D.DEALER_CD                              ");
            strSql.AppendLine("                            AND D.CHRG_ID = E.EMP_ID                               ");
            strSql.AppendLine("                            AND A.KERATYPE = '매출'                                ");
            strSql.AppendLine("                            AND A.J_DATE BETWEEN '"+ sDateFrom + "' AND '"+ sDateYmd + "'     ");
            strSql.AppendLine("                            AND A.J_LOTNO = '4'                                    ");
            strSql.AppendLine("                          GROUP BY A.J_DATE, A.J_ID1)Z1                            ");
            strSql.AppendLine("                  GROUP BY Z1.J_DATE) X1                                            ");
            strSql.AppendLine("         ON A1.DATES = X1.J_DATE                                                       ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                          ");
            strSql.AppendLine("     SELECT A1.DATES                                                                   ");
            strSql.AppendLine("          , X1.DH                                                                     ");
            strSql.AppendLine("          , X1.GITA                                                                   ");
            strSql.AppendLine("       FROM DATERANGE A1                                                              ");
            strSql.AppendLine("       LEFT JOIN (SELECT Z1.J_DATE                                                    ");
            strSql.AppendLine("                       , SUM(CASE WHEN Z1.J_ID1 = 6541754044 THEN Z1.DANJUNG END) AS DH  ");
            strSql.AppendLine("                       , SUM(CASE WHEN Z1.J_ID1 NOT IN(6541754044) THEN Z1.DANJUNG END) AS GITA");
            strSql.AppendLine("                    FROM (SELECT A.J_DATE                                           ");
            strSql.AppendLine("                                , A.J_ID1                                          ");
            strSql.AppendLine("                                , SUM(A.DANJUNG) / 1000 AS DANJUNG                 ");
            strSql.AppendLine("                             FROM INLIST A, JAJAE B, ACC_DEALER_CD D, HR_EMP_BASIS E ");
            strSql.AppendLine("                            WHERE A.J_SERIAL = B.J_SERIAL                           ");
            strSql.AppendLine("                              AND A.J_ID1 = D.DEALER_CD                             ");
            strSql.AppendLine("                              AND D.CHRG_ID = E.EMP_ID                              ");
            strSql.AppendLine("                              AND A.KERATYPE = '매출'                               ");
            strSql.AppendLine("                              AND A.J_DATE BETWEEN '"+ sDateFrom + "' AND '"+ sDateYmd + "'    ");
            strSql.AppendLine("                              AND A.J_LOTNO <> '4'                                  ");
            strSql.AppendLine("                            GROUP BY A.J_DATE, A.J_ID1)Z1                           ");
            strSql.AppendLine("                   GROUP BY Z1.J_DATE) X1                                           ");
            strSql.AppendLine("        ON A1.DATES = X1.J_DATE                                                      ");
            strSql.AppendLine(" )                                                                                    ");
            strSql.AppendLine("                                                                                      ");
            strSql.AppendLine(" SELECT T1.DATES                                                                      ");
            strSql.AppendLine("      , T1.MOKDAY                                                                     ");
            strSql.AppendLine("      , T1.DH AS DH1                                                                  ");
            strSql.AppendLine("      , T1.GITA AS GITA1                                                              ");
            strSql.AppendLine("      , T2.DH AS DH2                                                                  ");
            strSql.AppendLine("      , T2.GITA AS GITA2                                                              ");
            strSql.AppendLine("      , CASE WHEN ISNULL(T1.DH, 0) + ISNULL(T1.GITA, 0) = 0 THEN NULL                           ");
            strSql.AppendLine("             ELSE ISNULL(T1.DH, 0) + ISNULL(T1.GITA, 0) END AS TOTJN                              ");
            strSql.AppendLine("      , CASE WHEN ISNULL(T2.DH, 0) + ISNULL(T2.GITA, 0) = 0 THEN NULL                           ");
            strSql.AppendLine("             ELSE ISNULL(T2.DH, 0) + ISNULL(T2.GITA, 0) END AS TOTYA                              ");
            strSql.AppendLine("      , CASE WHEN ISNULL(T1.DH, 0) + ISNULL(T1.GITA, 0)                                         ");
            strSql.AppendLine("                  + ISNULL(T2.DH, 0) + ISNULL(T2.GITA, 0) = 0 THEN NULL                         ");
            strSql.AppendLine("             ELSE ISNULL(T1.DH, 0) + ISNULL(T1.GITA, 0)                                           ");
            strSql.AppendLine("                  + ISNULL(T2.DH, 0) + ISNULL(T2.GITA, 0) END AS TOTSUM                         ");
            strSql.AppendLine("      , CASE WHEN (ISNULL(T1.DH, 0) + ISNULL(T1.GITA, 0)                                          ");
            strSql.AppendLine("                  + ISNULL(T2.DH, 0) + ISNULL(T2.GITA, 0)) -ISNULL(T1.MOKDAY, 0) = 0 THEN NULL");
            strSql.AppendLine("             ELSE (ISNULL(T1.DH, 0) + ISNULL(T1.GITA, 0)                                        ");
            strSql.AppendLine("                  + ISNULL(T2.DH, 0) + ISNULL(T2.GITA, 0)) - ISNULL(T1.MOKDAY, 0) END AS CHA    ");
            strSql.AppendLine("   FROM TEMP1 T1            ");                                                                                    
            strSql.AppendLine("   LEFT JOIN TEMP2 T2       ");
            strSql.AppendLine("     ON T1.DATES = T2.DATES ");

            DataTable dtRetr4 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridControl4.DataSource = dtRetr4;


            #endregion
            Cursor = Cursors.Default;
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            FileInfo_1 fileInfo = new FileInfo_1("5");

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

                string sDate = DateYMD.EditValue?.ToString().Substring(0, 10);

                ExcelApp = new Excel.Application();
                wb = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                ws = wb.Worksheets.get_Item(1);

                //일자부분 세팅
                ws.Range["B1"].Value = sDate;
                //월 세팅
                ws.Range["B4"].Value = sDate.Substring(5, 2);
                //근무일수
                ws.Range["G1"].Value = tE_WorkDay.Text;

                //목표량 설정
                string sMockC1 = GridBandC.Caption;
                string sMockW = GridBandW.Caption;
                string sMockG = GridBandG.Caption;
                string totMechul = GridBandTotMeChul.Caption;

                //매출
                ws.Range["C5"].Value = sMockC1;
                ws.Range["G5"].Value = sMockW;
                ws.Range["K5"].Value = sMockG;
                ws.Range["O5"].Value = totMechul;

                // 매출
                DataTable dtMaChul = (DataTable)gridControl1.DataSource;

                if (dtMaChul != null)
                {
                    int iStartRow = 7;

                    for (int i = 0; i < dtMaChul.Rows.Count; i++)
                    {
                        string sDATES = dtMaChul.Rows[i]["DATES"]?.ToString().Substring(0, 10);
                        string sMOKDAY1 = dtMaChul.Rows[i]["MOKDAY1"]?.ToString();
                        string sYK1 = dtMaChul.Rows[i]["YK1"]?.ToString();
                        string sGITA1 = dtMaChul.Rows[i]["GITA1"]?.ToString();
                        string sCHA1 = dtMaChul.Rows[i]["CHA1"]?.ToString();
                        string sMOKDAY2 = dtMaChul.Rows[i]["MOKDAY2"]?.ToString();
                        string sYK2 = dtMaChul.Rows[i]["YK2"]?.ToString();
                        string sGITA2 = dtMaChul.Rows[i]["GITA2"]?.ToString();
                        string sCHA2 = dtMaChul.Rows[i]["CHA2"]?.ToString();
                        string sMOKDAY3 = dtMaChul.Rows[i]["MOKDAY3"]?.ToString();
                        string sYK3 = dtMaChul.Rows[i]["YK3"]?.ToString();
                        string sGITA3 = dtMaChul.Rows[i]["GITA3"]?.ToString();
                        string sCHA3 = dtMaChul.Rows[i]["CHA3"]?.ToString();
                        string sTMOK = dtMaChul.Rows[i]["TMOK"]?.ToString();
                        string sTYK = dtMaChul.Rows[i]["TYK"]?.ToString();
                        string sTGITA = dtMaChul.Rows[i]["TGITA"]?.ToString();
                        string sTCHA = dtMaChul.Rows[i]["TCHA"]?.ToString();
                        string sASR = dtMaChul.Rows[i]["ASR"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["B" + iApplyRowIdx].Value = sDATES;
                        ws.Range["C" + iApplyRowIdx].Value = sMOKDAY1;
                        ws.Range["D" + iApplyRowIdx].Value = sYK1;
                        ws.Range["E" + iApplyRowIdx].Value = sGITA1;
                        ws.Range["F" + iApplyRowIdx].Value = sCHA1;
                        ws.Range["G" + iApplyRowIdx].Value = sMOKDAY2;
                        ws.Range["H" + iApplyRowIdx].Value = sYK2;
                        ws.Range["I" + iApplyRowIdx].Value = sGITA2;
                        ws.Range["J" + iApplyRowIdx].Value = sCHA2;
                        ws.Range["K" + iApplyRowIdx].Value = sMOKDAY3;
                        ws.Range["L" + iApplyRowIdx].Value = sYK3;
                        ws.Range["M" + iApplyRowIdx].Value = sGITA3;
                        ws.Range["N" + iApplyRowIdx].Value = sCHA3;
                        ws.Range["O" + iApplyRowIdx].Value = sTMOK;
                        ws.Range["P" + iApplyRowIdx].Value = sTYK;
                        ws.Range["Q" + iApplyRowIdx].Value = sTGITA;
                        ws.Range["R" + iApplyRowIdx].Value = sTCHA;
                        ws.Range["S" + iApplyRowIdx].Value = sASR;
                    }
                }

                // 매입
                ws = wb.Worksheets.get_Item(2);

                //일자부분 세팅
                ws.Range["B1"].Value = sDate;
                //월 세팅
                ws.Range["B4"].Value = sDate.Substring(5, 2);
                //근무일수
                ws.Range["H1"].Value = tE_WorkDay.Text;

                string sMockC2 = GridBandPe.Caption;
                string sMockD = GridBandWe.Caption;
                string sMockP = GridBandGs.Caption;
                string sTotMeip = GridBandTotMeip.Caption;

                //매입
                ws.Range["C5"].Value = sMockC2;
                ws.Range["F5"].Value = sMockD;
                ws.Range["I5"].Value = sMockP;
                ws.Range["L5"].Value = sTotMeip;

                DataTable dtMaIP = (DataTable)gridControl5.DataSource;

                if (dtMaIP != null)
                {
                    int iStartRow = 7;

                    for (int i = 0; i < dtMaIP.Rows.Count; i++)
                    {
                        string sDATES = dtMaIP.Rows[i]["DATES"]?.ToString().Substring(0, 10);
                        string sMOKDAY1 = dtMaIP.Rows[i]["MOKDAY1"]?.ToString();
                        string sW1 = dtMaIP.Rows[i]["W1"]?.ToString();
                        string sCHA1 = dtMaIP.Rows[i]["CHA1"]?.ToString();
                        string sMOKDAY2 = dtMaIP.Rows[i]["MOKDAY2"]?.ToString();
                        string sW2 = dtMaIP.Rows[i]["W2"]?.ToString();
                        string sCHA2 = dtMaIP.Rows[i]["CHA2"]?.ToString();
                        string sMOKDAY3 = dtMaIP.Rows[i]["MOKDAY3"]?.ToString();
                        string sW3 = dtMaIP.Rows[i]["W3"]?.ToString();
                        string sCHA3 = dtMaIP.Rows[i]["CHA3"]?.ToString();
                        string sTMOK = dtMaIP.Rows[i]["TMOK"]?.ToString();
                        string sTW = dtMaIP.Rows[i]["TW"]?.ToString();
                        string sTCHA = dtMaIP.Rows[i]["TCHA"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["B" + iApplyRowIdx].Value = sDATES;
                        ws.Range["C" + iApplyRowIdx].Value = sMOKDAY1;
                        ws.Range["D" + iApplyRowIdx].Value = sW1;
                        ws.Range["E" + iApplyRowIdx].Value = sCHA1;
                        ws.Range["F" + iApplyRowIdx].Value = sMOKDAY2;
                        ws.Range["G" + iApplyRowIdx].Value = sW2;
                        ws.Range["H" + iApplyRowIdx].Value = sCHA2;
                        ws.Range["I" + iApplyRowIdx].Value = sMOKDAY3;
                        ws.Range["J" + iApplyRowIdx].Value = sW3;
                        ws.Range["K" + iApplyRowIdx].Value = sCHA3;
                        ws.Range["L" + iApplyRowIdx].Value = sTMOK;
                        ws.Range["M" + iApplyRowIdx].Value = sTW;
                        ws.Range["N" + iApplyRowIdx].Value = sTCHA;

                    }
                }

                //재고
                string sMockW2 = GridBandC2.Caption;
                string sMockG2 = GridBandD.Caption;

                ws.Range["Q5"].Value = sMockC2;
                ws.Range["S5"].Value = sMockD;

                DataTable dtJaego = (DataTable)gridControl3.DataSource;

                if (dtJaego != null)
                {
                    int iStartRow = 7;

                    for (int i = 0; i < dtJaego.Rows.Count; i++)
                    {
                        string sDATES = dtJaego.Rows[i]["DATES"]?.ToString().Substring(0, 10);
                        string sW1 = dtJaego.Rows[i]["W1"]?.ToString();
                        //string sCNT1 = dtJaego.Rows[i]["CNT1"]?.ToString();
                        //string sW2 = dtJaego.Rows[i]["W2"]?.ToString();
                        //string sCNT2 = dtJaego.Rows[i]["CNT2"]?.ToString();
                        //string sTOTCP = dtJaego.Rows[i]["TOTCP"]?.ToString();
                        //string sTOTDUST = dtJaego.Rows[i]["TOTDUST"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["P" + iApplyRowIdx].Value = sDATES  ;
                        ws.Range["Q" + iApplyRowIdx].Value = sW1     ;
                        //ws.Range["R" + iApplyRowIdx].Value = sCNT1   ;
                        //ws.Range["S" + iApplyRowIdx].Value = sW2     ;
                        //ws.Range["T" + iApplyRowIdx].Value = sCNT2   ;
                    }
                }

                #region 2022-06-20 이전
                //DataTable dt = (DataTable)gridControl6.DataSource;

                //if(dt != null)
                //{
                //    int iStartRow = 8;

                //    for(int i = 0; i < dt.Rows.Count; i++)
                //    {
                //        string sDATES   = dt.Rows[i]["DATES"]?.ToString().Substring(0,10);
                //        string sM1      = dt.Rows[i]["M1"]?.ToString();
                //        string sYK1     = dt.Rows[i]["YK1"]?.ToString();
                //        string sDH1     = dt.Rows[i]["DH1"]?.ToString();
                //        string sGITA1   = dt.Rows[i]["GITA1"]?.ToString();
                //        string sCHA1    = dt.Rows[i]["CHA1"]?.ToString();
                //        string sM2      = dt.Rows[i]["M2"]?.ToString();
                //        string sYK2     = dt.Rows[i]["YK2"]?.ToString();
                //        string sDH2     = dt.Rows[i]["DH2"]?.ToString();
                //        string sGITA2   = dt.Rows[i]["GITA2"]?.ToString();
                //        string sCHA2    = dt.Rows[i]["CHA2"]?.ToString();
                //        string sM3      = dt.Rows[i]["M3"]?.ToString();
                //        string sYK3     = dt.Rows[i]["YK3"]?.ToString();
                //        string sDH3     = dt.Rows[i]["DH3"]?.ToString();
                //        string sGITA3   = dt.Rows[i]["GITA3"]?.ToString();
                //        string sCHA3    = dt.Rows[i]["CHA3"]?.ToString();
                //        string sASR     = dt.Rows[i]["ASR"]?.ToString();
                //        string sM4      = dt.Rows[i]["M4"]?.ToString();
                //     string sW1      = dt.Rows[i]["W1"]?.ToString();
                //     string sCHA4    = dt.Rows[i]["CHA4"]?.ToString();
                //     string sM5      = dt.Rows[i]["M5"]?.ToString();
                //     string sW2      = dt.Rows[i]["W2"]?.ToString();
                //     string sCHA5    = dt.Rows[i]["CHA5"]?.ToString();
                //     string sM6      = dt.Rows[i]["M6"]?.ToString();
                //     string sW3      = dt.Rows[i]["W3"]?.ToString();
                //     string sCHA6    = dt.Rows[i]["CHA6"]?.ToString();
                //     string sW4      = dt.Rows[i]["W4"]?.ToString();
                //     string sCNT1    = dt.Rows[i]["CNT1"]?.ToString();
                //     string sW5      = dt.Rows[i]["W5"]?.ToString();
                //     string sCNT2    = dt.Rows[i]["CNT2"]?.ToString();

                //        int iApplyRowIdx = iStartRow + (i + 1);

                //        ws.Range["B" + iApplyRowIdx].Value = sDATES;
                //        ws.Range["C" + iApplyRowIdx].Value = sM1;
                //        ws.Range["D" + iApplyRowIdx].Value = sYK1;
                //        ws.Range["E" + iApplyRowIdx].Value = sDH1;
                //        ws.Range["F" + iApplyRowIdx].Value = sGITA1;
                //        ws.Range["G" + iApplyRowIdx].Value = sCHA1;
                //        ws.Range["H" + iApplyRowIdx].Value = sM2;
                //        ws.Range["I" + iApplyRowIdx].Value = sYK2;
                //        ws.Range["J" + iApplyRowIdx].Value = sDH2;
                //        ws.Range["K" + iApplyRowIdx].Value = sGITA2;
                //        ws.Range["L" + iApplyRowIdx].Value = sCHA2;
                //        ws.Range["M" + iApplyRowIdx].Value = sM3;
                //        ws.Range["N" + iApplyRowIdx].Value = sYK3;
                //        ws.Range["O" + iApplyRowIdx].Value = sDH3;
                //        ws.Range["P" + iApplyRowIdx].Value = sGITA3;
                //        ws.Range["Q" + iApplyRowIdx].Value = sCHA3;
                //        ws.Range["R" + iApplyRowIdx].Value = sASR;


                //        ws.Range["T" + iApplyRowIdx].Value = sDATES;
                //        ws.Range["U" + iApplyRowIdx].Value = sM4;
                //        ws.Range["V" + iApplyRowIdx].Value = sW1;
                //        ws.Range["W" + iApplyRowIdx].Value = sCHA4;
                //        ws.Range["X" + iApplyRowIdx].Value = sM5;
                //        ws.Range["Y" + iApplyRowIdx].Value = sW2;
                //        ws.Range["Z" + iApplyRowIdx].Value = sCHA5;
                //        ws.Range["AA" + iApplyRowIdx].Value = sM6;
                //        ws.Range["AB" + iApplyRowIdx].Value = sW3;
                //        ws.Range["AC" + iApplyRowIdx].Value = sCHA6;


                //        ws.Range["AE" + iApplyRowIdx].Value = sDATES;
                //        ws.Range["AF" + iApplyRowIdx].Value = sW4;
                //        ws.Range["AG" + iApplyRowIdx].Value = sCNT1;
                //        ws.Range["AH" + iApplyRowIdx].Value = sW5;
                //        ws.Range["AI" + iApplyRowIdx].Value = sCNT2;
                //    }
                //}
                #endregion


                //직송
                Excel.Worksheet ws2 = wb.Worksheets.get_Item(3);

                //월 세팅
                ws2.Range["A4"].Value = sDate.Substring(5, 2);
                //목표량 세팅
                string sMockJK = GridBandJK.Caption;
                ws2.Range["B5"].Value = sMockJK;

                DataTable dtJK = (DataTable)gridControl4.DataSource;

                if (dtJK != null)
                {
                    int iStartRow = 7;

                    for (int i = 0; i < dtJK.Rows.Count; i++)
                    {
                        string sDATES = dtJK.Rows[i]["DATES"]?.ToString().Substring(0, 10);
                        string sMOKDAY = dtJK.Rows[i]["MOKDAY"]?.ToString();
                        string sDH1 = dtJK.Rows[i]["DH1"]?.ToString();
                        string sGITA1 = dtJK.Rows[i]["GITA1"]?.ToString();
                        string sTOTJN = dtJK.Rows[i]["TOTJN"]?.ToString();
                        string sDH2 = dtJK.Rows[i]["DH2"]?.ToString();
                        string sGITA2 = dtJK.Rows[i]["GITA2"]?.ToString();
                        string sTOTYA = dtJK.Rows[i]["TOTYA"]?.ToString();
                        string sTOTSUM = dtJK.Rows[i]["TOTSUM"]?.ToString();
                        //string sCHA = dtJK.Rows[i]["CHA"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws2.Range["A" + iApplyRowIdx].Value = sDATES;
                        ws2.Range["B" + iApplyRowIdx].Value = sMOKDAY;
                        ws2.Range["C" + iApplyRowIdx].Value = sDH1;
                        ws2.Range["D" + iApplyRowIdx].Value = sGITA1;
                        ws2.Range["E" + iApplyRowIdx].Value = sTOTJN;
                        ws2.Range["F" + iApplyRowIdx].Value = sDH2;
                        ws2.Range["G" + iApplyRowIdx].Value = sGITA2;
                        ws2.Range["H" + iApplyRowIdx].Value = sTOTYA;
                        ws2.Range["I" + iApplyRowIdx].Value = sTOTSUM;
                        //ws2.Range["L" + iApplyRowIdx].Value = sCHA;
                    }
                }


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

        private void IN12001F01_KeyDown(object sender, KeyEventArgs e)
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

        private void BGridViewRetr1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void BGridViewRetr6_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            if(e.Column.FieldName == "ASR" || e.Column.FieldName == "CHA6")
            {
                ComnGridFunc.CellDrawHelper.DoDefaultDrawCell(BGridViewRetr6, e);
                ComnGridFunc.CellDrawHelper.DrawCellBorder(e);
                e.Handled = true;
            }
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateYMD.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateYMD.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateYMD.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateYMD.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }

        private void IN12001F01_TextChanged(object sender, EventArgs e)
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

        private void BGridViewRetr4_CustomDrawFooterCell(object sender, FooterCellCustomDrawEventArgs e)
        {
            if(string.Compare(e.Column.FieldName, "TOTJN") == 0)
            {
                e.Appearance.BackColor = Color.Yellow;
                e.Appearance.Font = new Font(DefaultFont, FontStyle.Bold);
                                
            }
            if (string.Compare(e.Column.FieldName, "TOTYA") == 0)
            {
                e.Appearance.BackColor = Color.Yellow;
                e.Appearance.Font = new Font(DefaultFont, FontStyle.Bold);
                
            }
            if (string.Compare(e.Column.FieldName, "TOTSUM") == 0)
            {
                e.Appearance.BackColor2 = Color.Yellow;
                e.Appearance.Font = new Font(DefaultFont, FontStyle.Bold);
                
            }
            e.Painter.DrawObject(e.Info);
            e.Handled = true;
            
        }
    }
}

