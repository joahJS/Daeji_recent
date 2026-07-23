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
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
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
    public partial class AC10001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC10001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC10001F01_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            DateEditYmd.EditValue = today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { advBandedGridView1 };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }


            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmd = DateEditYmd.EditValue?.ToString().Substring(0, 10);
            GridBal.DataSource = GetBalancesSummaryList(sYmd);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ComnEtcFunc.ExportExcelFile(this.Text, GridBal);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private DataTable GetBalancesSummaryList(string sYmd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("SELECT ISNULL(JCAMT,0) AS C_BAL_AMT, ISNULL(ACAMT, 0) AS C_TOT_AMT, ACCOD, SEQNO, CONCAT(PRTSN, ACDSP) AS CVNAM, ISNULL(ADAMT, 0) AS D_TOT_AMT, ISNULL(JDAMT, 0) AS D_BAL_AMT ");
            strSql.AppendLine(" FROM(                                                                                                                                                                        ");
            strSql.AppendLine("    SELECT A1.ACCOD, A1.SEQNO, ISNULL(A1.PRTSN, '')PRTSN, A1.ACDSP, SUM(A2.ACAMT)ACAMT, SUM(A2.ADAMT)ADAMT, SUM(JCAMT)JCAMT, SUM(JDAMT)JDAMT                                  ");
            strSql.AppendLine("    FROM ACTOPF A1                                                                                                                                                            ");
            strSql.AppendLine("         LEFT JOIN                                                                                                                                                            ");
            strSql.AppendLine("         (SELECT CONCAT(A1.ACCOD, 'X')ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT) ACAMT, SUM(ADAMT) ADAMT                                                                           ");
            strSql.AppendLine("               , SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE 0 END) JCAMT                                                                      ");
            strSql.AppendLine("               , SUM(CASE WHEN B1.ACRDR = '2' THEN ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) ELSE 0 END) JDAMT                                                                      ");
            strSql.AppendLine("          FROM   ACTRAN A1                                                                                                                                                    ");
            strSql.AppendLine("          LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                                                                    ");
            strSql.AppendLine("          WHERE A1.TDATE BETWEEN SUBSTRING('" + sYmd + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), CONVERT(DATE, '" + sYmd + "'), 112)                                             ");
            strSql.AppendLine("            AND A1.APVYN = 'Y'                                                                                                        ");
            strSql.AppendLine("          GROUP  BY A1.ACCOD, B1.ACDSP, B1.ACRDR                                                                                      ");
            strSql.AppendLine("          UNION  ALL                                                                                                                  ");
            strSql.AppendLine("          SELECT CONCAT(A1.ACCOD, 'X'), B1.ACDSP, B1.ACRDR, 0, 0, SUM(ACDRJN), SUM(ACCRJN)                                            ");
            strSql.AppendLine("          FROM   ACJANF A1                                                                                                            ");
            strSql.AppendLine("          LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                            ");
            strSql.AppendLine("          WHERE  A1.ACYEAR = SUBSTRING('" + sYmd + "', 1, 4)                                                                           ");
            strSql.AppendLine("          GROUP  BY A1.ACCOD, B1.ACDSP, B1.ACRDR                                                                                      ");
            strSql.AppendLine("         ) A2 ON A2.ACCOD BETWEEN A1.AFROM AND A1.ATO                                                                                 ");
            strSql.AppendLine("    GROUP BY A1.ACCOD, A1.SEQNO, A1.PRTSN, A1.ACDSP                                                                                   ");
            strSql.AppendLine("   UNION ALL                                                                                                                          ");
            strSql.AppendLine("   SELECT A1.ACCOD, 99, ' ', A1.ACDSP, SUM(A1.ACAMT), SUM(A1.ADAMT), SUM(A1.JCAMT), SUM(JDAMT)                                        ");
            strSql.AppendLine("   FROM(                                                                                                                              ");
            strSql.AppendLine("         SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT)ACAMT, SUM(ADAMT)ADAMT                                                        ");
            strSql.AppendLine("              , SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE 0 END) JCAMT                               ");
            strSql.AppendLine("              , SUM(CASE WHEN B1.ACRDR = '2' THEN ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) ELSE 0 END) JDAMT                               ");
            strSql.AppendLine("         FROM   ACTRAN A1                                                                                                             ");
            strSql.AppendLine("                LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                      ");
            strSql.AppendLine("         WHERE  A1.TDATE BETWEEN SUBSTRING('" + sYmd + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), CONVERT(DATE, '" + sYmd + "'), 112) ");
            strSql.AppendLine("           AND A1.APVYN = 'Y'                                                                                                         ");
            strSql.AppendLine("         GROUP  BY A1.ACCOD, B1.ACDSP, B1.ACRDR                                 ");
            strSql.AppendLine("         UNION  ALL                                                             ");
            strSql.AppendLine("         SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, 0, 0, SUM(ACDRJN), SUM(ACCRJN)    ");
            strSql.AppendLine("         FROM   ACJANF A1                                                       ");
            strSql.AppendLine("         LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                       ");
            strSql.AppendLine("         WHERE  A1.ACYEAR = SUBSTRING('" + sYmd + "', 1, 4)                     ");
            strSql.AppendLine("         GROUP  BY A1.ACCOD, B1.ACDSP, B1.ACRDR                                 ");
            strSql.AppendLine("         )A1                                                                    ");
            strSql.AppendLine("   GROUP  BY A1.ACCOD, A1.ACDSP                                                 ");
            strSql.AppendLine("   )A1                                                                          ");
            strSql.AppendLine("   ORDER BY ACCOD, SEQNO                                                        ");

            #region mariaDB
            //strSql.AppendLine(" SELECT IFNULL(JCAMT,0) AS C_BAL_AMT, IFNULL(ACAMT,0) AS C_TOT_AMT, ACCOD, SEQNO, CONCAT(PRTSN, ACDSP) AS CVNAM, IFNULL(ADAMT,0) AS D_TOT_AMT, IFNULL(JDAMT,0) AS D_BAL_AMT ");
            //strSql.AppendLine(" FROM ( ");
            //strSql.AppendLine("    SELECT A1.ACCOD, A1.SEQNO, IFNULL(A1.PRTSN,'')PRTSN, A1.ACDSP, SUM(A2.ACAMT)ACAMT, SUM(A2.ADAMT)ADAMT, SUM(JCAMT)JCAMT, SUM(JDAMT)JDAMT ");
            //strSql.AppendLine("    FROM ACTOPF A1  ");
            //strSql.AppendLine("         LEFT JOIN  ");
            //strSql.AppendLine("         (SELECT CONCAT(A1.ACCOD,'X')ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT) ACAMT, SUM(ADAMT) ADAMT ");
            //strSql.AppendLine("               , SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT,0)-IFNULL(ADAMT,0) ELSE 0 END) JCAMT ");
            //strSql.AppendLine("               , SUM(CASE WHEN B1.ACRDR='2' THEN IFNULL(ADAMT,0)-IFNULL(ACAMT,0) ELSE 0 END) JDAMT ");
            //strSql.AppendLine("          FROM   ACTRAN A1                                               ");
            //strSql.AppendLine("          LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD               ");
            //strSql.AppendLine("          WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmd + "', '%Y0101') AND DATE_FORMAT('" + sYmd + "', '%Y%m%d')                      ");
            //strSql.AppendLine("            AND A1.APVYN = 'Y' ");
            //strSql.AppendLine("          GROUP  BY A1.ACCOD, B1.ACDSP, B1.ACRDR ");
            //strSql.AppendLine("          UNION  ALL                                                      ");
            //strSql.AppendLine("          SELECT CONCAT(A1.ACCOD,'X'), B1.ACDSP, B1.ACRDR, 0,0, SUM(ACDRJN), SUM(ACCRJN) ");
            //strSql.AppendLine("          FROM   ACJANF A1                                               ");
            //strSql.AppendLine("          LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD              ");
            //strSql.AppendLine("          WHERE  A1.ACYEAR = DATE_FORMAT('" + sYmd + "', '%Y')                ");
            //strSql.AppendLine("          GROUP  BY A1.ACCOD , B1.ACDSP , B1.ACRDR  ");
            //strSql.AppendLine("         ) A2 ON A2.ACCOD BETWEEN A1.AFROM AND A1.ATO  ");
            //strSql.AppendLine("    GROUP BY A1.ACCOD, A1.SEQNO, A1.PRTSN, A1.ACDSP    ");
            //strSql.AppendLine("   UNION ALL  ");
            //strSql.AppendLine("   SELECT A1.ACCOD, 99, ' ', A1.ACDSP, SUM(A1.ACAMT), SUM(A1.ADAMT), SUM(A1.JCAMT), SUM(JDAMT) ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("         SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT)ACAMT, SUM(ADAMT)ADAMT ");
            //strSql.AppendLine("              , SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT,0)-IFNULL(ADAMT,0) ELSE 0 END) JCAMT ");
            //strSql.AppendLine("              , SUM(CASE WHEN B1.ACRDR='2' THEN IFNULL(ADAMT,0)-IFNULL(ACAMT,0) ELSE 0 END) JDAMT ");
            //strSql.AppendLine("         FROM   ACTRAN A1                                               ");
            //strSql.AppendLine("                LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD               ");
            //strSql.AppendLine("         WHERE  A1.TDATE BETWEEN DATE_FORMAT('" + sYmd + "', '%Y0101') AND DATE_FORMAT('" + sYmd + "', '%Y%m%d')  ");
            //strSql.AppendLine("           AND A1.APVYN = 'Y' ");
            //strSql.AppendLine("         GROUP  BY A1.ACCOD, B1.ACDSP, B1.ACRDR ");
            //strSql.AppendLine("         UNION  ALL                                                      ");
            //strSql.AppendLine("         SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, 0,0, SUM(ACDRJN), SUM(ACCRJN) ");
            //strSql.AppendLine("         FROM   ACJANF A1                                               ");
            //strSql.AppendLine("         LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD              ");
            //strSql.AppendLine("         WHERE  A1.ACYEAR = DATE_FORMAT('" + sYmd + "', '%Y')     ");
            //strSql.AppendLine("         GROUP  BY A1.ACCOD , B1.ACDSP , B1.ACRDR  ");
            //strSql.AppendLine("         )A1  ");
            //strSql.AppendLine("   GROUP  BY A1.ACCOD, A1.ACDSP  ");
            //strSql.AppendLine("   )A1 ");
            //strSql.AppendLine("   ORDER BY ACCOD,  SEQNO ");
            #endregion

            #region 이전코드
            //strSql.Clear();
            //strSql.AppendLine(" SELECT IFNULL(JCAMT,0) AS C_BAL_AMT, IFNULL(ACAMT,0) AS C_TOT_AMT, ACCOD, SEQNO, CONCAT(PRTSN, ACDSP) AS CVNAM, IFNULL(ADAMT,0) AS D_TOT_AMT, IFNULL(JDAMT,0) AS D_BAL_AMT ");
            //strSql.AppendLine("   FROM (  ");
            //strSql.AppendLine("          SELECT A1.ACCOD, A1.SEQNO, IFNULL(A1.PRTSN,'')PRTSN, A1.ACDSP, SUM(A2.ACAMT)ACAMT, SUM(A2.ADAMT)ADAMT, SUM(JCAMT)JCAMT, SUM(JDAMT)JDAMT ");
            //strSql.AppendLine("            FROM ACTOPF A1  ");
            //strSql.AppendLine("            LEFT JOIN ( SELECT CONCAT(A1.ACCOD,'X')ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT) ACAMT, SUM(ADAMT) ADAMT, SUM(ACAMT)JCAMT, SUM(ADAMT)JDAMT ");
            //strSql.AppendLine("                          FROM ACTRAN A1 ");
            //strSql.AppendLine("                          LEFT OUTER JOIN ACMSTF B1  ");
            //strSql.AppendLine("                            ON A1.ACCOD=B1.ACCOD  ");
            //strSql.AppendLine("                         WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmd + "', '%Y0101') AND DATE_FORMAT('" + sYmd + "', '%Y%m%d') AND A1.APVYN = 'Y' ");
            //strSql.AppendLine("                         GROUP BY A1.ACCOD, B1.ACDSP, B1.ACRDR ");
            //strSql.AppendLine("                         UNION ALL  ");
            //strSql.AppendLine("                        SELECT CONCAT(A1.ACCOD,'X'), B1.ACDSP, B1.ACRDR, 0,0,SUM(ACCRJN), SUM(ACDRJN) ");
            //strSql.AppendLine("                          FROM ACJANF A1  ");
            //strSql.AppendLine("                          LEFT OUTER JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD  ");
            //strSql.AppendLine("                         WHERE  A1.ACYEAR = DATE_FORMAT('" + sYmd + "', '%Y')  ");
            //strSql.AppendLine("                         GROUP  BY A1.ACCOD , B1.ACDSP , B1.ACRDR  ");
            //strSql.AppendLine("                      ) A2  ");
            //strSql.AppendLine("              ON A2.ACCOD BETWEEN A1.AFROM AND A1.ATO  ");
            //strSql.AppendLine("           GROUP BY A1.ACCOD, A1.SEQNO, A1.PRTSN, A1.ACDSP  ");
            //strSql.AppendLine("           UNION ALL ");
            //strSql.AppendLine("          SELECT A1.ACCOD, 99, ' ', A1.ACDSP, SUM(A1.ACAMT), SUM(A1.ADAMT), SUM(A1.JCAMT), SUM(JDAMT) ");
            //strSql.AppendLine("            FROM ( ");
            //strSql.AppendLine("                   SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT)ACAMT, SUM(ADAMT)ADAMT, SUM(ACAMT) JCAMT, SUM(ADAMT) JDAMT ");
            //strSql.AppendLine("                     FROM ACTRAN A1 ");
            //strSql.AppendLine("                     LEFT OUTER JOIN ACMSTF B1  ");
            //strSql.AppendLine("                       ON A1.ACCOD=B1.ACCOD ");
            //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmd + "', '%Y0101') AND DATE_FORMAT('" + sYmd + "', '%Y%m%d') AND A1.APVYN = 'Y' ");
            //strSql.AppendLine("                    GROUP BY A1.ACCOD, B1.ACDSP, B1.ACRDR ");
            //strSql.AppendLine("                    UNION ALL ");
            //strSql.AppendLine("                   SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, 0,0, SUM(ACCRJN), SUM(ACDRJN) ");
            //strSql.AppendLine("                     FROM ACJANF A1 ");
            //strSql.AppendLine("                     LEFT OUTER JOIN ACMSTF B1 ");
            //strSql.AppendLine("                       ON A1.ACCOD=B1.ACCOD              ");
            //strSql.AppendLine("                    WHERE A1.ACYEAR = DATE_FORMAT('" + sYmd + "', '%Y')     ");
            //strSql.AppendLine("                    GROUP BY A1.ACCOD , B1.ACDSP , B1.ACRDR  ");
            //strSql.AppendLine("                 ) A1  ");
            //strSql.AppendLine("           GROUP BY A1.ACCOD, A1.ACDSP ");
            //strSql.AppendLine("        ) A1 ");
            //strSql.AppendLine("    ORDER BY ACCOD, SEQNO ");

            //strSql.AppendLine(" SELECT CASE WHEN Z1.ACRDR = '1' THEN (Z1.C_AMT - Z1.D_AMT) ELSE 0 END AS C_BAL_AMT ");
            //strSql.AppendLine("      , CASE WHEN Z1.ACRDR = '1' THEN (Z1.C_AMT) ELSE 0 END AS C_TOT ");
            //strSql.AppendLine("      , Z1.ACDSP AS ACNAM ");
            //strSql.AppendLine("      , CASE WHEN Z1.ACRDR = '2' THEN (Z1.D_AMT) ELSE 0 END AS D_TOT ");
            //strSql.AppendLine("      , CASE WHEN Z1.ACRDR = '2' THEN (Z1.D_AMT - Z1.C_AMT) ELSE 0 END AS C_BAL_AMT ");
            //strSql.AppendLine("   FROM( ");
            //strSql.AppendLine(" SELECT X1.ACCOD ");
            //strSql.AppendLine("      , X1.ACDSP ");
            //strSql.AppendLine("      , X1.ACRDR ");
            //strSql.AppendLine(" 	 , SUM(IFNULL(X1.ACAMT, 0)) AS C_AMT ");
            //strSql.AppendLine(" 	 , SUM(IFNULL(X1.ADAMT, 0)) AS D_AMT ");
            //strSql.AppendLine("   FROM( ");
            //strSql.AppendLine("         SELECT (B1.ACCOD + B1.SEQNO) AS ACCOD ");
            //strSql.AppendLine("         	  , B1.ACDSP ");
            //strSql.AppendLine("         	  , B1.ACRDR ");
            //strSql.AppendLine("         	  , SUM(B2.ACAMT) AS ACAMT ");
            //strSql.AppendLine("         	  , SUM(B2.ADAMT) AS ADAMT ");
            //strSql.AppendLine("           FROM ACTOPF B1 ");
            //strSql.AppendLine("           LEFT OUTER JOIN( ");
            //strSql.AppendLine("                            SELECT A1.TDATE, A1.ACCOD, A1.ACAMT, A1.ADAMT ");
            //strSql.AppendLine("                              FROM ACTRAN A1  ");
            //strSql.AppendLine("                             UNION ALL  ");
            //strSql.AppendLine("                            SELECT A2.TDATE, '" + ComnEtcFunc.CashCode + "', A2.ADAMT, A2.ACAMT  ");
            //strSql.AppendLine("                              FROM ACTRAN A2 ");
            //strSql.AppendLine("                              LEFT OUTER JOIN ACMSTF A4");
            //strSql.AppendLine("                                ON A2.ACCOD = A4.ACCOD");
            //strSql.AppendLine("                             WHERE A4.ASMCD='" + ComnEtcFunc.CashCode + "' ");
            //strSql.AppendLine("                             UNION ALL ");
            //strSql.AppendLine("                            SELECT '" + sYmd + "', A3.ACCOD, A3.ACCRJN, A3.ACDRJN ");
            //strSql.AppendLine("                              FROM ACJANF A3 ");
            //strSql.AppendLine("                             WHERE A3.ACYEAR= '" + sYmd.Substring(0, 4) + "' ");
            //strSql.AppendLine("                          ) B2 ");
            //strSql.AppendLine("             ON B2.ACCOD BETWEEN B1.AFROM AND B1.ATO ");
            //strSql.AppendLine("            AND B2.TDATE BETWEEN '" + sYmd + "' AND '" + sYmd + "' ");
            //strSql.AppendLine("          GROUP BY B1.ACCOD, B1.SEQNO, B1.ACDSP, B1.ACRDR ");
            //strSql.AppendLine("          UNION ALL ");
            //strSql.AppendLine("         SELECT C1.ACCOD, D1.ACDSP, D1.ACRDR, C1.ACAMT, C1.ADAMT  ");
            //strSql.AppendLine("           FROM ACTRAN C1 ");
            //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF D1  ");
            //strSql.AppendLine("             ON C1.ACCOD = D1.ACCOD ");
            //strSql.AppendLine("          WHERE C1.TDATE BETWEEN '" + sYmd + "' AND '" + sYmd + "'  ");
            //strSql.AppendLine("            AND C1.ACCOD <> '" + ComnEtcFunc.CashCode + "' ");
            //strSql.AppendLine("          UNION ALL ");
            //strSql.AppendLine("         SELECT '" + ComnEtcFunc.CashCode + "', F1.ACDSP, F1.ACRDR, E1.ADAMT, E1.ACAMT ");
            //strSql.AppendLine("           FROM ACTRAN E1 ");
            //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF F1 ");
            //strSql.AppendLine("             ON E1.ACCOD = F1.ACCOD ");
            //strSql.AppendLine("          WHERE E1.TDATE BETWEEN '" + sYmd + "' AND '" + sYmd + "'  ");
            //strSql.AppendLine("            AND F1.ASMCD = '" + ComnEtcFunc.CashCode + "' ");
            //strSql.AppendLine("          UNION ALL ");
            //strSql.AppendLine("         SELECT G1.ACCOD, F1.ACDSP, F1.ACRDR, G1.ACCRJN, G1.ACDRJN ");
            //strSql.AppendLine("           FROM ACJANF G1 ");
            //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF F1 ");
            //strSql.AppendLine("             ON G1.ACCOD = F1.ACCOD ");
            //strSql.AppendLine("          WHERE G1.ACYEAR = '" + sYmd.Substring(0, 4) + "' ");
            //strSql.AppendLine("       ) X1 ");
            //strSql.AppendLine("  WHERE IFNULL(X1.ACAMT, 0) <> 0 OR IFNULL(X1.ADAMT, 0) <> 0 ");
            //strSql.AppendLine("  GROUP BY X1.ACCOD, X1.ACDSP, X1.ACRDR ");
            //strSql.AppendLine("  ORDER BY X1.ACCOD ");
            //strSql.AppendLine(" ) Z1 ");
            #endregion

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        #region[GridView's Design]

        private void advBandedGridView1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void advBandedGridView1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridView's Design]

        private void AC10001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void AC10001F01_TextChanged(object sender, EventArgs e)
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

        private void advBandedGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                GridView view = sender as GridView;
                GridHitInfo hitInfo = view.CalcHitInfo(e.Location);
                if (hitInfo.InRow || hitInfo.InColumnPanel)
                {
                    DevExpress.Utils.DXMouseEventArgs args = DevExpress.Utils.DXMouseEventArgs.GetMouseArgs(e);
                    if (args != null) args.Handled = true;
                    view.FocusedRowHandle = hitInfo.RowHandle;
                    view.FocusedColumn = hitInfo.Column;
                    //show context menu here  
                }
            }
        }

        private void DateEditYmd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateEditYmd.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateEditYmd.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateEditYmd.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateEditYmd.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }
    }
}