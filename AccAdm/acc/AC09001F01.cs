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
using DevExpress.XtraBars;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraPrinting;
using System.IO;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * 
 * --------------------HISTORY------------------
 * 1. 2021-02-04 (현업요청)
 *    거래처별 계정상세에서 외상매입출금의 경우 마감자료와 JOIN하여 마감항목 출력
 *    
 * 2. 2021-02-05 (현업요청)
 *    외상매입출금과 관련된 마감항목 추가 (마감단가)
 *    
 * 3. 2021-02-07 (현업요청)
 *    거래처 초성검색 추가
 *    
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 *            
 * 수정일자 : 2023-01-02
 * 수정자   : 정은영
 * ID       : #0001
 * 수정내용 : 1.코리아테크, 미래테크 자료 섞이는 현상발생 => 코리아테크의 상호명이 미래테크로 변경되어서 발생한 문제
 *              (현업요청)
 *              현재 상호명인 미래테크로 합쳐서 조회되도록 요청. => 원장조회시 전표에서 저장된 거래처명이 아닌 거래처 코드로 현재 코드명을 가져오도록 수정
 *            
 *            2.거래처 코드 숫자 변환오류 수정
 *            
 * 수정일자 : 2023-01-20
 * 수정자   : 정은영
 * ID       : #0002
 * 수정내용 : 1. 이월금이 있으면 전표가 없어도 조회 되도록 수정
 * 
 * 
 * 수정일자 : 2023-02-07
 * 수정자   : 정은영
 * ID       : #0003
 * 수정내용 : 1. 월계 추가
 * 
 */
namespace AccAdm
{
    public partial class AC09001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC09001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;

        private GridColumn[] ArrGridCol;

        private string PROCEDURE_ID = "DP_AC09001F0122";//#0002

        private void AC09001F01_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            DateEditFrom.EditValue = today.AddDays(1 - today.Day);
            DateEditTo.EditValue = today;
            InitControls();
            UpdateDropDownButton(BtnAcCodSummary);

            ArrGridCol = new GridColumn[] { GridColJBNum, GridColGubun1, GridColDanjung, GridColR2Danga };
            
            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout2.xaml";
            if (File.Exists(sFile))
            {
                layoutControl2.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout3.xaml";
            if (File.Exists(sFile))
            {
                layoutControl3.RestoreLayoutFromXml(sFile);
            }

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewL1, GridViewR1, GridViewL2, GridViewR2 };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sCarryOverAmtGb = RdgbCarryOver.EditValue?.ToString();
            string sCvCod = BtnEditCvCod.EditValue?.ToString();

            /*
             * 2021-01-06 현업요청
             * 원장의 경우 잔액부분 때문에 조회일자 From이 2021년 이전일 경우 메시지 출력 및 2021년도로 다시 세팅
             */
            DateTime dateFrom = DateTime.Parse(DateEditFrom.EditValue?.ToString());
            DateTime dateTo = DateTime.Parse(DateEditTo.EditValue?.ToString());
            if (dateFrom.Year < 2021 || dateTo.Year < 2021)
            {
                XtraMessageBox.Show("거래처 원장은 2021년 01월 01일부터 조회가 가능합니다.");
                DateEditFrom.EditValue = dateFrom.AddYears(2021 - dateFrom.Year);
                DateEditTo.EditValue = dateTo.AddYears(2021 - dateTo.Year);
                return;
            }

            if (XtTab.SelectedTabPageIndex == 0)
            {
                GridL1.DataSource = null;
                GridL1.DataSource = GetDealerInfo(sYmdFrom, sYmdTo, sCarryOverAmtGb, sCvCod);
                if (GridViewL1.RowCount == 0)
                {
                    GridR1.DataSource = null;
                }
            }
            else
            {
                string sAcCod = BtnEditAcCod.EditValue?.ToString().Trim();
                if (string.IsNullOrEmpty(sAcCod))
                {
                    XtraMessageBox.Show("계정코드를 설정하세요.");
                    BtnEditAcCod.Focus();
                    BtnEditAcCod.SelectAll();
                    return;
                }
                GridL2.DataSource = null;
                GridL2.DataSource = GetDealerInfoDetail(sYmdFrom, sYmdTo, sAcCod, sCarryOverAmtGb, sCvCod);
                if(GridViewL2.RowCount == 0)
                {
                    GridR2.DataSource = null;
                }
            }
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC09001F01_KeyDown(object sender, KeyEventArgs e)
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
                DropBtnExcel.PerformClick();
            }
        }
        
        private void GridViewL1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (GridViewL1.RowCount < 1)
            {
                GridR1.DataSource = null;
                return;
            }

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sCvCod = GridViewL1.GetFocusedRowCellValue("CVCOD")?.ToString();

            GridR1.DataSource = GetAccountOfSlipSummary(sYmdFrom, sYmdTo, sCvCod);
        }


        private void GridViewL2_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
                return;

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sAcCod = BtnEditAcCod.EditValue?.ToString().Trim();
            string sCvCod = GridViewL2.GetFocusedRowCellValue("CVCOD")?.ToString();

            GridR2.DataSource = GetAccountDetailOfDealers(sYmdFrom, sYmdTo, sCvCod, sAcCod);
        }

        #region[Query]

        private DataTable GetDealerInfo(string sSlipYmdFrom, string sSlipYmdTo, string sCarryOverAmtGb, string sCvCod)
        {
            #region #0002 이전코드
            //StringBuilder strSql = new StringBuilder();

            //strSql.Clear();
            //strSql.AppendLine(" ");
            //#region mariaDB
            ////strSql.AppendLine(" SELECT DISTINCT A.CVCOD ");
            ////strSql.AppendLine(" 	 , B.DEALER_NM AS CVNAM  ");
            ////strSql.AppendLine(" 	 , B.IDT_NO  ");
            ////strSql.AppendLine(" 	 , B.REP_NM ");
            ////strSql.AppendLine("   FROM ACTRAN A  ");
            ////strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B  ");
            ////strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            ////strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            ////strSql.AppendLine("    AND A.APVYN = 'Y' ");
            ////strSql.AppendLine("    AND (('" + sCvCod + "' = '' AND 1 = 1 ) ");
            ////strSql.AppendLine("         OR ");
            ////strSql.AppendLine("         ('" + sCvCod + "' <> '' AND A.CVCOD = '" + sCvCod + "')) ");
            ////strSql.AppendLine("  ORDER BY REPLACE(B.DEALER_NM, '(주)', '') ");
            //#endregion

            //strSql.AppendLine(" SELECT A.CVCOD ");
            //strSql.AppendLine(" 	 , B.DEALER_NM AS CVNAM  ");
            //strSql.AppendLine(" 	 , B.IDT_NO  ");
            //strSql.AppendLine(" 	 , B.REP_NM ");
            //strSql.AppendLine("   FROM ACTRAN A  ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B  ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            //strSql.AppendLine("    AND A.APVYN = 'Y' ");
            //strSql.AppendLine("    AND (('" + sCvCod + "' = '' AND 1 = 1 ) ");
            //strSql.AppendLine("         OR ");
            ////#0001
            //strSql.AppendLine("         ('" + sCvCod + "' <> '' AND TRY_CONVERT(NUMERIC, '" + sCvCod + "') IS NOT NULL AND A.CVCOD = '" + sCvCod + "')) ");
            //strSql.AppendLine("  GROUP BY A.CVCOD, B.DEALER_NM, B.IDT_NO, B.REP_NM  ");
            //strSql.AppendLine("  ORDER BY REPLACE(B.DEALER_NM, '(주)', '') ");

            //return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            #endregion
            //#0002
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST1");
            dicParams.Add("DATEF", sSlipYmdFrom);
            dicParams.Add("DATET", sSlipYmdTo);
            if (!string.IsNullOrEmpty(sCvCod))
                dicParams.Add("CVCOD", sCvCod);

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }

        private DataTable GetDealerInfoDetail(string sSlipYmdFrom, string sSlipYmdTo, string sAcCod, string sCarryOverAmtGb, string sDealerCd)
        {
            #region #0002 이전코드
            //StringBuilder strSql = new StringBuilder();

            //strSql.Clear();
            //#region mariaDB
            ////strSql.AppendLine(" SELECT * ");
            ////strSql.AppendLine("   FROM ( ");
            ////strSql.AppendLine("          SELECT Z.CVCOD, Z.CVNAM, Z.IDT_NO, SUM(Z.JAMT) AS JAMT ");
            ////strSql.AppendLine("            FROM ( ");
            ////strSql.AppendLine("                   SELECT A1.CVCOD, A1.CVNAM, A1.IDT_NO, SUM(A1.JAMT) AS JAMT   ");
            ////strSql.AppendLine("                     FROM (  ");
            ////strSql.AppendLine("                            SELECT A1.CVCOD, B1.DEALER_NM AS CVNAM, B1.IDT_NO, IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT  ");
            ////strSql.AppendLine("                              FROM ACJANF A1   ");
            ////strSql.AppendLine("                              LEFT JOIN ACC_DEALER_CD B1  ");
            ////strSql.AppendLine("                                ON A1.CVCOD = B1.DEALER_CD  ");
            ////strSql.AppendLine("                             WHERE A1.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "' ");
            ////strSql.AppendLine("                               AND A1.ACCOD = '" + sAcCod + "' ");
            ////strSql.AppendLine("                               AND (('" + sDealerCd + "' = '' AND 1 = 1) ");
            ////strSql.AppendLine("                                    OR  ");
            ////strSql.AppendLine("                                    ('" + sDealerCd + "' <> '' AND A1.CVCOD = '" + sDealerCd + "')) ");
            ////strSql.AppendLine("                             UNION ALL  ");
            ////strSql.AppendLine("                            SELECT A1.CVCOD, A1.CVNAM, C1.IDT_NO AS IDT_NO, SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0)-IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0)-IFNULL(ACAMT, 0) END) JAMT  ");
            ////strSql.AppendLine("                              FROM ACTRAN A1  ");
            ////strSql.AppendLine("                              LEFT JOIN ACMSTF B1 ");
            ////strSql.AppendLine("                                ON A1.ACCOD = B1.ACCOD  ");
            ////strSql.AppendLine("                              LEFT OUTER JOIN ACC_DEALER_CD C1  ");
            ////strSql.AppendLine("                                ON A1.CVCOD = C1.DEALER_CD ");
            ////strSql.AppendLine("                             WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sSlipYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sSlipYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d') ");
            ////strSql.AppendLine("                               AND A1.ACCOD = '" + sAcCod + "'  ");
            ////strSql.AppendLine("                               AND A1.APVYN = 'Y'  ");
            ////strSql.AppendLine("                               AND (('" + sDealerCd + "' = '' AND 1 = 1) ");
            ////strSql.AppendLine("                                    OR  ");
            ////strSql.AppendLine("                                    ('" + sDealerCd + "' <> '' AND A1.CVCOD = '" + sDealerCd + "')) ");
            ////strSql.AppendLine("                             GROUP BY A1.CVCOD, A1.CVNAM  ");
            ////strSql.AppendLine("                          ) A1  ");
            ////strSql.AppendLine("                    GROUP BY A1.CVCOD, A1.CVNAM, A1.IDT_NO  ");
            ////strSql.AppendLine("                    UNION ALL  ");
            ////strSql.AppendLine("                   SELECT A1.CVCOD, A1.CVNAM, C1.IDT_NO AS IDT_NO, SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0)-IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0)-IFNULL(ACAMT, 0) END) JAMT  ");
            ////strSql.AppendLine("                     FROM ACTRAN A1  ");
            ////strSql.AppendLine("                     LEFT JOIN ACMSTF B1 ");
            ////strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD  ");
            ////strSql.AppendLine("                     LEFT OUTER JOIN ACC_DEALER_CD C1  ");
            ////strSql.AppendLine("                       ON A1.CVCOD = C1.DEALER_CD ");
            ////strSql.AppendLine("                    WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            ////strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'  ");
            ////strSql.AppendLine("                      AND A1.APVYN = 'Y'  ");
            ////strSql.AppendLine("                      AND (('" + sDealerCd + "' = '' AND 1 = 1) ");
            ////strSql.AppendLine("                           OR  ");
            ////strSql.AppendLine("                           ('" + sDealerCd + "' <> '' AND A1.CVCOD = '" + sDealerCd + "')) ");
            ////strSql.AppendLine("                    GROUP BY A1.CVCOD  ");
            ////strSql.AppendLine("                ) Z ");
            ////strSql.AppendLine("          GROUP BY Z.CVCOD ");
            ////strSql.AppendLine("        ) XY1");
            ////strSql.AppendLine("  WHERE (('" + sCarryOverAmtGb + "' = '1' AND 1 = 1) ");
            ////strSql.AppendLine("         OR  ");
            ////strSql.AppendLine("         ('" + sCarryOverAmtGb + "' = '2' AND XY1.JAMT != 0) ");
            ////strSql.AppendLine("         OR  ");
            ////strSql.AppendLine("         ('" + sCarryOverAmtGb + "' = '3' AND XY1.JAMT = 0) ) ");
            //#endregion

            ////#0001
            //strSql.AppendLine("SELECT* ");
            //strSql.AppendLine("   FROM(                                                                                                                       ");
            //strSql.AppendLine("          SELECT Z.CVCOD, Z.CVNAM, Z.IDT_NO, SUM(Z.JAMT) AS JAMT                                                               ");
            //strSql.AppendLine("            FROM(                                                                                                              ");
            //strSql.AppendLine("                   SELECT A1.CVCOD, A1.CVNAM, A1.IDT_NO, SUM(A1.JAMT) AS JAMT                                                  ");
            //strSql.AppendLine("                     FROM(                                                                                                     ");
            //strSql.AppendLine("                            SELECT A1.CVCOD, B1.DEALER_NM AS CVNAM, B1.IDT_NO, ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT");
            //strSql.AppendLine("                              FROM ACJANF A1                                                                                   ");
            //strSql.AppendLine("                              LEFT JOIN ACC_DEALER_CD B1                                                                       ");
            //strSql.AppendLine("                                ON A1.CVCOD = B1.DEALER_CD                                                                     ");
            //strSql.AppendLine("                             WHERE A1.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "'                                                                          ");
            //strSql.AppendLine("                               AND A1.ACCOD = '" + sAcCod + "'                                                                           ");
            //strSql.AppendLine("                               AND(('" + sDealerCd + "' = '' AND 1 = 1)                                                               ");
            //strSql.AppendLine("                                    OR                                                                                         ");
            //strSql.AppendLine("                                    ('" + sDealerCd + "' <> '' AND TRY_CONVERT(NUMERIC, '" + sDealerCd + "') IS NOT NULL AND A1.CVCOD = '" + sDealerCd + "'))                                          ");
            //strSql.AppendLine("                             UNION ALL                                                                                         ");
            //strSql.AppendLine("                            SELECT A1.CVCOD                                                                                    ");
            //strSql.AppendLine("                                 , C1.DEALER_NM AS CVNAM                                                                                    ");
            //strSql.AppendLine("                                 , C1.IDT_NO AS IDT_NO                                                                         ");
            //strSql.AppendLine("                                 , SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(A1.ACAMT, 0) - ISNULL(A1.ADAMT, 0)                 ");
            //strSql.AppendLine("                                            ELSE ISNULL(A1.ADAMT, 0) - ISNULL(ACAMT, 0) END) JAMT                              ");
            //strSql.AppendLine("                              FROM ACTRAN A1                                                                                   ");
            //strSql.AppendLine("                              LEFT JOIN ACMSTF B1                                                                              ");
            //strSql.AppendLine("                                ON A1.ACCOD = B1.ACCOD                                                                         ");
            //strSql.AppendLine("                              LEFT OUTER JOIN ACC_DEALER_CD C1                                                                 ");
            //strSql.AppendLine("                                ON A1.CVCOD = C1.DEALER_CD                                                                     ");
            //strSql.AppendLine("                             WHERE A1.TDATE BETWEEN SUBSTRING('" + sSlipYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + sSlipYmdFrom + "')), 112)");
            //strSql.AppendLine("                               AND A1.ACCOD = '" + sAcCod + "'                                                     ");
            //strSql.AppendLine("                               AND A1.APVYN = 'Y'                                                        ");
            //strSql.AppendLine("                               AND(('" + sDealerCd + "' = '' AND 1 = 1)                                         ");
            //strSql.AppendLine("                                    OR                                                                   ");
            //strSql.AppendLine("                                    ('" + sDealerCd + "' <> '' AND TRY_CONVERT(NUMERIC, '" + sDealerCd + "') IS NOT NULL AND A1.CVCOD = '" + sDealerCd + "'))                    ");
            //strSql.AppendLine("                             GROUP BY A1.CVCOD, C1.DEALER_NM, C1.IDT_NO                                      ");
            //strSql.AppendLine("                          ) A1                                                                           ");
            //strSql.AppendLine("                    GROUP BY A1.CVCOD, A1.CVNAM, A1.IDT_NO                                               ");
            //strSql.AppendLine("                    UNION ALL                                                                            ");
            //strSql.AppendLine("                   SELECT A1.CVCOD                                                                       ");
            //strSql.AppendLine("                        , C1.DEALER_NM AS CVNAM                                                                       ");
            //strSql.AppendLine("                        , C1.IDT_NO AS IDT_NO                                                            ");
            //strSql.AppendLine("                        , SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(A1.ACAMT, 0) - ISNULL(A1.ADAMT, 0)    ");
            //strSql.AppendLine("                                   ELSE ISNULL(A1.ADAMT, 0) - ISNULL(ACAMT, 0) END) JAMT                 ");
            //strSql.AppendLine("                       FROM ACTRAN A1                                                                    ");
            //strSql.AppendLine("                     LEFT JOIN ACMSTF B1                                                                 ");
            //strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD                                                            ");
            //strSql.AppendLine("                     LEFT OUTER JOIN ACC_DEALER_CD C1                                                    ");
            //strSql.AppendLine("                       ON A1.CVCOD = C1.DEALER_CD                                                        ");
            //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "'                                     ");
            //strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'                                                              ");
            //strSql.AppendLine("                      AND A1.APVYN = 'Y'                                                                 ");
            //strSql.AppendLine("                      AND(('" + sDealerCd + "' = '' AND 1 = 1)                                                  ");
            //strSql.AppendLine("                           OR                                                                            ");
            //strSql.AppendLine("                           ('" + sDealerCd + "' <> '' AND TRY_CONVERT(NUMERIC, '" + sDealerCd + "') IS NOT NULL AND A1.CVCOD = '" + sDealerCd + "'))                             ");
            //strSql.AppendLine("                    GROUP BY A1.CVCOD, C1.DEALER_NM, C1.IDT_NO                                               ");
            //strSql.AppendLine("                ) Z                                                                                      ");
            //strSql.AppendLine("          GROUP BY Z.CVCOD, Z.CVNAM, Z.IDT_NO                                                            ");
            //strSql.AppendLine("        ) XY1                                                                                            ");
            //strSql.AppendLine("  WHERE (('" + sCarryOverAmtGb + "' = '1' AND 1 = 1) ");
            //strSql.AppendLine("         OR  ");
            //strSql.AppendLine("         ('" + sCarryOverAmtGb + "' = '2' AND XY1.JAMT != 0) ");
            //strSql.AppendLine("         OR  ");
            //strSql.AppendLine("         ('" + sCarryOverAmtGb + "' = '3' AND XY1.JAMT = 0) ) ");

            //return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            #region 이전코드
            //if (string.IsNullOrEmpty(sDealerCd))
            //{
            //    strSql.AppendLine(" SELECT A1.CVCOD, A1.CVNAM, A1.IDT_NO, SUM(A1.JAMT) AS JAMT   ");
            //    strSql.AppendLine("   FROM (  ");
            //    strSql.AppendLine("          SELECT A1.CVCOD, B1.DEALER_NM AS CVNAM, B1.IDT_NO, IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT  ");
            //    strSql.AppendLine("            FROM ACJANF A1   ");
            //    strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B1  ");
            //    strSql.AppendLine("              ON A1.CVCOD = B1.DEALER_CD  ");
            //    strSql.AppendLine("           WHERE A1.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "' ");
            //    strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "' ");
            //    strSql.AppendLine("             AND (('" + sDealerCd + "' = '' AND 1 = 1) ");
            //    strSql.AppendLine("                  OR  ");
            //    strSql.AppendLine("                  ('" + sDealerCd + "' <> '' AND A1.CVCOD = '" + sDealerCd + "')) ");
            //    strSql.AppendLine("           UNION ALL  ");
            //    strSql.AppendLine("          SELECT A1.CVCOD, A1.CVNAM, C1.IDT_NO AS IDT_NO, SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0)-IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0)-IFNULL(ACAMT, 0) END) JAMT  ");
            //    strSql.AppendLine("            FROM ACTRAN A1  ");
            //    strSql.AppendLine("            LEFT JOIN ACMSTF B1 ");
            //    strSql.AppendLine("              ON A1.ACCOD = B1.ACCOD  ");
            //    strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C1  ");
            //    strSql.AppendLine("              ON A1.CVCOD = C1.DEALER_CD ");
            //    strSql.AppendLine("           WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            //    strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "'  ");
            //    strSql.AppendLine("             AND A1.APVYN = 'Y'  ");
            //    strSql.AppendLine("             AND (('" + sDealerCd + "' = '' AND 1 = 1) ");
            //    strSql.AppendLine("                  OR  ");
            //    strSql.AppendLine("                  ('" + sDealerCd + "' <> '' AND A1.CVCOD = '" + sDealerCd + "')) ");
            //    strSql.AppendLine("           GROUP BY A1.CVCOD, A1.CVNAM  ");
            //    strSql.AppendLine("        ) A1  ");
            //    strSql.AppendLine("  GROUP BY A1.CVCOD, A1.CVNAM, A1.IDT_NO  ");
            //    strSql.AppendLine("  ORDER BY A1.CVCOD   ");
            //}
            //else
            //{
            //    strSql.AppendLine(" SELECT A1.CVCOD, A1.CVNAM, A1.IDT_NO, SUM(A1.JAMT) AS JAMT   ");
            //    strSql.AppendLine("   FROM (  ");
            //    strSql.AppendLine("          SELECT A1.CVCOD, B1.DEALER_NM AS CVNAM, B1.IDT_NO, IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT  ");
            //    strSql.AppendLine("            FROM ACJANF A1   ");
            //    strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B1  ");
            //    strSql.AppendLine("              ON A1.CVCOD = B1.DEALER_CD  ");
            //    strSql.AppendLine("           WHERE A1.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "' ");
            //    strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "' ");
            //    strSql.AppendLine("             AND A1.CVCOD = " + sDealerCd + " ");
            //    strSql.AppendLine("           UNION ALL  ");
            //    strSql.AppendLine("          SELECT A1.CVCOD, A1.CVNAM, C1.IDT_NO AS IDT_NO, SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0)-IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0)-IFNULL(ACAMT, 0) END) JAMT  ");
            //    strSql.AppendLine("            FROM ACTRAN A1  ");
            //    strSql.AppendLine("            LEFT JOIN ACMSTF B1 ");
            //    strSql.AppendLine("              ON A1.ACCOD = B1.ACCOD  ");
            //    strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C1  ");
            //    strSql.AppendLine("              ON A1.CVCOD = C1.DEALER_CD ");
            //    strSql.AppendLine("           WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            //    strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "'  ");
            //    strSql.AppendLine("             AND A1.APVYN = 'Y'  ");
            //    strSql.AppendLine("             AND A1.CVCOD = " + sDealerCd + " ");
            //    strSql.AppendLine("           GROUP BY A1.CVCOD, A1.CVNAM  ");
            //    strSql.AppendLine("        ) A1  ");
            //    strSql.AppendLine("  GROUP BY A1.CVCOD, A1.CVNAM, A1.IDT_NO  ");
            //    strSql.AppendLine("  UNION ALL ");
            //    strSql.AppendLine(" SELECT A1.CVCOD, A1.CVNAM, A1.IDT_NO, SUM(A1.JAMT) AS JAMT   ");
            //    strSql.AppendLine("   FROM (  ");
            //    strSql.AppendLine("          SELECT A1.CVCOD, B1.DEALER_NM AS CVNAM, B1.IDT_NO, IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT  ");
            //    strSql.AppendLine("            FROM ACJANF A1   ");
            //    strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B1  ");
            //    strSql.AppendLine("              ON A1.CVCOD = B1.DEALER_CD  ");
            //    strSql.AppendLine("           WHERE A1.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "' ");
            //    strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "' ");
            //    strSql.AppendLine("             AND A1.CVCOD <> " + sDealerCd + " ");
            //    strSql.AppendLine("           UNION ALL  ");
            //    strSql.AppendLine("          SELECT A1.CVCOD, A1.CVNAM, C1.IDT_NO AS IDT_NO, SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0)-IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0)-IFNULL(ACAMT, 0) END) JAMT  ");
            //    strSql.AppendLine("            FROM ACTRAN A1  ");
            //    strSql.AppendLine("            LEFT JOIN ACMSTF B1 ");
            //    strSql.AppendLine("              ON A1.ACCOD = B1.ACCOD  ");
            //    strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C1  ");
            //    strSql.AppendLine("              ON A1.CVCOD = C1.DEALER_CD ");
            //    strSql.AppendLine("           WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            //    strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "'  ");
            //    strSql.AppendLine("             AND A1.APVYN = 'Y'  ");
            //    strSql.AppendLine("             AND A1.CVCOD <> " + sDealerCd + " ");
            //    strSql.AppendLine("           GROUP BY A1.CVCOD, A1.CVNAM  ");
            //    strSql.AppendLine("        ) A1  ");
            //    strSql.AppendLine("  GROUP BY A1.CVCOD, A1.CVNAM, A1.IDT_NO  ");
            //}

            //strSql.Clear();
            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT X1.CVCOD ");
            //strSql.AppendLine("      , X3.DEALER_NM AS CVNAM ");
            //strSql.AppendLine("      , X3.IDT_NO ");
            //strSql.AppendLine("      , X3.REP_NM ");
            //strSql.AppendLine("      , SUM(CASE WHEN X2.ACRDR = '1' THEN X1.ACCRJN - X1.ACDRJN ");
            //strSql.AppendLine("                                     ELSE X1.ACDRJN - X1.ACCRJN END ) AS BAL_AMT ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A1.CVCOD ");
            //strSql.AppendLine("          	  , A1.ACCRJN  ");
            //strSql.AppendLine("          	  , A1.ACDRJN ");
            //strSql.AppendLine("            FROM ACJANF A1 ");
            //strSql.AppendLine("           WHERE A1.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "' ");
            //strSql.AppendLine("             AND A1.ACCOD = '" + sAcCod + "' ");
            //strSql.AppendLine("           UNION ALL  ");
            //strSql.AppendLine("          SELECT A2.CVCOD  ");
            //strSql.AppendLine("          	  , SUM(A2.ACAMT) AS CAMT  ");
            //strSql.AppendLine("          	  , SUM(A2.ADAMT) AS DAMT ");
            //strSql.AppendLine("            FROM ACTRAN A2  ");
            //strSql.AppendLine("           WHERE A2.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            //strSql.AppendLine("             AND A2.ACCOD = '" + sAcCod + "'  ");
            //strSql.AppendLine("           GROUP BY CVCOD ");
            //strSql.AppendLine("        ) X1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF X2  ");
            //strSql.AppendLine("     ON X2.ACCOD = '" + sAcCod + "' ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD X3  ");
            //strSql.AppendLine("     ON X1.CVCOD = X3.DEALER_CD ");
            //strSql.AppendLine("  GROUP BY X1.CVCOD ");
            //strSql.AppendLine("  ORDER BY X3.DEALER_NM ");
            #endregion
            #endregion

            string sJanGb = RdgbCarryOver.EditValue?.ToString();
            //#0002
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            
            dicParams.Clear();
            dicParams.Add("CMD", "LIST3");
            dicParams.Add("DATEF", sSlipYmdFrom);
            dicParams.Add("DATET", sSlipYmdTo);
            if(!string.IsNullOrEmpty(sDealerCd))
                dicParams.Add("CVCOD", sDealerCd);
            dicParams.Add("JANGB", sJanGb);
            dicParams.Add("ACCOD", sAcCod);

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }

        private DataTable GetAccountOfSlipSummary(string sSlipYmdFrom, string sSlipYmdTo, string sCvCod)
        {
            #region #0002 이전코드
            //StringBuilder strSql = new StringBuilder();

            //strSql.Clear();
            //strSql.AppendLine(" ");
            //#region mariaDB
            ////strSql.AppendLine(" SELECT A1.ACCOD ");
            ////strSql.AppendLine("      , B1.ACDSP AS ACNAM ");
            ////strSql.AppendLine("      , B1.ACRDR ");
            ////strSql.AppendLine("      , (CASE WHEN B1.ACRDR='1' THEN   ");
            ////strSql.AppendLine("                  IFNULL(B2.ACDRJN ,0) - IFNULL(B2.ACCRJN, 0)   ");
            ////strSql.AppendLine("                + IFNULL(CAMT,0) - IFNULL(DAMT,0)   ");
            ////strSql.AppendLine("             ELSE IFNULL(B2.ACCRJN ,0)-IFNULL(B2.ACDRJN, 0)   ");
            ////strSql.AppendLine("                + IFNULL(DAMT,0) - IFNULL(CAMT,0) END) AS CARRY_OVER_AMT   ");
            ////strSql.AppendLine("      , (CASE WHEN B1.ACRDR='1' THEN    ");
            ////strSql.AppendLine("                  IFNULL(B2.ACDRJN ,0) - IFNULL(B2.ACCRJN, 0) + IFNULL(CAMT,0)     ");
            ////strSql.AppendLine("                - IFNULL(DAMT,0) + SUM(IFNULL(ACAMT, 0)) - SUM(IFNULL(ADAMT, 0))    ");
            ////strSql.AppendLine("             ELSE IFNULL(B2.ACCRJN ,0) - IFNULL(B2.ACDRJN, 0) + IFNULL(DAMT,0)     ");
            ////strSql.AppendLine("                - IFNULL(CAMT,0) + SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) END) AS BAL_AMT    ");
            ////strSql.AppendLine("      , SUM(IFNULL(ACAMT, 0)) AS ACAMT   ");
            ////strSql.AppendLine("      , SUM(IFNULL(ADAMT, 0)) AS ADAMT   ");
            //////strSql.AppendLine("      , (CASE WHEN B1.ACRDR='1' THEN  ");
            //////strSql.AppendLine("                 IFNULL(B2.ACCRJN, 0) - IFNULL(B2.ACDRJN ,0)  ");
            //////strSql.AppendLine("                + IFNULL(CAMT,0) - IFNULL(DAMT,0)  ");
            //////strSql.AppendLine("            ELSE IFNULL(B2.ACDRJN ,0)-IFNULL(B2.ACCRJN, 0)  ");
            //////strSql.AppendLine("                + IFNULL(DAMT,0) - IFNULL(CAMT,0) END) AS CARRY_OVER_AMT  ");
            //////strSql.AppendLine("      , (CASE WHEN B1.ACRDR='1' THEN  ");
            //////strSql.AppendLine("                 IFNULL(B2.ACCRJN ,0) - IFNULL(B2.ACDRJN, 0) + IFNULL(CAMT,0)   ");
            //////strSql.AppendLine("                - IFNULL(DAMT,0) + SUM(IFNULL(ACAMT, 0)) - SUM(IFNULL(ADAMT, 0))  ");
            //////strSql.AppendLine("            ELSE IFNULL(B2.ACDRJN, 0) - IFNULL(B2.ACCRJN, 0) + IFNULL(DAMT,0)   ");
            //////strSql.AppendLine("                - IFNULL(CAMT,0) + SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) END) AS BAL_AMT  ");
            //////strSql.AppendLine("      , SUM(IFNULL(ACAMT, 0)) AS ACAMT  ");
            //////strSql.AppendLine("      , SUM(IFNULL(ADAMT, 0)) AS ADAMT  ");
            ////strSql.AppendLine("  FROM ACTRAN A1 ");
            ////strSql.AppendLine("  LEFT OUTER JOIN ACMSTF B1  ");
            ////strSql.AppendLine("    ON A1.ACCOD = B1.ACCOD ");
            ////strSql.AppendLine("  LEFT OUTER JOIN ACJANF B2  ");
            ////strSql.AppendLine("    ON A1.ACCOD = B2.ACCOD ");
            ////strSql.AppendLine("   AND A1.CVCOD = B2.CVCOD ");
            ////strSql.AppendLine("   AND B2.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "'  ");
            ////strSql.AppendLine("  LEFT OUTER JOIN ");
            ////strSql.AppendLine("      (SELECT A1.ACCOD, A1.CVCOD, SUM(ACAMT)CAMT, SUM(ADAMT)DAMT ");
            ////strSql.AppendLine("         FROM ACTRAN A1 ");
            ////strSql.AppendLine("        WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sSlipYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sSlipYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d') ");
            ////strSql.AppendLine("        GROUP BY A1.ACCOD, A1.CVCOD ");
            ////strSql.AppendLine("        ) B3 ON A1.ACCOD = B3.ACCOD AND A1.CVCOD = B3.CVCOD ");
            ////strSql.AppendLine(" WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            ////strSql.AppendLine("   AND A1.CVCOD = '" + sCvCod + "' ");
            ////strSql.AppendLine("   AND A1.APVYN = 'Y' ");
            ////strSql.AppendLine(" GROUP BY A1.ACCOD, B1.ACDSP, B1.ACRDR, B2.ACCRJN, B2.ACDRJN, CAMT, DAMT ");
            ////strSql.AppendLine(" ORDER BY A1.ACCOD ");

            //////strSql.AppendLine(" SELECT A1.ACCOD          ");
            //////strSql.AppendLine("      , B1.ACDSP AS ACNAM ");
            //////strSql.AppendLine("      , B1.ACRDR          ");
            //////strSql.AppendLine("      , (CASE WHEN B1.ACRDR = '1' THEN                           ");
            //////strSql.AppendLine("                   (IFNULL(B2.ACDRJN ,0) - IFNULL(B2.ACCRJN, 0)) ");
            //////strSql.AppendLine("                 + IFNULL(CAMT,0) - IFNULL(DAMT,0)               ");
            //////strSql.AppendLine("              ELSE (IFNULL(B2.ACCRJN, 0) - IFNULL(B2.ACDRJN ,0)) ");
            //////strSql.AppendLine("                 + IFNULL(DAMT,0) - IFNULL(CAMT,0) END) AS CARRY_OVER_AMT /*이월금액*/  ");
            //////strSql.AppendLine("      , IFNULL(SUM(A1.ACAMT),0) AS ACAMT /*차변금액*/ ");
            //////strSql.AppendLine("      , IFNULL(SUM(A1.ADAMT),0) AS ADAMT /*대변금액*/ ");
            //////strSql.AppendLine("      , (CASE WHEN B1.ACRDR = '1' THEN                ");
            //////strSql.AppendLine("                   IFNULL(B2.ACDRJN ,0) - IFNULL(B2.ACCRJN, 0) + IFNULL(CAMT,0)         ");
            //////strSql.AppendLine("                 - IFNULL(DAMT,0) + ifnull(SUM(A1.ACAMT),0) - ifnull(SUM(A1.ADAMT),0)   ");
            //////strSql.AppendLine("              ELSE IFNULL(B2.ACCRJN, 0) - IFNULL(B2.ACDRJN, 0) + IFNULL(DAMT,0)         ");
            //////strSql.AppendLine("                 - IFNULL(CAMT,0) + ifnull(SUM(A1.ADAMT),0) - ifnull(SUM(A1.ACAMT),0) END) AS BAL_AMT /*잔액*/");
            //////strSql.AppendLine("  FROM ACTRAN A1            ");
            //////strSql.AppendLine("  LEFT OUTER JOIN ACMSTF B1 ");
            //////strSql.AppendLine("    ON A1.ACCOD = B1.ACCOD  ");
            //////strSql.AppendLine("  LEFT OUTER JOIN ACJANF B2 ");
            //////strSql.AppendLine("    ON A1.ACCOD = B2.ACCOD ");
            //////strSql.AppendLine("   AND A1.CVCOD = B2.CVCOD ");
            //////strSql.AppendLine("   AND B2.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "'  ");
            //////strSql.AppendLine("  LEFT OUTER JOIN   ");
            //////strSql.AppendLine("      (SELECT A1.ACCOD, A1.CVCOD, SUM(ACAMT)CAMT, SUM(ADAMT)DAMT  ");
            //////strSql.AppendLine("         FROM ACTRAN A1  ");
            //////strSql.AppendLine("        WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom.Substring(0, 4) + "0101' AND DATE_SUB('" + sSlipYmdFrom + "',INTERVAL 1 DAY)");
            //////strSql.AppendLine("        GROUP BY A1.ACCOD, A1.CVCOD ");
            //////strSql.AppendLine("        ) B3 ON A1.ACCOD = B3.ACCOD AND A1.CVCOD = B3.CVCOD ");
            //////strSql.AppendLine(" WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "' ");
            //////strSql.AppendLine("   AND A1.CVCOD = '" + sCvCod + "' ");
            //////strSql.AppendLine(" GROUP BY A1.ACCOD, B1.ACDSP, B1.ACRDR, B2.ACCRJN, B2.ACDRJN, CAMT, DAMT  ");
            //////strSql.AppendLine(" ORDER BY A1.ACCOD ");
            //#endregion

            ////#0001
            //strSql.AppendLine("SELECT A1.ACCOD                                                                                                                             ");
            //strSql.AppendLine("     , B1.ACDSP AS ACNAM                                                                                                                    ");
            //strSql.AppendLine("     , B1.ACRDR                                                                                                                             ");
            //strSql.AppendLine("     , (CASE WHEN B1.ACRDR = '1' THEN                                                                                                       ");
            //strSql.AppendLine("                 ISNULL(B2.ACDRJN, 0) - ISNULL(B2.ACCRJN, 0)                                                                                ");
            //strSql.AppendLine("               + ISNULL(CAMT, 0) - ISNULL(DAMT, 0)                                                                                          ");
            //strSql.AppendLine("            ELSE ISNULL(B2.ACCRJN ,0)-ISNULL(B2.ACDRJN, 0)                                                                                  ");
            //strSql.AppendLine("               + ISNULL(DAMT, 0) - ISNULL(CAMT, 0) END) AS CARRY_OVER_AMT                                                                   ");
            //strSql.AppendLine("       , (CASE WHEN B1.ACRDR = '1' THEN                                                                                                     ");
            //strSql.AppendLine("                                                                                                                                            ");
            //strSql.AppendLine("                   ISNULL(B2.ACDRJN, 0) - ISNULL(B2.ACCRJN, 0) + ISNULL(CAMT, 0)                                                            ");
            //strSql.AppendLine("                 - ISNULL(DAMT, 0) + SUM(ISNULL(ACAMT, 0)) - SUM(ISNULL(ADAMT, 0))                                                          ");
            //strSql.AppendLine("                                                                                                                                            ");
            //strSql.AppendLine("              ELSE ISNULL(B2.ACCRJN, 0) - ISNULL(B2.ACDRJN, 0) + ISNULL(DAMT, 0)                                                            ");
            //strSql.AppendLine("                 - ISNULL(CAMT, 0) + SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) END) AS BAL_AMT                                          ");
            //strSql.AppendLine("        , SUM(ISNULL(ACAMT, 0)) AS ACAMT                                                                                                    ");
            //strSql.AppendLine("        , SUM(ISNULL(ADAMT, 0)) AS ADAMT                                                                                                    ");
            //strSql.AppendLine(" FROM ACTRAN A1                                                                                                                             ");
            //strSql.AppendLine(" LEFT OUTER JOIN ACMSTF B1                                                                                                                  ");
            //strSql.AppendLine("   ON A1.ACCOD = B1.ACCOD                                                                                                                   ");
            //strSql.AppendLine(" LEFT OUTER JOIN ACJANF B2                                                                                                                  ");
            //strSql.AppendLine("   ON A1.ACCOD = B2.ACCOD                                                                                                                   ");
            //strSql.AppendLine("  AND A1.CVCOD = B2.CVCOD                                                                                                                   ");
            //strSql.AppendLine("  AND B2.ACYEAR = '" + sSlipYmdFrom.Substring(0, 4) + "'                                                                                                                    ");
            //strSql.AppendLine(" LEFT OUTER JOIN                                                                                                                            ");
            //strSql.AppendLine("     (SELECT A1.ACCOD, A1.CVCOD, SUM(ACAMT)CAMT, SUM(ADAMT)DAMT                                                                             ");
            //strSql.AppendLine("        FROM ACTRAN A1                                                                                                                      ");
            //strSql.AppendLine("       WHERE A1.TDATE BETWEEN SUBSTRING('" + sSlipYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + sSlipYmdFrom + "')), 112)");
            //strSql.AppendLine("       GROUP BY A1.ACCOD, A1.CVCOD                                                                                                          ");
            //strSql.AppendLine("       ) B3 ON A1.ACCOD = B3.ACCOD AND A1.CVCOD = B3.CVCOD                                                                                  ");
            //strSql.AppendLine("WHERE A1.TDATE BETWEEN '" + sSlipYmdFrom + "' AND '" + sSlipYmdTo + "'                                                                                            ");
            //strSql.AppendLine("  AND A1.CVCOD = '" + sCvCod + "'                                                                                                               ");
            //strSql.AppendLine("  AND A1.APVYN = 'Y'                                                                                                                        ");
            //strSql.AppendLine("GROUP BY A1.ACCOD, B1.ACDSP, B1.ACRDR, B2.ACCRJN, B2.ACDRJN, CAMT, DAMT                                                                     ");
            //strSql.AppendLine("ORDER BY A1.ACCOD                                                                                                                           ");

            //return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            #endregion
            //#0002
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST2");
            dicParams.Add("DATEF", sSlipYmdFrom);
            dicParams.Add("DATET", sSlipYmdTo);
            dicParams.Add("CVCOD", sCvCod);

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }

        private DataTable GetAccountDetailOfDealers(string sYmdFrom, string sYmdTo, string sCvCod, string sAcCod)
        {
            #region #0002 이전코드
            //StringBuilder strSql = new StringBuilder();

            ///*
            // * 2021-02-04(현업요청)
            // * 계정상세내역에서 마감자료와 JOIN걸어 등급, 차번, 인수량을 가져오도록 수정
            // * 대상은 외상매입금, 외상매출금을 대상으로 하여 계정코드가 해당 코드일 경우 분기를 태운다
            // */
            //if (sAcCod.Equals("0251") || sAcCod.Equals("0108"))
            //{
            //    //외상매입금 및 외상매출금의 경우 추가되는 컬럼 Visible true;
            //    foreach(GridColumn col in ArrGridCol)
            //    {
            //        col.Visible = true;
            //    }

            //    strSql.Clear();
            //    strSql.AppendLine("--상세(거래처별)                                                                     ");
            //    strSql.AppendLine("WITH DAY_INFO AS(                                                                    ");
            //    strSql.AppendLine("    SELECT '2' AS IDX                                                                ");
            //    strSql.AppendLine("         , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE1               ");
            //    strSql.AppendLine("         , '[일계]' AS TDATE                                                         ");
            //    strSql.AppendLine("         , MIN(A1.SEQNO) AS SEQNO                                                                   ");
            //    strSql.AppendLine("         , MIN(A1.LINNO) AS LINNO                                                                  ");
            //    strSql.AppendLine("         , '' AS ACRDR                                                               ");
            //    strSql.AppendLine("         , '' AS ATEXT                                                               ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ACAMT, 0)) AS ACAMT                                         ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ADAMT, 0)) AS ADAMT                                         ");
            //    strSql.AppendLine("         , '' AS ACCOD                                                               ");
            //    strSql.AppendLine("         , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT                                        ");
            //    strSql.AppendLine("         , SUM(C1.DANJUNG) AS DANJUNG                                                ");
            //    strSql.AppendLine("      FROM ACTRAN A1                                                                 ");
            //    strSql.AppendLine("      LEFT JOIN ACMSTF B1                                                            ");
            //    strSql.AppendLine("        ON A1.ACCOD = B1.ACCOD                                                       ");
            //    strSql.AppendLine("      LEFT JOIN INLIST C1                                                            ");
            //    strSql.AppendLine("        ON A1.REF1 = CONVERT(VARCHAR,C1.JUNPYOID)                                                     ");
            //    strSql.AppendLine("     WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'                                ");
            //    strSql.AppendLine("       AND A1.ACCOD = '" + sAcCod + "'                                                         ");
            //    strSql.AppendLine("       AND A1.CVCOD = " + sCvCod + "                                                     ");
            //    strSql.AppendLine("       AND A1.APVYN = 'Y'                                                            ");
            //    strSql.AppendLine("     GROUP BY A1.TDATE                                          ");
            //    strSql.AppendLine("     UNION ALL                                                                       ");
            //    strSql.AppendLine("    SELECT '3' AS IDX                                                                ");
            //    strSql.AppendLine("         , CONCAT(CONVERT(VARCHAR(10), CONVERT(DATE, '" + sYmdTo + "'), 23), '1') AS TDATE1");
            //    strSql.AppendLine("         , '[누계]' AS TDATE, '9999' AS SEQNO                                        ");
            //    strSql.AppendLine("         , 99 AS LINNO                                                               ");
            //    strSql.AppendLine("         , '' AS ACRDR                                                               ");
            //    strSql.AppendLine("         , '' AS ATEXT                                                               ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ACAMT, 0)) AS ACAMT                                         ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD                            ");
            //    strSql.AppendLine("         , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT                                        ");
            //    strSql.AppendLine("         , SUM(C1.DANJUNG) AS DANJUNG                                                ");
            //    strSql.AppendLine("      FROM ACTRAN A1                                                                 ");
            //    strSql.AppendLine("      LEFT JOIN ACMSTF B1                                                            ");
            //    strSql.AppendLine("        ON A1.ACCOD = B1.ACCOD                                                       ");
            //    strSql.AppendLine("      LEFT JOIN INLIST C1                                                            ");
            //    strSql.AppendLine("        ON A1.REF1 = CONVERT(VARCHAR,C1.JUNPYOID)                                                     ");
            //    strSql.AppendLine("     WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'                                ");
            //    strSql.AppendLine("       AND A1.ACCOD = '" + sAcCod + "'                                                         ");
            //    strSql.AppendLine("       AND A1.CVCOD = " + sCvCod + "                                                     ");
            //    strSql.AppendLine("       AND A1.APVYN = 'Y'                                                            ");
            //    strSql.AppendLine(")                                                                                    ");
            //    strSql.AppendLine("SELECT*                                                                              ");
            //    strSql.AppendLine("  FROM(                                                                              ");
            //    strSql.AppendLine("        SELECT A1.IDX, A1.TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A1.ACNAM");
            //    strSql.AppendLine("             , 0 AS JJAMT                                                                                 ");
            //    strSql.AppendLine("             , '' AS REF1                                                                                 ");
            //    strSql.AppendLine("             , '' AS REF2                                                                                 ");
            //    strSql.AppendLine("             , '' AS REF3                                                                                 ");
            //    strSql.AppendLine("             , FORMAT(A1.DANJUNG, '#,0') AS DANJUNG                                                       ");
            //    strSql.AppendLine("             , '' AS GUBUN1                                                                               ");
            //    strSql.AppendLine("             , '' AS J_BNUM                                                                               ");
            //    strSql.AppendLine("             , '' AS DANGA                                                                                ");
            //    strSql.AppendLine("          FROM DAY_INFO A1                                                                                ");
            //    strSql.AppendLine("          UNION ALL                                                                                       ");
            //    strSql.AppendLine("         SELECT A1.IDX, A1.TDATE AS TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A2.ACNAM");
            //    strSql.AppendLine("              , SUM((CASE WHEN A1.ACRDR = '1' THEN ISNULL(A1.ACAMT, 0) - ISNULL(A1.ADAMT, 0) ELSE ISNULL(A1.ADAMT, 0) - ISNULL(A1.ACAMT, 0) END) + JAMT) OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) JJAMT");
            //    strSql.AppendLine("              , A1.REF1                              ");
            //    strSql.AppendLine("              , A1.REF2                              ");
            //    strSql.AppendLine("              , A1.REF3                              ");
            //    strSql.AppendLine("              , FORMAT(A1.DANJUNG, '#,0') AS DANJUNG ");
            //    strSql.AppendLine("              , A1.GUBUN1                            ");
            //    strSql.AppendLine("              , A1.J_BNUM                            ");
            //    strSql.AppendLine("              , FORMAT(A1.DANGA, '#,0') AS DANGA     ");
            //    strSql.AppendLine("           FROM(SELECT '0' AS IDX                    ");
            //    strSql.AppendLine("                      , ' [이월]' TDATE              ");
            //    strSql.AppendLine("                      , ' 'SEQNO                     ");
            //    strSql.AppendLine("                      , 0 LINNO                      ");
            //    strSql.AppendLine("                      , '' ACRDR                     ");
            //    strSql.AppendLine("                      , ' 'ATEXT                     ");
            //    strSql.AppendLine("                      , 0 ACAMT                      ");
            //    strSql.AppendLine("                      , 0 ADAMT                      ");
            //    strSql.AppendLine("                      , SUM(JAMT)JAMT                ");
            //    strSql.AppendLine("                      , A1.ACCOD                     ");
            //    strSql.AppendLine("                      , '' AS REF1                   ");
            //    strSql.AppendLine("                      , '' AS REF2                   ");
            //    strSql.AppendLine("                      , '' AS REF3                   ");
            //    strSql.AppendLine("                      , '' AS DANJUNG                ");
            //    strSql.AppendLine("                      , '' AS GUBUN1                 ");
            //    strSql.AppendLine("                      , '' AS J_BNUM                 ");
            //    strSql.AppendLine("                      , '' AS DANGA                  ");
            //    strSql.AppendLine("                   FROM(                             ");
            //    strSql.AppendLine("                          SELECT ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT, A1.ACCOD");
            //    strSql.AppendLine("                            FROM ACJANF A1                                                 ");
            //    strSql.AppendLine("                           WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'                                        ");
            //    strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'                                         ");
            //    strSql.AppendLine("                             AND A1.CVCOD = " + sCvCod + "                                     ");
            //    strSql.AppendLine("                           UNION ALL                                                       ");
            //    strSql.AppendLine("                          SELECT ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0)");
            //    strSql.AppendLine("                                                ELSE ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) END), 0) AS JAMT    ");
            //    strSql.AppendLine("                               , A1.ACCOD                                                                    ");
            //    strSql.AppendLine("                            FROM ACTRAN A1                                                                   ");
            //    strSql.AppendLine("                            LEFT JOIN ACMSTF B1                                                              ");
            //    strSql.AppendLine("                              ON A1.ACCOD = B1.ACCOD                                                         ");
            //    strSql.AppendLine("                           WHERE A1.TDATE BETWEEN SUBSTRING('" + sYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + sYmdFrom + "')), 112)");
            //    strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'                          ");
            //    strSql.AppendLine("                             AND A1.CVCOD = " + sCvCod + "                      ");
            //    strSql.AppendLine("                             AND A1.APVYN = 'Y'                             ");
            //    strSql.AppendLine("                           GROUP BY A1.ACCOD                                ");
            //    strSql.AppendLine("                        ) A1                                                ");
            //    strSql.AppendLine("                    GROUP BY A1.ACCOD                                       ");
            //    strSql.AppendLine("                    UNION  ALL                                              ");
            //    strSql.AppendLine("                   SELECT '1' AS IDX                                        ");
            //    strSql.AppendLine("                        , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) ");
            //    strSql.AppendLine("                        , A1.SEQNO                                          ");
            //    strSql.AppendLine("                        , A1.LINNO                                          ");
            //    strSql.AppendLine("                        , B1.ACRDR                                          ");
            //    strSql.AppendLine("                        , A1.ATEXT                                          ");
            //    strSql.AppendLine("                        , ISNULL(A1.ACAMT, 0) AS ACAMT                      ");
            //    strSql.AppendLine("                        , ISNULL(A1.ADAMT, 0) AS ADAMT                      ");
            //    strSql.AppendLine("                        , 0 JAMT                                            ");
            //    strSql.AppendLine("                        , A1.ACCOD                                          ");
            //    strSql.AppendLine("                        , A1.REF1                                           ");
            //    strSql.AppendLine("                        , A1.REF2                                           ");
            //    strSql.AppendLine("                        , A1.REF3                                           ");
            //    strSql.AppendLine("                        , C1.DANJUNG                                        ");
            //    strSql.AppendLine("                        , C1.GUBUN1                                         ");
            //    strSql.AppendLine("                        , D1.J_BNUM                                         ");
            //    strSql.AppendLine("                        , C1.DANGA AS DANGA                                 ");
            //    strSql.AppendLine("                     FROM ACTRAN A1                                         ");
            //    strSql.AppendLine("                     LEFT JOIN ACMSTF B1                                    ");
            //    strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD                               ");
            //    strSql.AppendLine("                     LEFT JOIN INLIST C1                                    ");
            //    strSql.AppendLine("                       ON A1.REF1 = CONVERT(VARCHAR, C1.JUNPYOID)                             ");
            //    strSql.AppendLine("                     LEFT JOIN MESURING D1                                  ");
            //    strSql.AppendLine("                       ON C1.J_ID = CASE WHEN C1.KERATYPE = '매입' THEN D1.IPCHULGO_MAIPID ELSE D1.IPCHULGO_MACHULID END");
            //    strSql.AppendLine("                    WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //    strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'                          ");
            //    strSql.AppendLine("                      AND A1.CVCOD = " + sCvCod + "                      ");
            //    strSql.AppendLine("                      AND A1.APVYN = 'Y'                             ");
            //    strSql.AppendLine("                ) A1                                                 ");
            //    strSql.AppendLine("           LEFT OUTER JOIN ACMSTF A2                                 ");
            //    strSql.AppendLine("             ON A1.ACCOD = A2.ACCOD                                  ");
            //    strSql.AppendLine("        ) Y1                                                         ");
            //    strSql.AppendLine(" ORDER BY Y1.TDATE1, Y1.IDX, Y1.SEQNO, Y1.LINNO                      ");

            //    #region mariaDB
            //    //strSql.AppendLine(" SELECT *  ");
            //    //strSql.AppendLine("  FROM (");
            //    //strSql.AppendLine("        -- 상세 (거래처별)         ");
            //    //strSql.AppendLine("        WITH DAY_INFO AS ( ");
            //    //strSql.AppendLine("            SELECT '2' AS IDX, DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1, '[일계]' AS TDATE, A1.SEQNO, A1.LINNO, '' AS ACRDR, '' AS ATEXT, SUM(IFNULL(A1.ACAMT, 0)) AS ACAMT , SUM(IFNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD ");
            //    //strSql.AppendLine("                 , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            //    //strSql.AppendLine("                 , SUM(C1.DANJUNG) AS DANJUNG ");
            //    //strSql.AppendLine("              FROM ACTRAN A1  ");
            //    //strSql.AppendLine("              LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("              LEFT JOIN INLIST C1 ");
            //    //strSql.AppendLine("                ON A1.REF1 = C1.JUNPYOID ");
            //    //strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    //strSql.AppendLine("               AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("               AND A1.CVCOD = " + sCvCod + "  ");
            //    //strSql.AppendLine("               AND A1.APVYN = 'Y'  ");
            //    //strSql.AppendLine("             GROUP BY A1.TDATE ");
            //    //strSql.AppendLine("             UNION ALL  ");
            //    //strSql.AppendLine("            SELECT '3' AS IDX, CONCAT(DATE_FORMAT('" + sYmdTo + "', '%Y-%m-%d'), '1') AS TDATE1, '[누계]' AS TDATE, '9999' AS SEQNO, 99 AS LINNO, '' AS ACRDR, '' AS ATEXT, SUM(IFNULL(A1.ACAMT, 0)) AS ACAMT , SUM(IFNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD ");
            //    //strSql.AppendLine("                 , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            //    //strSql.AppendLine("                 , SUM(C1.DANJUNG) AS DANJUNG ");
            //    //strSql.AppendLine("              FROM ACTRAN A1  ");
            //    //strSql.AppendLine("              LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("              LEFT JOIN INLIST C1 ");
            //    //strSql.AppendLine("                ON A1.REF1 = C1.JUNPYOID ");
            //    //strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    //strSql.AppendLine("               AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("               AND A1.CVCOD = " + sCvCod + "  ");
            //    //strSql.AppendLine("               AND A1.APVYN = 'Y'  ");
            //    //strSql.AppendLine("        ) ");
            //    //strSql.AppendLine("        SELECT A1.IDX, A1.TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A1.ACNAM  ");
            //    //strSql.AppendLine("             , 0 AS JJAMT    ");
            //    //strSql.AppendLine("             , '' AS REF1 ");
            //    //strSql.AppendLine("             , '' AS REF2 ");
            //    //strSql.AppendLine("             , '' AS REF3 ");
            //    //strSql.AppendLine("             , FORMAT(A1.DANJUNG, 0) AS DANJUNG ");
            //    //strSql.AppendLine("             , '' AS GUBUN1  ");
            //    //strSql.AppendLine("             , '' AS J_BNUM ");
            //    //strSql.AppendLine("             , '' AS DANGA ");
            //    //strSql.AppendLine("          FROM DAY_INFO A1 ");
            //    //strSql.AppendLine("          UNION ALL ");
            //    //strSql.AppendLine("         SELECT A1.IDX, A1.TDATE AS TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A2.ACNAM  ");
            //    //strSql.AppendLine("              , SUM((CASE WHEN A1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0) - IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0) - IFNULL(A1.ACAMT, 0) END) + JAMT) OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) JJAMT    ");
            //    //strSql.AppendLine("              , A1.REF1 ");
            //    //strSql.AppendLine("              , A1.REF2 ");
            //    //strSql.AppendLine("              , A1.REF3 ");
            //    //strSql.AppendLine("              , FORMAT(A1.DANJUNG, 0) AS DANJUNG ");
            //    //strSql.AppendLine("              , A1.GUBUN1  ");
            //    //strSql.AppendLine("              , A1.J_BNUM ");
            //    //strSql.AppendLine("              , FORMAT(A1.DANGA, 0) AS DANGA ");
            //    //strSql.AppendLine("           FROM (SELECT '0' AS IDX, ' [이월]' TDATE, ' 'SEQNO, 0 LINNO, '' ACRDR, ' 'ATEXT, 0 ACAMT,0 ADAMT, SUM(JAMT)JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                      , '' AS REF1 ");
            //    //strSql.AppendLine("                      , '' AS REF2 ");
            //    //strSql.AppendLine("                      , '' AS REF3 ");
            //    //strSql.AppendLine("                      , '' AS DANJUNG ");
            //    //strSql.AppendLine("                      , '' AS GUBUN1 ");
            //    //strSql.AppendLine("                      , '' AS J_BNUM ");
            //    //strSql.AppendLine("                      , '' AS DANGA ");
            //    //strSql.AppendLine("                   FROM (  ");
            //    //strSql.AppendLine("                          SELECT IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                            FROM ACJANF A1   ");
            //    //strSql.AppendLine("                           WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'   ");
            //    //strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("                             AND A1.CVCOD =  " + sCvCod + "  ");
            //    //strSql.AppendLine("                           UNION ALL  ");
            //    //strSql.AppendLine("                          SELECT IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                            FROM ACTRAN A1  ");
            //    //strSql.AppendLine("                            LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                              ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("                           WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')  ");
            //    //strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("                             AND A1.CVCOD = " + sCvCod + "   ");
            //    //strSql.AppendLine("                             AND A1.APVYN = 'Y'   ");
            //    //strSql.AppendLine("                        ) A1          ");
            //    //strSql.AppendLine("                    UNION  ALL  ");
            //    //strSql.AppendLine("                   SELECT '1' AS IDX, DATE_FORMAT(A1.TDATE, '%Y-%m-%d'), A1.SEQNO, A1.LINNO, B1.ACRDR, A1.ATEXT, IFNULL(A1.ACAMT, 0) AS ACAMT , IFNULL(A1.ADAMT, 0) AS ADAMT, 0 JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                        , A1.REF1 ");
            //    //strSql.AppendLine("                        , A1.REF2 ");
            //    //strSql.AppendLine("                        , A1.REF3 ");
            //    //strSql.AppendLine("                        , C1.DANJUNG ");
            //    //strSql.AppendLine("                        , C1.GUBUN1 ");
            //    //strSql.AppendLine("                        , D1.J_BNUM ");
            //    //strSql.AppendLine("                        , C1.DANGA AS DANGA ");
            //    //strSql.AppendLine("                     FROM ACTRAN A1  ");
            //    //strSql.AppendLine("                     LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("                     LEFT JOIN INLIST C1 ");
            //    //strSql.AppendLine("                       ON A1.REF1 = C1.JUNPYOID ");
            //    //strSql.AppendLine("                     LEFT JOIN MESURING D1 ");
            //    //strSql.AppendLine("                       ON C1.J_ID = CASE WHEN C1.KERATYPE = '매입' THEN D1.IPCHULGO_MAIPID ELSE D1.IPCHULGO_MACHULID END ");
            //    //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    //strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("                      AND A1.CVCOD = " + sCvCod + "  ");
            //    //strSql.AppendLine("                      AND A1.APVYN = 'Y'  ");
            //    //strSql.AppendLine("                ) A1  ");
            //    //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF A2  ");
            //    //strSql.AppendLine("             ON A1.ACCOD = A2.ACCOD  ");
            //    //strSql.AppendLine("        ) Y1 ");
            //    //strSql.AppendLine(" ORDER BY Y1.TDATE1, Y1.IDX, Y1.SEQNO, Y1.LINNO ");
            //    #endregion
            //}
            //else
            //{
            //    //외상매입금 및 외상매출금의 경우 추가되는 컬럼 Visible true;
            //    foreach (GridColumn col in ArrGridCol)
            //    {
            //        col.Visible = false;
            //    }

            //    /*
            //     * 2021-01-04 현업요청
            //     * 계정상세내역 중 일계 및 누계 추가요청에 따라 쿼리 수정
            //     */
            //    strSql.Clear();

            //    strSql.AppendLine("--상세(거래처별)                                                        ");
            //    strSql.AppendLine("WITH DAY_INFO AS(                                                       ");
            //    strSql.AppendLine("    SELECT '2' AS IDX                                                   ");
            //    strSql.AppendLine("         , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23) AS TDATE1  ");
            //    strSql.AppendLine("         , '[일계]' AS TDATE                                            ");
            //    strSql.AppendLine("         , MIN(A1.SEQNO) AS SEQNO                                       ");
            //    strSql.AppendLine("         , MIN(A1.LINNO) AS LINNO                                       ");
            //    strSql.AppendLine("         , '' AS ACRDR                                                  ");
            //    strSql.AppendLine("         , '' AS ATEXT                                                  ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ACAMT, 0)) AS ACAMT                            ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ADAMT, 0)) AS ADAMT                            ");
            //    strSql.AppendLine("         , '' AS ACCOD                                                  ");
            //    strSql.AppendLine("         , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT                           ");
            //    strSql.AppendLine("      FROM ACTRAN A1                                                    ");
            //    strSql.AppendLine("      LEFT JOIN ACMSTF B1                                               ");
            //    strSql.AppendLine("        ON A1.ACCOD = B1.ACCOD                                          ");
            //    strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    strSql.AppendLine("               AND A1.ACCOD = '" + sAcCod + "'   ");
            //    strSql.AppendLine("               AND A1.CVCOD = " + sCvCod + "  ");
            //    strSql.AppendLine("       AND A1.APVYN = 'Y'                                               ");
            //    strSql.AppendLine("     GROUP BY A1.TDATE                                                  ");
            //    strSql.AppendLine("     UNION ALL                                                          ");
            //    strSql.AppendLine("    SELECT '3' AS IDX");
            //    strSql.AppendLine("         , CONCAT(CONVERT(VARCHAR(10), CONVERT(DATE, '" + sYmdTo + "'), 23), '1') AS TDATE1");
            //    strSql.AppendLine("         , '[누계]' AS TDATE                  ");
            //    strSql.AppendLine("         , '9999' AS SEQNO                    ");
            //    strSql.AppendLine("         , 99 AS LINNO                        ");
            //    strSql.AppendLine("         , '' AS ACRDR                        ");
            //    strSql.AppendLine("         , '' AS ATEXT                        ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ACAMT, 0)) AS ACAMT  ");
            //    strSql.AppendLine("         , SUM(ISNULL(A1.ADAMT, 0)) AS ADAMT  ");
            //    strSql.AppendLine("         , '' AS ACCOD                        ");
            //    strSql.AppendLine("         , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            //    strSql.AppendLine("      FROM ACTRAN A1                          ");
            //    strSql.AppendLine("      LEFT JOIN ACMSTF B1                     ");
            //    strSql.AppendLine("        ON A1.ACCOD = B1.ACCOD                ");
            //    strSql.AppendLine("     WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
            //    strSql.AppendLine("       AND A1.ACCOD = '" + sAcCod + "'        ");
            //    strSql.AppendLine("       AND A1.CVCOD = " + sCvCod + "          ");
            //    strSql.AppendLine("       AND A1.APVYN = 'Y'           ");
            //    strSql.AppendLine(")                                   ");
            //    strSql.AppendLine("SELECT*                             ");
            //    strSql.AppendLine("  FROM(                             ");
            //    strSql.AppendLine("        SELECT A1.IDX               ");
            //    strSql.AppendLine("             , A1.TDATE1            ");
            //    strSql.AppendLine("             , A1.TDATE             ");
            //    strSql.AppendLine("             , A1.SEQNO             ");
            //    strSql.AppendLine("             , A1.LINNO             ");
            //    strSql.AppendLine("             , A1.ATEXT             ");
            //    strSql.AppendLine("             , ACAMT                ");
            //    strSql.AppendLine("             , ADAMT                ");
            //    strSql.AppendLine("             , A1.ACCOD             ");
            //    strSql.AppendLine("             , A1.ACNAM             ");
            //    strSql.AppendLine("             , 0 AS JJAMT           ");
            //    strSql.AppendLine("          FROM DAY_INFO A1          ");
            //    strSql.AppendLine("          UNION ALL                 ");
            //    strSql.AppendLine("         SELECT A1.IDX              ");
            //    strSql.AppendLine("              , A1.TDATE AS TDATE1  ");
            //    strSql.AppendLine("              , A1.TDATE            ");
            //    strSql.AppendLine("              , A1.SEQNO            ");
            //    strSql.AppendLine("              , A1.LINNO            ");
            //    strSql.AppendLine("              , A1.ATEXT            ");
            //    strSql.AppendLine("              , ACAMT               ");
            //    strSql.AppendLine("              , ADAMT               ");
            //    strSql.AppendLine("              , A1.ACCOD            ");
            //    strSql.AppendLine("              , A2.ACNAM            ");
            //    strSql.AppendLine("              , SUM((CASE WHEN A1.ACRDR = '1' THEN ISNULL(A1.ACAMT, 0) - ISNULL(A1.ADAMT, 0) ELSE ISNULL(A1.ADAMT, 0) - ISNULL(A1.ACAMT, 0) END) + JAMT) OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) JJAMT");
            //    strSql.AppendLine("           FROM(SELECT '0' AS IDX         ");
            //    strSql.AppendLine("                      , ' [이월]' TDATE   ");
            //    strSql.AppendLine("                      , ' 'SEQNO          ");
            //    strSql.AppendLine("                      , 0 LINNO           ");
            //    strSql.AppendLine("                      , '' ACRDR          ");
            //    strSql.AppendLine("                      , ' 'ATEXT          ");
            //    strSql.AppendLine("                      , 0 ACAMT           ");
            //    strSql.AppendLine("                      , 0 ADAMT           ");
            //    strSql.AppendLine("                      , SUM(JAMT)JAMT     ");
            //    strSql.AppendLine("                      , A1.ACCOD          ");
            //    strSql.AppendLine("                   FROM(                  ");
            //    strSql.AppendLine("                          SELECT ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT ");
            //    strSql.AppendLine("                               , A1.ACCOD                                         ");
            //    strSql.AppendLine("                            FROM ACJANF A1                                        ");
            //    strSql.AppendLine("                           WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'                               ");
            //    strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'                                ");
            //    strSql.AppendLine("                             AND A1.CVCOD = " + sCvCod + "                                  ");
            //    strSql.AppendLine("                           UNION ALL                                              ");
            //    strSql.AppendLine("                          SELECT ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) END), 0) AS JAMT");
            //    strSql.AppendLine("                               , A1.ACCOD            ");
            //    strSql.AppendLine("                            FROM ACTRAN A1           ");
            //    strSql.AppendLine("                            LEFT JOIN ACMSTF B1      ");
            //    strSql.AppendLine("                              ON A1.ACCOD = B1.ACCOD ");
            //    strSql.AppendLine("                           WHERE A1.TDATE BETWEEN SUBSTRING('" + sYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + sYmdFrom + "')), 112)");
            //    strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'  ");
            //    strSql.AppendLine("                             AND A1.CVCOD = " + sCvCod + "    ");
            //    strSql.AppendLine("                             AND A1.APVYN = 'Y'     ");
            //    strSql.AppendLine("                           GROUP BY A1.ACCOD        ");
            //    strSql.AppendLine("                        ) A1                        ");
            //    strSql.AppendLine("                    GROUP BY A1.ACCOD               ");
            //    strSql.AppendLine("                    UNION  ALL                      ");
            //    strSql.AppendLine("                   SELECT '1' AS IDX                ");
            //    strSql.AppendLine("                        , CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23)");
            //    strSql.AppendLine("                        , A1.SEQNO                      ");
            //    strSql.AppendLine("                        , A1.LINNO                      ");
            //    strSql.AppendLine("                        , B1.ACRDR                      ");
            //    strSql.AppendLine("                        , A1.ATEXT                      ");
            //    strSql.AppendLine("                        , ISNULL(A1.ACAMT, 0) AS ACAMT  ");
            //    strSql.AppendLine("                        , ISNULL(A1.ADAMT, 0) AS ADAMT  ");
            //    strSql.AppendLine("                        , 0 JAMT                        ");
            //    strSql.AppendLine("                        , A1.ACCOD                      ");
            //    strSql.AppendLine("                     FROM ACTRAN A1                     ");
            //    strSql.AppendLine("                     LEFT JOIN ACMSTF B1                ");
            //    strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD           ");
            //    strSql.AppendLine("                    WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
            //    strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'      ");
            //    strSql.AppendLine("                      AND A1.CVCOD = " + sCvCod + "        ");
            //    strSql.AppendLine("                      AND A1.APVYN = 'Y'         ");
            //    strSql.AppendLine("                ) A1                             ");
            //    strSql.AppendLine("           LEFT OUTER JOIN ACMSTF A2             ");
            //    strSql.AppendLine("             ON A1.ACCOD = A2.ACCOD              ");
            //    strSql.AppendLine("        ) Y1                                     ");
            //    strSql.AppendLine(" ORDER BY Y1.TDATE1, Y1.IDX, Y1.SEQNO, Y1.LINNO  ");

            //    #region mariaDB
            //    //strSql.AppendLine(" SELECT *  ");
            //    //strSql.AppendLine("  FROM (");
            //    //strSql.AppendLine("        -- 상세 (거래처별)         ");
            //    //strSql.AppendLine("        WITH DAY_INFO AS ( ");
            //    //strSql.AppendLine("            SELECT '2' AS IDX, DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1, '[일계]' AS TDATE, A1.SEQNO, A1.LINNO, '' AS ACRDR, '' AS ATEXT, SUM(IFNULL(A1.ACAMT, 0)) AS ACAMT , SUM(IFNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD ");
            //    //strSql.AppendLine("                 , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            //    //strSql.AppendLine("              FROM ACTRAN A1  ");
            //    //strSql.AppendLine("              LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    //strSql.AppendLine("               AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("               AND A1.CVCOD = " + sCvCod + "  ");
            //    //strSql.AppendLine("               AND A1.APVYN = 'Y'  ");
            //    //strSql.AppendLine("             GROUP BY A1.TDATE ");
            //    //strSql.AppendLine("             UNION ALL  ");
            //    //strSql.AppendLine("            SELECT '3' AS IDX, CONCAT(DATE_FORMAT('" + sYmdTo + "', '%Y-%m-%d'), '1') AS TDATE1, '[누계]' AS TDATE, '9999' AS SEQNO, 99 AS LINNO, '' AS ACRDR, '' AS ATEXT, SUM(IFNULL(A1.ACAMT, 0)) AS ACAMT , SUM(IFNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD ");
            //    //strSql.AppendLine("                 , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            //    //strSql.AppendLine("              FROM ACTRAN A1  ");
            //    //strSql.AppendLine("              LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    //strSql.AppendLine("               AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("               AND A1.CVCOD = " + sCvCod + "  ");
            //    //strSql.AppendLine("               AND A1.APVYN = 'Y'  ");
            //    //strSql.AppendLine("        ) ");
            //    //strSql.AppendLine("        SELECT A1.IDX, A1.TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A1.ACNAM  ");
            //    //strSql.AppendLine("             , 0 AS JJAMT    ");
            //    //strSql.AppendLine("          FROM DAY_INFO A1 ");
            //    //strSql.AppendLine("          UNION ALL ");
            //    //strSql.AppendLine("         SELECT A1.IDX, A1.TDATE AS TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A2.ACNAM  ");
            //    //strSql.AppendLine("              , SUM((CASE WHEN A1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0) - IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0) - IFNULL(A1.ACAMT, 0) END) + JAMT) OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) JJAMT    ");
            //    //strSql.AppendLine("           FROM (SELECT '0' AS IDX, ' [이월]' TDATE, ' 'SEQNO, 0 LINNO, '' ACRDR, ' 'ATEXT, 0 ACAMT,0 ADAMT, SUM(JAMT)JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                   FROM (  ");
            //    //strSql.AppendLine("                          SELECT IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                            FROM ACJANF A1   ");
            //    //strSql.AppendLine("                           WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'   ");
            //    //strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("                             AND A1.CVCOD =  " + sCvCod + "  ");
            //    //strSql.AppendLine("                           UNION ALL  ");
            //    //strSql.AppendLine("                          SELECT IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                            FROM ACTRAN A1  ");
            //    //strSql.AppendLine("                            LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                              ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("                           WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')  ");
            //    //strSql.AppendLine("                             AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("                             AND A1.CVCOD = " + sCvCod + "   ");
            //    //strSql.AppendLine("                             AND A1.APVYN = 'Y'   ");
            //    //strSql.AppendLine("                        ) A1          ");
            //    //strSql.AppendLine("                    UNION  ALL  ");
            //    //strSql.AppendLine("                   SELECT '1' AS IDX, DATE_FORMAT(A1.TDATE, '%Y-%m-%d'), A1.SEQNO, A1.LINNO, B1.ACRDR, A1.ATEXT, IFNULL(A1.ACAMT, 0) AS ACAMT , IFNULL(A1.ADAMT, 0) AS ADAMT, 0 JAMT, A1.ACCOD  ");
            //    //strSql.AppendLine("                     FROM ACTRAN A1  ");
            //    //strSql.AppendLine("                     LEFT JOIN ACMSTF B1   ");
            //    //strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD  ");
            //    //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //    //strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'   ");
            //    //strSql.AppendLine("                      AND A1.CVCOD = " + sCvCod + "  ");
            //    //strSql.AppendLine("                      AND A1.APVYN = 'Y'  ");
            //    //strSql.AppendLine("                ) A1  ");
            //    //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF A2  ");
            //    //strSql.AppendLine("             ON A1.ACCOD = A2.ACCOD  ");
            //    //strSql.AppendLine("        ) Y1 ");
            //    //strSql.AppendLine(" ORDER BY Y1.TDATE1, Y1.IDX, Y1.SEQNO, Y1.LINNO ");
            //    #endregion
            //}

            //#region[2021-02-04 이전쿼리]

            /////*
            //// * 2021-01-04 현업요청
            //// * 계정상세내역 중 일계 및 누계 추가요청에 따라 쿼리 수정
            //// */
            ////strSql.Clear();
            ////strSql.AppendLine(" SELECT *  ");
            ////strSql.AppendLine("  FROM ( ");
            ////strSql.AppendLine("         -- 상세 (거래처별)         ");
            ////strSql.AppendLine("         WITH DAY_INFO AS ( ");
            ////strSql.AppendLine("             SELECT '2' AS IDX, DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE1, '[일계]' AS TDATE, A1.SEQNO, A1.LINNO, '' AS ACRDR, '' AS ATEXT, SUM(IFNULL(A1.ACAMT, 0)) AS ACAMT , SUM(IFNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD ");
            ////strSql.AppendLine("                  , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            ////strSql.AppendLine("               FROM ACTRAN A1  ");
            ////strSql.AppendLine("               LEFT JOIN ACMSTF B1   ");
            ////strSql.AppendLine("                 ON A1.ACCOD = B1.ACCOD  ");
            ////strSql.AppendLine("              WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            ////strSql.AppendLine("                AND A1.ACCOD = '" + sAcCod + "'   ");
            ////strSql.AppendLine("                AND A1.CVCOD = " + sCvCod + "  ");
            ////strSql.AppendLine("                AND A1.APVYN = 'Y'  ");
            ////strSql.AppendLine("              GROUP BY A1.TDATE ");
            ////strSql.AppendLine("              UNION ALL  ");
            ////strSql.AppendLine("             SELECT '3' AS IDX, CONCAT(DATE_FORMAT('" + sYmdTo + "', '%Y-%m-%d'), '1') AS TDATE1, '[누계]' AS TDATE, '9999' AS SEQNO, 99 AS LINNO, '' AS ACRDR, '' AS ATEXT, SUM(IFNULL(A1.ACAMT, 0)) AS ACAMT , SUM(IFNULL(A1.ADAMT, 0)) AS ADAMT, '' AS ACCOD ");
            ////strSql.AppendLine("                  , '' AS ACNAM, 0 AS JAMT, 0 AS JJAMT ");
            ////strSql.AppendLine("               FROM ACTRAN A1  ");
            ////strSql.AppendLine("               LEFT JOIN ACMSTF B1   ");
            ////strSql.AppendLine("                 ON A1.ACCOD = B1.ACCOD  ");
            ////strSql.AppendLine("              WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            ////strSql.AppendLine("                AND A1.ACCOD = '" + sAcCod + "'   ");
            ////strSql.AppendLine("                AND A1.CVCOD = " + sCvCod + "  ");
            ////strSql.AppendLine("                AND A1.APVYN = 'Y'  ");
            ////strSql.AppendLine(" ) ");
            ////strSql.AppendLine(" SELECT A1.IDX, A1.TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A1.ACNAM  ");
            ////strSql.AppendLine("      , 0 AS JJAMT    ");
            ////strSql.AppendLine("   FROM DAY_INFO A1 ");
            ////strSql.AppendLine("   UNION ALL ");
            ////strSql.AppendLine("  SELECT A1.IDX, A1.TDATE AS TDATE1, A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A2.ACNAM  ");
            ////strSql.AppendLine("       , SUM((CASE WHEN A1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0) - IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0) - IFNULL(A1.ACAMT, 0) END) + JAMT) OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) JJAMT    ");
            ////strSql.AppendLine("    FROM (SELECT '0' AS IDX, ' [이월]' TDATE, ' 'SEQNO, 0 LINNO, '' ACRDR, ' 'ATEXT, 0 ACAMT,0 ADAMT, SUM(JAMT)JAMT, A1.ACCOD  ");
            ////strSql.AppendLine("            FROM (  ");
            ////strSql.AppendLine("                   SELECT IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT, A1.ACCOD  ");
            ////strSql.AppendLine("                     FROM ACJANF A1   ");
            ////strSql.AppendLine("                    WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'   ");
            ////strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'   ");
            ////strSql.AppendLine("                      AND A1.CVCOD =  " + sCvCod + "  ");
            ////strSql.AppendLine("                    UNION ALL  ");
            ////strSql.AppendLine("                   SELECT IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT, A1.ACCOD  ");
            ////strSql.AppendLine("                     FROM ACTRAN A1  ");
            ////strSql.AppendLine("                     LEFT JOIN ACMSTF B1   ");
            ////strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD  ");
            ////strSql.AppendLine("                    WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')  ");
            ////strSql.AppendLine("                      AND A1.ACCOD = '" + sAcCod + "'   ");
            ////strSql.AppendLine("                      AND A1.CVCOD = " + sCvCod + "   ");
            ////strSql.AppendLine("                      AND A1.APVYN = 'Y'   ");
            ////strSql.AppendLine("                 ) A1          ");
            ////strSql.AppendLine("             UNION  ALL  ");
            ////strSql.AppendLine("            SELECT '1' AS IDX, DATE_FORMAT(A1.TDATE, '%Y-%m-%d'), A1.SEQNO, A1.LINNO, B1.ACRDR, A1.ATEXT, IFNULL(A1.ACAMT, 0) AS ACAMT , IFNULL(A1.ADAMT, 0) AS ADAMT, 0 JAMT, A1.ACCOD  ");
            ////strSql.AppendLine("              FROM ACTRAN A1  ");
            ////strSql.AppendLine("              LEFT JOIN ACMSTF B1   ");
            ////strSql.AppendLine("                ON A1.ACCOD = B1.ACCOD  ");
            ////strSql.AppendLine("             WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            ////strSql.AppendLine("               AND A1.ACCOD = '" + sAcCod + "'   ");
            ////strSql.AppendLine("               AND A1.CVCOD = " + sCvCod + "  ");
            ////strSql.AppendLine("               AND A1.APVYN = 'Y'  ");
            ////strSql.AppendLine("         ) A1  ");
            ////strSql.AppendLine("    LEFT OUTER JOIN ACMSTF A2  ");
            ////strSql.AppendLine("      ON A1.ACCOD = A2.ACCOD  ");
            ////strSql.AppendLine(" ) Y1 ");
            ////strSql.AppendLine(" ORDER BY Y1.TDATE1, Y1.IDX, Y1.SEQNO, Y1.LINNO ");

            //#endregion[2021-02-04 이전쿼리]

            //#region[2021-01-14 이전 쿼리]

            ////strSql.Clear();
            ////strSql.AppendLine(" -- 상세 (거래처별)        ");
            ////strSql.AppendLine(" SELECT A1.TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, ACAMT, ADAMT, A1.ACCOD, A2.ACNAM ");
            ////strSql.AppendLine("      , SUM((CASE WHEN A1.ACRDR='1' THEN IFNULL(A1.ACAMT, 0) - IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0) - IFNULL(A1.ACAMT, 0) END) + JAMT) OVER(ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO) JJAMT   ");
            ////strSql.AppendLine("   FROM (SELECT ' [이월]' TDATE, ' 'SEQNO, 0 LINNO, '' ACRDR, ' 'ATEXT, 0 ACAMT,0 ADAMT, SUM(JAMT)JAMT, A1.ACCOD ");
            ////strSql.AppendLine("           FROM ( ");
            ////strSql.AppendLine("                  SELECT IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT, A1.ACCOD ");
            ////strSql.AppendLine("                    FROM ACJANF A1  ");
            ////strSql.AppendLine("                   WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'  ");
            ////strSql.AppendLine("                     AND A1.ACCOD = '" + sAcCod + "'  ");
            ////strSql.AppendLine("                     AND A1.CVCOD =  " + sCvCod + " ");
            ////strSql.AppendLine("                   UNION ALL ");
            ////strSql.AppendLine("                  SELECT IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT, A1.ACCOD ");
            ////strSql.AppendLine("                    FROM ACTRAN A1 ");
            ////strSql.AppendLine("                    LEFT JOIN ACMSTF B1  ");
            ////strSql.AppendLine("                      ON A1.ACCOD = B1.ACCOD ");
            ////strSql.AppendLine("                   WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d') ");
            ////strSql.AppendLine("                     AND A1.ACCOD = '" + sAcCod + "'  ");
            ////strSql.AppendLine("                     AND A1.CVCOD = " + sCvCod + "  ");
            ////strSql.AppendLine("                     AND A1.APVYN = 'Y'  ");
            ////strSql.AppendLine("                ) A1         ");
            ////strSql.AppendLine("            UNION  ALL ");
            ////strSql.AppendLine("           SELECT DATE_FORMAT(A1.TDATE, '%Y-%m-%d'), A1.SEQNO, A1.LINNO, B1.ACRDR, A1.ATEXT, IFNULL(A1.ACAMT, 0) AS ACAMT , IFNULL(A1.ADAMT, 0) AS ADAMT, 0 JAMT, A1.ACCOD ");
            ////strSql.AppendLine("             FROM ACTRAN A1 ");
            ////strSql.AppendLine("             LEFT JOIN ACMSTF B1  ");
            ////strSql.AppendLine("               ON A1.ACCOD = B1.ACCOD ");
            ////strSql.AppendLine("            WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            ////strSql.AppendLine("              AND A1.ACCOD = '" + sAcCod + "'  ");
            ////strSql.AppendLine("              AND A1.CVCOD = " + sCvCod + " ");
            ////strSql.AppendLine("              AND A1.APVYN = 'Y' ");
            ////strSql.AppendLine("        ) A1 ");
            ////strSql.AppendLine("   LEFT OUTER JOIN ACMSTF A2 ");
            ////strSql.AppendLine("     ON A1.ACCOD = A2.ACCOD ");


            //#endregion[2021-01-14 이전 쿼리]

            //return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            #endregion
            //#0002
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            /*
             * 2021-02-04(현업요청)
             * 계정상세내역에서 마감자료와 JOIN걸어 등급, 차번, 인수량을 가져오도록 수정
             * 대상은 외상매입금, 외상매출금을 대상으로 하여 계정코드가 해당 코드일 경우 분기를 태운다
             */
            if (sAcCod.Equals("0251") || sAcCod.Equals("0108"))
            {
                //외상매입금 및 외상매출금의 경우 추가되는 컬럼 Visible true;
                foreach (GridColumn col in ArrGridCol)
                {
                    col.Visible = true;
                }

                dicParams.Add("CMD", "LIST5");
            }
            else
            {
                //외상매입금 및 외상매출금의 경우 추가되는 컬럼 Visible false;
                foreach (GridColumn col in ArrGridCol)
                {
                    col.Visible = false;
                }

                dicParams.Add("CMD", "LIST4");
            }
                
            dicParams.Add("DATEF", sYmdFrom);
            dicParams.Add("DATET", sYmdTo);
            dicParams.Add("CVCOD", sCvCod);
            dicParams.Add("ACCOD", sAcCod);

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }
        #endregion[Query]

        #region[GridView's Design]

        private void GridViewL1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewL1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewR2_RowStyle(object sender, RowStyleEventArgs e)
        {
            #region #0002 이전
            ///*
            // * 2021-01-14
            // * 일계 및 누계는 색깔 달리함 (IDX : 2, 3)
            // * 쿼리참조
            // */
            //string sIDX = GridViewR2.GetRowCellValue(e.RowHandle, GridColIDX)?.ToString();
            //if(!string.IsNullOrEmpty(sIDX) && (sIDX.Equals("2") || sIDX.Equals("3") || sIDX.Equals("0")))
            //{
            //    e.Appearance.BackColor = Color.PaleGreen;
            //}
            //else if (e.RowHandle % 2 == 0)
            //{
            //    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            //}
            #endregion

            /*
             * #0002
             * 이월,일계, 누계 색 다르게
             * 
             * #0003
             * 월계 추가
             */
            string sTDATE = GridViewR2.GetRowCellValue(e.RowHandle, GridColR2TDate)?.ToString();
            if (!string.IsNullOrEmpty(sTDATE) && (sTDATE.Equals("[이월]") || sTDATE.Equals("[일계]") || sTDATE.Equals("[월계]") || sTDATE.Equals("[누계]")))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        #endregion[GridView's Design]

        public DataRow DrPopupInfo;
        private void BtnEditAcCod_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString();

            AC01001F03 frm = new AC01001F03();
            frm.P_AC09001F01 = this;
            frm.AccCd = sVal;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                btnEdit.EditValue = DrPopupInfo["ACCOD"];
                TxtAcNam.EditValue = DrPopupInfo["ACNAM"];
            }
        }

        private void BtnEditAcCod_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                //계정코드는 계정상세 탭에서 필수 입력사항으로 무조건 띄움
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sVal = btnEdit.EditValue?.ToString();
                //if (string.IsNullOrEmpty(sVal))
                //{
                //    TxtAcNam.EditValue = string.Empty;
                //    return;
                //}

                DataTable dt = GetAccInfo(sVal);
                if(dt.Rows.Count == 1)
                {
                    btnEdit.EditValue = dt.Rows[0]["ACCOD"];
                    TxtAcNam.EditValue = dt.Rows[0]["ACNAM"];
                }
                else
                {
                    BtnEditAcCod_ButtonClick(sender, null);
                }

            }
        }

        private DataTable GetAccInfo(string sVal)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT ACCOD ");
            strSql.AppendLine("      , ACNAM ");
            strSql.AppendLine("   FROM ACMSTF ");
            strSql.AppendLine("  WHERE ACCOD = '" + sVal + "' ");
            strSql.AppendLine("     OR ACNAM LIKE '%" + sVal + "%' ");
            
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void SetVisibility(string sGb)
        {
            if (sGb.Equals("0"))
            {
                LayoutAcCod.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutTxtAcNam.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutCarryOver.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else
            {
                LayoutAcCod.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutTxtAcNam.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutCarryOver.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
        }
        
        private void XtTab_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            SetVisibility(XtTab.SelectedTabPageIndex.ToString());
        }

        private void GridViewR1_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                string sCvCod = GridViewL1.GetFocusedRowCellValue("CVCOD")?.ToString();
                string sAcCod = GridViewR1.GetFocusedRowCellValue("ACCOD")?.ToString();
                string sAcNam = GridViewR1.GetFocusedRowCellValue("ACNAM")?.ToString();
                string sCarryOverAmtGb = RdgbCarryOver.EditValue?.ToString();
                XtTab.SelectedTabPageIndex = 1;
                BtnEditAcCod.EditValue = sAcCod;
                TxtAcNam.EditValue = sAcNam;
                //BtnRetr_Click(null, null);

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                GridL2.DataSource = null;
                GridL2.DataSource = GetDealerInfoDetail(sYmdFrom, sYmdTo, sAcCod, sCarryOverAmtGb, sCvCod);
            }
        }

        private void AC09001F01_TextChanged(object sender, EventArgs e)
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

        public DataRow DrDealerInfo;
        private void BtnCvCod_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            AC02001F03 frm = new AC02001F03();
            frm.P_AC09001F01 = this;
            frm.DealerCd = string.Empty;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                btnEdit.EditValue = DrDealerInfo["DEALER_CD"];
                TxtCvNam.EditValue = DrDealerInfo["DEALER_NM"];
            }
        }

        private DataTable GetDealerInfo(string sVal)
        {
            StringBuilder strSql = new StringBuilder();

            /*
             * 수정일자 : 2021-02-07 (현업요청)
             * 수정자 : 고혜성
             * 수정내용 : 거래처초성검색 추가
             */
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.DEALER_CD  ");
            strSql.AppendLine("      , A.DEALER_NM ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A ");
            strSql.AppendLine("  WHERE A.DEALER_NM LIKE '%" + sVal + "%' OR A.INITIAL_NM LIKE '%" + sVal + "%' ");
            if (double.TryParse(sVal, out double result))
            {
                strSql.AppendLine("OR A.DEALER_CD = " + sVal);
            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        
        private void BtnCvCod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sVal = btnEdit.EditValue?.ToString().Trim();
                if (string.IsNullOrEmpty(sVal))
                {
                    TxtCvNam.EditValue = string.Empty;
                    return;
                }

                DataTable dt = GetDealerInfo(sVal);
                if (dt != null)
                {
                    if (dt.Rows.Count == 1)
                    {
                        btnEdit.EditValue = dt.Rows[0]["DEALER_CD"];
                        TxtCvNam.EditValue = dt.Rows[0]["DEALER_NM"];
                    }
                    else
                    {
                        AC02001F03 frm = new AC02001F03();
                        frm.P_AC09001F01 = this;
                        frm.DealerCd = sVal;
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            btnEdit.EditValue = DrDealerInfo["DEALER_CD"];
                            TxtCvNam.EditValue = DrDealerInfo["DEALER_NM"];
                        }
                    }
                }
            }
        }

        private void RdgbCarryOver_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewL1_MouseDown(object sender, MouseEventArgs e)
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

        private void DropBtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "계정집계리스트")
            {
                AcCodSummary();
            }
            else if (tag == "거래처별 계정잔액리스트")
            {
                DealerCarryOver();
            }
            else if (tag == "계정상세리스트")
            {
                AcCodList();
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnAcCodSummary;
        BarButtonItem BtnDealerCarryOver;
        BarButtonItem BtnAcCodList;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnAcCodSummary = new BarButtonItem(barManager1, "계정집계리스트");
            BtnDealerCarryOver = new BarButtonItem(barManager1, "거래처별 계정잔액리스트");
            BtnAcCodList = new BarButtonItem(barManager1, "계정상세리스트");
            popupMenu1.AddItem(BtnAcCodSummary);
            popupMenu1.AddItem(BtnDealerCarryOver);
            popupMenu1.AddItem(BtnAcCodList);

            DropBtnExcel.DropDownControl = popupMenu1;

            BtnAcCodSummary.Tag = "계정집계리스트";
            BtnAcCodSummary.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnAcCodSummary_ItemClick);
            
            BtnDealerCarryOver.Tag = "거래처별 계정잔액리스트";
            BtnDealerCarryOver.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnDealerCarryOver_ItemClick);

            BtnAcCodList.Tag = "계정상세리스트";
            BtnAcCodList.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnAcCodList_ItemClick);
        }

        private void BtnAcCodSummary_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            AcCodSummary();
        }

        private void BtnDealerCarryOver_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            DealerCarryOver();
        }

        private void BtnAcCodList_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            AcCodList();
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnExcel.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnExcel.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnExcel.Tag = submenuItem.Tag;
        }

        private void AcCodSummary()
        {
            if(XtTab.SelectedTabPageIndex != 0)
            {
                XtraMessageBox.Show("'거래처별 계정집계' 탭을 클릭하세요.");
                return;
            }

            if(GridViewR1.RowCount == 0)
            {
                XtraMessageBox.Show("계정집계 리스트(오른쪽 리스트)에 데이터가 없습니다.");
                return;
            }
            
            ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}", this.Text, "_계정집계리스트"), GridR1);
        }

        private void DealerCarryOver()
        {
            if (XtTab.SelectedTabPageIndex != 1)
            {
                XtraMessageBox.Show("'거래처별 계정상세' 탭을 클릭하세요.");
                return;
            }

            if (GridViewL2.RowCount == 0)
            {
                XtraMessageBox.Show("거래처별 계정잔액 리스트(왼쪽 리스트) 에 데이터가 없습니다.");
                return;
            }
            ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}", this.Text, "_거래처별 계정잔액리스트"), GridL2);
        }

        private void AcCodList()
        {
            if (XtTab.SelectedTabPageIndex != 1)
            {
                XtraMessageBox.Show("'거래처별 계정상세' 탭을 클릭하세요.");
                return;
            }

            if (GridViewR2.RowCount == 0)
            {
                XtraMessageBox.Show("계정상세 리스트(오른쪽 리스트) 에 데이터가 없습니다.");
                return;
            }
            //ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}", this.Text, "계정상세리스트"), GridR2);

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sAcCod = BtnEditAcCod.EditValue?.ToString().Trim();
            string sCvCod = GridViewL2.GetFocusedRowCellValue("CVCOD")?.ToString();

            ReportViewer fm = new ReportViewer(GetAccountDetailOfDealersPRINT(sYmdFrom, sYmdTo, sCvCod, sAcCod), GetAccountDetailOfDealersPRINT(sYmdFrom, sYmdTo, sCvCod, sAcCod), "RptAcc");
            fm.ShowDialog();
        }
        private DataTable GetAccountDetailOfDealersPRINT(string sYmdFrom, string sYmdTo, string sCvCod, string sAcCod)
        {
            //#0002
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            /*
             * 2021-02-04(현업요청)
             * 계정상세내역에서 마감자료와 JOIN걸어 등급, 차번, 인수량을 가져오도록 수정
             * 대상은 외상매입금, 외상매출금을 대상으로 하여 계정코드가 해당 코드일 경우 분기를 태운다
             */
            if (sAcCod.Equals("0251") || sAcCod.Equals("0108"))
            {
                //외상매입금 및 외상매출금의 경우 추가되는 컬럼 Visible true;
                foreach (GridColumn col in ArrGridCol)
                {
                    col.Visible = true;
                }

                dicParams.Add("CMD", "LIST5PRINT");
            }
            else
            {
                //외상매입금 및 외상매출금의 경우 추가되는 컬럼 Visible false;
                foreach (GridColumn col in ArrGridCol)
                {
                    col.Visible = false;
                }

                dicParams.Add("CMD", "LIST4PRINT");
            }

            dicParams.Add("DATEF", sYmdFrom);
            dicParams.Add("DATET", sYmdTo);
            dicParams.Add("CVCOD", sCvCod);
            dicParams.Add("ACCOD", sAcCod);

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }
    }
}