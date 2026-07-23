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
using DevExpress.XtraGrid;
using ComLib;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

/*
 * 작성일자 : 2021-02-09
 * 작성자 : 고혜성
 * ------------------------------HISTORY-------------------------
 * 수정일자 : 2021-02-23
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            INLIST에 이력성 컬럼인 CHRG_ID(담당자)로 업체담당자 표기하도록 쿼리 수정
 *            
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자: 2021-09-10
 * 수정자: 정은영
 * ID     : #00001
 * 수정내용 : (현업요청)
 *           1. 매출처별 조회시 전표에서 인센티브 데이터도 가져와 매출리스트에 표기 (등급:인센티브 , 금액/도착도금액: 전표의 공급가액, 나머지 숫자는 0처리)
 *           
 * 수정일자: 2022-12-26
 * 수정자  : 정은영
 * ID      : #00002
 * 수정내용: (현업요청)
 *           1. 인센티브 전표 승인자료만
 */
namespace AccAdm
{
    public partial class IN010F00 : DevExpress.XtraEditors.XtraForm
    {
        public IN010F00()
        {
            InitializeComponent();
        }

        public GridView[] arrGrdView;
        private void IN010F00_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Now;
            DateEditTo.EditValue = DateTime.Now;

            arrGrdView = new GridView[] { GridViewRetrL1, GridViewRetrR1_1, GridViewRetrR1_2, GridViewRetrL2, GridViewRetrR2_1, GridViewRetrR2_2 };
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
            try
            {
                Cursor = Cursors.WaitCursor;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                string sFindIdx = CboFindIdx.SelectedIndex.ToString();
                string sFindWord = TxtFindWord.EditValue?.ToString().Trim();

                if (string.IsNullOrEmpty(sYmdFrom))
                {
                    XtraMessageBox.Show("일자를 올바르게 설정하세요.");
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }
                else if (string.IsNullOrEmpty(sYmdTo))
                {
                    XtraMessageBox.Show("일자를 올바르게 설정하세요.");
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);
                dicParams.Add("FIND_IDX", sFindIdx);
                dicParams.Add("FIND_WORD", sFindWord);

                GridControl Grid = new GridControl();
                if (TabPage.SelectedTabPage == TabPagePurc)
                {
                    dicParams.Add("KERATYPE", "매입");
                    Grid = GridRetrL1;
                }
                else if (TabPage.SelectedTabPage == TabPageSale)
                {
                    dicParams.Add("KERATYPE", "매출");
                    Grid = GridRetrL2;
                }

                Grid.DataSource = GetDealerInfo(dicParams);
                if (Grid.MainView.RowCount > 0)
                {
                    Grid.MainView.Focus();
                }
                else
                {
                    if (string.IsNullOrEmpty(sFindWord))
                    {
                        DateEditFrom.SelectAll();
                        DateEditFrom.Focus();
                    }
                    else
                    {
                        TxtFindWord.SelectAll();
                        TxtFindWord.Focus();
                    }
                }

                if (TabPage.SelectedTabPage == TabPagePurc)
                {
                    SelDtlData1();
                }
                else if (TabPage.SelectedTabPage == TabPageSale)
                {
                    SelDtlData2();
                }

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private DataTable GetDealerInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB 
            //strSql.AppendLine(" SELECT A.J_ID1 AS DEALER_CD ");
            //strSql.AppendLine("      , C.DEALER_NM ");
            //strSql.AppendLine("      , CONCAT(SUBSTRING(C.IDT_NO, 1, 3), '-', SUBSTRING(C.IDT_NO, 4, 2), '-', SUBSTRING(C.IDT_NO, 6, 5)) AS IDT_NO ");
            //strSql.AppendLine("      , C.REP_NM ");
            //strSql.AppendLine("      , D.EMP_NM ");
            //strSql.AppendLine("      , CASE WHEN ( C.CHRG_HP_NO IS NULL OR C.CHRG_HP_NO = '' ) THEN C.CHRG_TEL_NO ELSE C.CHRG_HP_NO END AS TEL_NO ");
            //strSql.AppendLine("      , COUNT(CASE WHEN B.KERATYPE = '입고' THEN B.KERATYPE END) AS IN_CNT ");
            //strSql.AppendLine("      , COUNT(CASE WHEN B.KERATYPE = '출고' THEN B.KERATYPE END) AS OUT_CNT ");
            //strSql.AppendLine("      , SUM(IFNULL(A.IKONGKEP, 0)) AS TOT_IKONGKEP ");
            //strSql.AppendLine("      , SUM(IFNULL(A.KONGKEP, 0)) AS TOT_OKONGKEP ");
            //strSql.AppendLine("      , SUM(IFNULL(A.CKONGKEP, 0)) AS TOT_CKONGKEP ");
            //strSql.AppendLine("   FROM INLIST A  ");
            //strSql.AppendLine("   LEFT JOIN MESURING B  ");
            //strSql.AppendLine("     ON A.J_ID = CASE WHEN B.KERATYPE = '입고' THEN B.IPCHULGO_MAIPID ELSE B.IPCHULGO_MACHULID END  ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C  ");
            //strSql.AppendLine("     ON A.J_ID1 = C.DEALER_CD ");
            //strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS D  ");
            //strSql.AppendLine("     ON A.CHRG_ID = D.EMP_ID ");
            //strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            //strSql.AppendLine("    AND A.KERATYPE = '" + dicParams["KERATYPE"] + "' ");
            //strSql.AppendLine("    AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 )");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '0' AND (C.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) OR C.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '1' AND CAST(A.J_ID1 AS CHAR) LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '2' AND D.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '3' AND C.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '4' AND C.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' )) ");
            //strSql.AppendLine("  GROUP BY A.J_ID1 ");
            //strSql.AppendLine("  ORDER BY REPLACE(C.DEALER_NM, '(주)', '') ");
            #endregion

            strSql.AppendLine(" SELECT Z1.DEALER_CD                                                                                                   ");
            strSql.AppendLine("      , SUM(Z1.IN_CNT) AS IN_CNT                                                                                       ");
            strSql.AppendLine("      , SUM(Z1.OUT_CNT) AS OUT_CNT                                                                                     ");
            strSql.AppendLine("      , Z2.DEALER_NM                                                                                                   ");
            strSql.AppendLine("      , CONCAT(SUBSTRING(Z2.IDT_NO, 1, 3), '-', SUBSTRING(Z2.IDT_NO, 4, 2), '-', SUBSTRING(Z2.IDT_NO, 6, 5)) AS IDT_NO ");
            strSql.AppendLine("      , Z2.REP_NM                                                                                                      ");
            strSql.AppendLine("      , Z3.EMP_NM                                                                                                      ");
            strSql.AppendLine("      , CASE WHEN(Z2.CHRG_HP_NO IS NULL OR Z2.CHRG_HP_NO = '') THEN Z2.CHRG_TEL_NO ELSE Z2.CHRG_HP_NO END AS TEL_NO    ");
            strSql.AppendLine("      , SUM(ISNULL(Z1.TOT_IKONGKEP, 0)) AS TOT_IKONGKEP                                                                ");
            strSql.AppendLine("      , SUM(ISNULL(Z1.TOT_OKONGKEP, 0)) AS TOT_OKONGKEP                                                                ");
            strSql.AppendLine("      , SUM(ISNULL(Z1.TOT_CKONGKEP, 0)) AS TOT_CKONGKEP                                                                ");
            strSql.AppendLine("   FROM(SELECT A.J_ID1 AS DEALER_CD                                                                                    ");
            strSql.AppendLine("             , COUNT(CASE WHEN B.KERATYPE = '입고' THEN B.KERATYPE END) AS IN_CNT                                      ");
            strSql.AppendLine("             , COUNT(CASE WHEN B.KERATYPE = '출고' THEN B.KERATYPE END) AS OUT_CNT                                     ");
            strSql.AppendLine("             , SUM(ISNULL(A.IKONGKEP, 0)) AS TOT_IKONGKEP                                                              ");
            strSql.AppendLine("             , SUM(ISNULL(A.KONGKEP, 0)) AS TOT_OKONGKEP                                                               ");
            strSql.AppendLine("             , SUM(ISNULL(A.CKONGKEP, 0)) AS TOT_CKONGKEP                                                              ");
            strSql.AppendLine("          FROM INLIST A                                                                                                ");
            strSql.AppendLine("          LEFT JOIN MESURING B                                                                                         ");
            strSql.AppendLine("            ON A.J_ID = CASE WHEN B.KERATYPE = '입고' THEN B.IPCHULGO_MAIPID ELSE B.IPCHULGO_MACHULID END              ");
            strSql.AppendLine("          LEFT JOIN HR_EMP_BASIS D                                                                                     ");
            strSql.AppendLine("            ON A.CHRG_ID = D.EMP_ID                                                                                    ");
            strSql.AppendLine("         WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "'                        ");
            strSql.AppendLine("           AND A.KERATYPE = '" + dicParams["KERATYPE"] + "'                                                            "); ;
            strSql.AppendLine("         GROUP BY A.J_ID1                                                                                              ");

            //#00001
            if (dicParams["KERATYPE"].Equals("매출"))
            {
                strSql.AppendLine("         UNION ALL                                                                                                     ");
                strSql.AppendLine("        SELECT C.DEALER_CD                                                                                             ");
                strSql.AppendLine("             , '' AS IN_CNT                                                                                            ");
                strSql.AppendLine("             , '' AS OUT_CNT                                                                                           ");
                strSql.AppendLine("             , 0 AS TOT_IKONGKEP                                                                                       ");
                strSql.AppendLine("             , A.ADAMT AS TOT_OKONGKEP                                                                                 ");
                strSql.AppendLine("             , 0 AS TOT_CKONGKEP                                                                                       ");
                strSql.AppendLine("          FROM ACTRAN A                                                                                                ");
                strSql.AppendLine("          LEFT OUTER JOIN ACMSTF B                                                                                     ");
                strSql.AppendLine("            ON A.ACCOD = B.ACCOD                                                                                       ");
                strSql.AppendLine("          LEFT OUTER JOIN ACC_DEALER_CD C                                                                              ");
                strSql.AppendLine("            ON A.CVCOD = C.DEALER_CD                                                                                   ");
                strSql.AppendLine("          LEFT JOIN HR_EMP_BASIS D                                                                                     ");
                strSql.AppendLine("            ON C.CHRG_ID = D.EMP_ID                                                                                    ");
                strSql.AppendLine("         WHERE A.TDATE BETWEEN REPLACE('" + dicParams["DATE_F"] + "', '-', '') AND REPLACE('" + dicParams["DATE_T"] + "', '-', '')                       ");
                strSql.AppendLine("           AND A.ACCOD = '0404'--제품매출금                                                                            ");
                strSql.AppendLine("           AND A.RK LIKE '%인센티브%'                                                                             ");
                //#00002
                strSql.AppendLine("           AND A.APVYN = 'Y' ");
            }

            strSql.AppendLine("           ) Z1 ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD Z2                                                                                          ");
            strSql.AppendLine("     ON Z1.DEALER_CD = Z2.DEALER_CD                                                                                    ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS Z3                                                                                           ");
            strSql.AppendLine("     ON Z2.CHRG_ID = Z3.EMP_ID                                                                                         ");
            strSql.AppendLine("  WHERE(('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1)                                                                                            ");
            strSql.AppendLine("         OR                                                                                                            ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '0' AND(Z2.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') OR Z2.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')                                            ");
            strSql.AppendLine("         OR                                                                                                            ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '1' AND Z1.DEALER_CD LIKE '%" + dicParams["FIND_WORD"] + "%')                                                                        ");
            strSql.AppendLine("         OR                                                                                                            ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '2' AND Z3.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%')                                                                           ");
            strSql.AppendLine("         OR                                                                                                            ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '3' AND Z2.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%')                                                                           ");
            strSql.AppendLine("         OR                                                                                                            ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '4' AND Z2.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))                                                                          ");
            strSql.AppendLine("  GROUP BY Z1.DEALER_CD                                                                                                ");
            strSql.AppendLine("         , Z2.DEALER_NM                                                                                                ");
            strSql.AppendLine("         , CONCAT(SUBSTRING(Z2.IDT_NO, 1, 3), '-', SUBSTRING(Z2.IDT_NO, 4, 2), '-', SUBSTRING(Z2.IDT_NO, 6, 5))        ");
            strSql.AppendLine("         , Z2.REP_NM                                                                                                   ");
            strSql.AppendLine("         , Z3.EMP_NM                                                                                                   ");
            strSql.AppendLine("         , CASE WHEN(Z2.CHRG_HP_NO IS NULL OR Z2.CHRG_HP_NO = '') THEN Z2.CHRG_TEL_NO ELSE Z2.CHRG_HP_NO END           ");
            strSql.AppendLine("  ORDER BY REPLACE(Z2.DEALER_NM, '(주)', '')                                                                           ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();

            try
            {
                GridView[] arrViews = new GridView[3];
                if (TabPage.SelectedTabPage == TabPagePurc)
                {
                    arrViews[0] = GridViewRetrL1;
                    arrViews[1] = GridViewRetrR1_1;
                    arrViews[2] = GridViewRetrR1_2;
                }
                else
                {
                    arrViews[0] = GridViewRetrL2;
                    arrViews[1] = GridViewRetrR2_1;
                    arrViews[2] = GridViewRetrR2_2;
                }

                /*
                 * 각 그리드 뷰를 개별 엑셀파일로 만들어 하나의 엑셀파일로 취합
                 * 여기서 Sheet이름은 GridView의 Tag속성을 따라감
                 */
                string sDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string sTempExcelPath = string.Format("{0}\\{1}", System.Windows.Forms.Application.StartupPath, "Temp_Excel");
                DirectoryInfo directory = new DirectoryInfo(sTempExcelPath);
                if (directory.Exists)
                {
                    directory.Delete(true);
                    directory.Create();
                }
                else
                {
                    directory.Create();
                }
                
                foreach (GridView view in arrViews)
                {
                    view.ExportToXlsx(string.Format("{0}\\{1}.xlsx", sTempExcelPath, view.Tag));
                }
                
                Excel.Workbook MainWorkbook = xlApp.Workbooks.Open(string.Format("{0}\\{1}.xlsx", sTempExcelPath, arrViews[0].Tag)); //가장 첫번째 엑셀파일을 기준하여 해당 파일에 Sheet복사
                MainWorkbook.Worksheets["Sheet"].Name = arrViews[0].Tag; //해당 시트 이름변경
                MainWorkbook.Save();

                //int iCnt = 1;
                //foreach (GridView view in arrViews)
                //{
                //    if (view == arrViews[0])
                //        continue;

                //    Excel.Workbook workbook = xlApp.Workbooks.Open(string.Format("{0}\\{1}.xlsx", sTempExcelPath, view.Tag));
                //    MainWorkbook.Worksheets.Copy(workbook.Worksheets[1]);
                //    //MainWorkbook.Worksheets["Sheet"].Name = view.Tag; //해당 시트 이름변경
                //    //MainWorkbook.Save();
                //    //workbook.Close();
                //}

                Excel.Workbook workbook = xlApp.Workbooks.Open(string.Format("{0}\\{1}.xlsx", sTempExcelPath, arrViews[1].Tag));
                Excel.Worksheet workSheet = workbook.Worksheets["Sheet"];

                //workbook.SaveAs(string.Format("{0}\\{1}.xlsx", sDesktopPath, "ddd"));

                Excel.Worksheet workSheetMain = MainWorkbook.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]) as Excel.Worksheet;
                workbook.Worksheets["Sheet1"].Copy(workbook.Worksheets["Sheet"]);

                //workSheetMain.Copy(Type.Missing, Type.Missing);
                //workSheetMain = workSheet;

                MainWorkbook.SaveAs(string.Format("{0}\\{1}.xlsx", sDesktopPath, this.Text));
                MainWorkbook.Close();
                workbook.Close(false);
                xlApp.Quit();

                //File.Copy(string.Format("{0}\\{1}.xlsx", sTempExcelPath, arrViews[0].Tag), string.Format("{0}\\{1}.xlsx", sDesktopPath, this.Text), true);
                Process.Start(string.Format("{0}\\{1}.xlsx", sDesktopPath, this.Text));

                
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            //finally
            //{
            //    xlApp.Quit();
            //}
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void TabPage_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewRetrL1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SelDtlData1();
        }

        private void SelDtlData1()
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("J_ID", GridViewRetrL1.GetFocusedRowCellValue(GridColL1DealerCd)?.ToString());

            GridRetrR1_1.DataSource = GetPurchaseInfo(dicParams);
            GridRetrR1_2.DataSource = GetDirectSendInfo(dicParams, "1");

            SetInitValue();
        }

        private DataTable GetPurchaseInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT A.J_DATE AS CLOSE_DT ");
            //strSql.AppendLine("      , G.J_DATE AS MESURE_DT  ");
            //strSql.AppendLine("      , D.DEALER_NM  "); //2020-11-24 현업요청에 따라 계근일자 빼고 매입처로 변경
            //strSql.AppendLine("      , B.J_BNUM  ");
            //strSql.AppendLine("      , G.WEIGHT AS LANDEDWEIGHT  ");
            //strSql.AppendLine("      , A.HALIN AS LOSS  ");
            //strSql.AppendLine("      , A.IWEIGHT AS ACCEPTWEIGHT  ");
            //strSql.AppendLine("      , IFNULL(A.IWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT  ");
            //strSql.AppendLine("      , F.GUBUN1  ");
            //strSql.AppendLine("      , A.MIDANGA AS STDDUNITPRICE  ");
            //strSql.AppendLine("      , A.DANGA AS PAYEDUNITPRICE  ");
            //strSql.AppendLine("      , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT - IFNULL(A.MIDANGA, 0) AS DIFFUNITPRICE  ");
            //strSql.AppendLine("      , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE  ");
            //strSql.AppendLine("      , IFNULL(A.CKONGKEP, 0) AS CARRYCOST  ");
            //strSql.AppendLine("      , A.IKONGKEP AS TOTALPRICE  ");
            //strSql.AppendLine("      , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE  ");
            //strSql.AppendLine("      , G.J_STATE AS LOSSCAUSE  ");
            //strSql.AppendLine("      , G.JUNPYOID AS IMAGE  ");
            //strSql.AppendLine("      , G.GUMSUBIGO AS INSPECTNOTE  ");
            //strSql.AppendLine("     , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE  ");
            //strSql.AppendLine("     , H.DEALER_NM AS ETC_DEALER_NM1  ");
            //strSql.AppendLine("     , G.ETC_COST1  ");
            //strSql.AppendLine("     , G.ETC_REMARK1  ");
            //strSql.AppendLine("     , I.DEALER_NM AS ETC_DEALER_NM2  ");
            //strSql.AppendLine("     , G.ETC_COST2  ");
            //strSql.AppendLine("     , G.ETC_REMARK2  ");
            //strSql.AppendLine("  FROM INLIST A  ");
            //strSql.AppendLine("  LEFT OUTER JOIN IPCHULGO B  ");
            //strSql.AppendLine("     ON B.J_ID = A.J_ID  ");
            //strSql.AppendLine("  LEFT OUTER JOIN JAJAE C  ");
            //strSql.AppendLine("    ON C.J_SERIAL = A.J_SERIAL  ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D  ");
            //strSql.AppendLine("    ON D.DEALER_CD = A.J_ID1  ");
            //strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS E  ");
            //strSql.AppendLine("    ON E.EMP_ID = A.CHRG_ID  ");
            //strSql.AppendLine("  LEFT OUTER JOIN JAJAE F  ");
            //strSql.AppendLine("    ON F.GUBUN1 = C.GUBUN1  ");
            //strSql.AppendLine("  LEFT OUTER JOIN MESURING G  ");
            //strSql.AppendLine("    ON G.IPCHULGO_MAIPID = A.J_ID  ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD H ");
            //strSql.AppendLine("    ON H.DEALER_CD = G.ETC_DEALER_CD1  ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD I ");
            //strSql.AppendLine("    ON I.DEALER_CD = G.ETC_DEALER_CD2  ");
            //strSql.AppendLine(" WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            //strSql.AppendLine("   AND A.KERATYPE = '매입' ");
            //strSql.AppendLine("   AND B.KERAGUBUN <> '직송'  ");
            //strSql.AppendLine("   AND G.J_CHECK = '1' ");
            //strSql.AppendLine("   AND A.DANGA != 0    ");
            //strSql.AppendLine("   AND A.J_ID1 = " + dicParams["J_ID"] + " ");
            //strSql.AppendLine(" ORDER BY A.J_DATE, D.DEALER_NM ");
            #endregion

            strSql.AppendLine(" SELECT A.J_DATE AS CLOSE_DT ");
            strSql.AppendLine("      , G.J_DATE AS MESURE_DT  ");
            strSql.AppendLine("      , D.DEALER_NM  "); //2020-11-24 현업요청에 따라 계근일자 빼고 매입처로 변경
            strSql.AppendLine("      , B.J_BNUM  ");
            strSql.AppendLine("      , G.WEIGHT AS LANDEDWEIGHT  ");
            strSql.AppendLine("      , A.HALIN AS LOSS  ");
            strSql.AppendLine("      , A.IWEIGHT AS ACCEPTWEIGHT  ");
            strSql.AppendLine("      , ISNULL(A.IWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT  ");
            strSql.AppendLine("      , F.GUBUN1  ");
            strSql.AppendLine("      , A.MIDANGA AS STDDUNITPRICE  ");
            strSql.AppendLine("      , A.DANGA AS PAYEDUNITPRICE  ");
            strSql.AppendLine("      , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0)) / A.IWEIGHT - ISNULL(A.MIDANGA, 0) AS DIFFUNITPRICE  ");
            strSql.AppendLine("      , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE  ");
            strSql.AppendLine("      , ISNULL(A.CKONGKEP, 0) AS CARRYCOST  ");
            strSql.AppendLine("      , A.IKONGKEP AS TOTALPRICE  ");
            strSql.AppendLine("      , (ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE  ");
            strSql.AppendLine("      , G.J_STATE AS LOSSCAUSE  ");
            strSql.AppendLine("      , G.JUNPYOID AS IMAGE  ");
            strSql.AppendLine("      , G.GUMSUBIGO AS INSPECTNOTE  ");
            strSql.AppendLine("     , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG), 0, 1) AS CARRY_UNIT_PRICE  ");
            strSql.AppendLine("     , H.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("     , G.ETC_COST1  ");
            strSql.AppendLine("     , G.ETC_REMARK1  ");
            strSql.AppendLine("     , I.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("     , G.ETC_COST2  ");
            strSql.AppendLine("     , G.ETC_REMARK2  ");
            strSql.AppendLine("  FROM INLIST A  ");
            strSql.AppendLine("  LEFT OUTER JOIN IPCHULGO B  ");
            strSql.AppendLine("     ON B.J_ID = A.J_ID  ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE C  ");
            strSql.AppendLine("    ON C.J_SERIAL = A.J_SERIAL  ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D  ");
            strSql.AppendLine("    ON D.DEALER_CD = A.J_ID1  ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS E  ");
            strSql.AppendLine("    ON E.EMP_ID = A.CHRG_ID  ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE F  ");
            strSql.AppendLine("    ON F.GUBUN1 = C.GUBUN1  ");
            strSql.AppendLine("  LEFT OUTER JOIN MESURING G  ");
            strSql.AppendLine("    ON G.IPCHULGO_MAIPID = A.J_ID  ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("    ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("    ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine(" WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("   AND A.KERATYPE = '매입' ");
            strSql.AppendLine("   AND B.KERAGUBUN <> '직송'  ");
            strSql.AppendLine("   AND G.J_CHECK = '1' ");
            strSql.AppendLine("   AND A.DANGA != 0    ");
            strSql.AppendLine("   AND A.J_ID1 = '" + dicParams["J_ID"] + "' ");
            strSql.AppendLine(" ORDER BY A.J_DATE, D.DEALER_NM ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewRetrL2_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SelDtlData2();
        }

        private void SelDtlData2()
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("J_ID", GridViewRetrL2.GetFocusedRowCellValue(GridColL2DealerCd)?.ToString());

            GridRetrR2_1.DataSource = SalesPerBuyer(dicParams);
            GridRetrR2_2.DataSource = GetDirectSendInfo(dicParams, "2");

            SetInitValue2();
        }

        private DataTable SalesPerBuyer(Dictionary<string, string> dicParams)
        {
            if (string.IsNullOrEmpty(dicParams["J_ID"]))
            {
                return null;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , A.JUNPYOID ");
            //strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("               , D.DEALER_NM ");
            //strSql.AppendLine("               , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("               , B.J_BNUM ");
            //strSql.AppendLine("               , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("               , A.HALIN AS LOSS ");
            //strSql.AppendLine("               , A.OWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("               , IFNULL(A.OWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("               , F.GUBUN1 ");
            //strSql.AppendLine("               , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("               , A.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine("               , ((IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("               , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("               , A.KONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("               , (IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("               , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("               , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("               , G.ETC_COST2 ");
            //strSql.AppendLine("               , G.ETC_REMARK2 ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("              ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("              ON E.EMP_ID = A.CHRG_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("              ON G.IPCHULGO_MACHULID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("              ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("              ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("           WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            //strSql.AppendLine("             AND A.J_ID1 = '" + dicParams["J_ID"] + "' ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");
            #endregion

            //#00001
            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT A.J_DATE ");
            strSql.AppendLine("               , A.JUNPYOID ");
            strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            strSql.AppendLine("               , D.DEALER_NM ");
            strSql.AppendLine("               , E.EMP_NM AS INSPECTOR ");
            strSql.AppendLine("               , B.J_BNUM ");
            strSql.AppendLine("               , G.WEIGHT AS LANDEDWEIGHT ");
            strSql.AppendLine("               , A.HALIN AS LOSS ");
            strSql.AppendLine("               , A.OWEIGHT AS ACCEPTWEIGHT ");
            strSql.AppendLine("               , ISNULL(A.OWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            strSql.AppendLine("               , F.GUBUN1 ");
            strSql.AppendLine("               , A.MIDANGA AS STDDUNITPRICE ");
            strSql.AppendLine("               , A.DANGA AS SALEUNITPRICE ");
            strSql.AppendLine("               , ((ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0)) / NULLIF(A.OWEIGHT,0) AS ARRVUNITPRICE ");
            strSql.AppendLine("               , ISNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            strSql.AppendLine("               , A.KONGKEP AS TOTALPRICE ");
            strSql.AppendLine("               , (ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("               , G.JUNPYOID AS IMAGE ");
            strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE ");
            strSql.AppendLine("               , ROUND((ISNULL(A.CKONGKEP, 0) / NULLIF(A.DANJUNG,0)), 0, 1) AS CARRY_UNIT_PRICE ");
            strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1 ");
            strSql.AppendLine("               , G.ETC_COST1 ");
            strSql.AppendLine("               , G.ETC_REMARK1 ");
            strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2 ");
            strSql.AppendLine("               , G.ETC_COST2 ");
            strSql.AppendLine("               , G.ETC_REMARK2 ");
            strSql.AppendLine("            FROM INLIST A ");
            strSql.AppendLine("            LEFT OUTER JOIN IPCHULGO B ");
            strSql.AppendLine("              ON B.J_ID = A.J_ID ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            strSql.AppendLine("              ON E.EMP_ID = A.CHRG_ID ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            strSql.AppendLine("              ON G.IPCHULGO_MACHULID = A.J_ID ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("           WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("             AND A.J_ID1 = '" + dicParams["J_ID"] + "' ");
            strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            strSql.AppendLine("             AND G.J_CHECK = '1' ");
            strSql.AppendLine("           UNION ALL                                                      ");
            strSql.AppendLine("          SELECT CONVERT(VARCHAR(10),CONVERT(DATE, A.TDATE),23) AS J_DATE ");
            strSql.AppendLine("               , NULL AS JUNPYOID                                         ");
            strSql.AppendLine("               , '' AS MESURE_DT                                          ");
            strSql.AppendLine("               , C.DEALER_NM                                              ");
	        strSql.AppendLine("               , '' AS INSPECTOR                                          ");
            strSql.AppendLine("               , '' AS J_BNUM                                             ");
            strSql.AppendLine("               , 0 AS LANDEDWEIGHT                                        ");
            strSql.AppendLine("               , 0 AS LOSS                                                ");
            strSql.AppendLine("               , 0 AS ACCEPTWEIGHT                                        ");
            strSql.AppendLine("               , 0 AS DIFFWEIGHT                                          ");
            strSql.AppendLine("               , '인센티브' AS GUBUN1                                     ");
            strSql.AppendLine("               , 0 AS STDDUNITPRICE                                       ");
            strSql.AppendLine("               , 0 AS SALEUNITPRICE                                       ");
            strSql.AppendLine("               , 0 AS ARRVUNITPRICE                                       ");
            strSql.AppendLine("               , 0 AS CARRYCOST                                           ");
            strSql.AppendLine("               , A.ADAMT AS TOTALPRICE                                    ");
	        strSql.AppendLine("               , A.ADAMT AS ARRVTOTALPRICE                                ");
	        strSql.AppendLine("               , '' AS LOSSCAUSE                                          ");
            strSql.AppendLine("               , '' AS IMAGE                                              ");
            strSql.AppendLine("               , '' AS INSPECTNOTE                                        ");
            strSql.AppendLine("               , 0 AS CARRY_UNIT_PRICE                                    ");
            strSql.AppendLine("               , '' AS ETC_DEALER_NM1                                     ");
            strSql.AppendLine("               , 0 AS ETC_COST1                                           ");
            strSql.AppendLine("               , '' AS ETC_REMARK1                                        ");
            strSql.AppendLine("               , '' AS ETC_DEALER_NM2                                     ");
            strSql.AppendLine("               , 0 AS ETC_COST2                                           ");
            strSql.AppendLine("               , '' AS ETC_REMARK2                                        ");
            strSql.AppendLine("            FROM ACTRAN A                                                 ");
            strSql.AppendLine("            LEFT OUTER JOIN ACMSTF B                                      ");
            strSql.AppendLine("              ON A.ACCOD = B.ACCOD                                        ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C                               ");
            strSql.AppendLine("              ON A.CVCOD = C.DEALER_CD                                    ");
            strSql.AppendLine("           WHERE A.TDATE BETWEEN REPLACE('" + dicParams["DATE_F"] + "', '-', '') AND REPLACE('" + dicParams["DATE_T"] + "','-','')"); 
            strSql.AppendLine("             AND A.ACCOD = '0404'--제품매출금                                                   ");
            strSql.AppendLine("             AND A.RK LIKE '%인센티브%'                                                         ");
            strSql.AppendLine("             AND C.DEALER_CD = '" + dicParams["J_ID"] + "' ");
            //#00002
            strSql.AppendLine("             AND A.APVYN = 'Y' ");
            strSql.AppendLine("      ) Y1 ");
            strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");   

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetDirectSendInfo(Dictionary<string, string> dicParams, string sGb)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT B.J_DATE ");
            //strSql.AppendLine("      , XX.EMP_NM AS NAME ");
            //strSql.AppendLine("      , H.GUBUN1 ");
            //strSql.AppendLine("      , A.J_BNUM ");
            //strSql.AppendLine("      , D.DEALER_NM AS SALE_DEALER_NM ");
            //strSql.AppendLine("      , B.DANJUNG AS OWEIGHT ");
            //strSql.AppendLine("      , B.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine("      , B.KONGKEP AS SALEPRICE ");
            //strSql.AppendLine("      , E.DEALER_NM AS PURC_DEALER_NM ");
            //strSql.AppendLine("      , C.DANGA AS PURCHUNITPRICE ");
            //strSql.AppendLine("      , C.IKONGKEP AS PURC_AMT ");
            //strSql.AppendLine("      , TRUNCATE(((IFNULL(B.DANJUNG, 0) * IFNULL(C.DANGA, 0)) + IFNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1) AS ARRIVEUNITPRICE  ");
            //strSql.AppendLine("      , IFNULL(B.CKONGKEP, 0) AS CARRYCOST  ");
            //strSql.AppendLine("      , TRUNCATE((IFNULL(B.CKONGKEP, 0) / B.DANJUNG), 1) AS CARRY_UNIT_PRICE  ");
            //strSql.AppendLine("      , (IFNULL(C.DANJUNG, 0) * IFNULL(C.DANGA, 0)) + IFNULL(B.CKONGKEP, 0) AS ARRVTOTALPRICE  ");
            ////strSql.AppendLine("      , IFNULL(B.DANGA, 0) - IFNULL(C.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , IFNULL(B.DANGA, 0) - TRUNCATE(((IFNULL(B.DANJUNG, 0) * IFNULL(C.DANGA, 0)) + IFNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1) AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , C.HALIN ");
            //strSql.AppendLine("      , F.DEALER_NM AS ETC_DEALER_NM1  ");
            //strSql.AppendLine("      , A.ETC_COST1  ");
            //strSql.AppendLine("      , A.ETC_REMARK1  ");
            //strSql.AppendLine("      , G.DEALER_NM AS ETC_DEALER_NM2  ");
            //strSql.AppendLine("      , A.ETC_COST2  ");
            //strSql.AppendLine("      , A.ETC_REMARK2  ");
            //strSql.AppendLine("   FROM MESURING A ");
            //strSql.AppendLine("  LEFT OUTER JOIN INLIST B  ");
            //strSql.AppendLine("    ON A.IPCHULGO_MACHULID = B.J_ID ");
            //strSql.AppendLine("  LEFT OUTER JOIN INLIST C  ");
            //strSql.AppendLine("    ON A.IPCHULGO_MAIPID = C.J_ID ");
            //strSql.AppendLine("  LEFT OUTER JOIN ( SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID ");
            //strSql.AppendLine("                      FROM ACC_DEALER_CD X1 ");
            //strSql.AppendLine("                      LEFT OUTER JOIN HR_EMP_BASIS X2 ");
            //strSql.AppendLine("                        ON X1.CHRG_ID = X2.EMP_ID ) D  ");
            //strSql.AppendLine("    ON A.J_ASSIGNID = D.DEALER_CD #매출처 ");
            //strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS XX ");
            //strSql.AppendLine("    ON C.CHRG_ID = XX.EMP_ID ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD E #매입처  ");
            //strSql.AppendLine("    ON C.J_ID1 = E.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD F ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD1 = F.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD G ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD2 = G.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN JAJAE H ");
            //strSql.AppendLine("    ON A.J_SERIAL = H.J_SERIAL  ");
            //strSql.AppendLine(" WHERE A.KERATYPE = '직송' ");
            //strSql.AppendLine("   AND B.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("   AND C.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("   AND B.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            //if (sGb.Equals("1"))
            //{
            //    strSql.AppendLine("   AND C.J_ID1 = " + dicParams["J_ID"] + " ");
            //}
            //else
            //{
            //    strSql.AppendLine("   AND B.J_ID1 = " + dicParams["J_ID"] + " ");
            //}
            //strSql.AppendLine("  ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");
            #endregion

            strSql.AppendLine(" SELECT B.J_DATE ");
            strSql.AppendLine("      , XX.EMP_NM AS NAME ");
            strSql.AppendLine("      , H.GUBUN1 ");
            strSql.AppendLine("      , A.J_BNUM ");
            strSql.AppendLine("      , D.DEALER_NM AS SALE_DEALER_NM ");
            strSql.AppendLine("      , B.DANJUNG AS OWEIGHT ");
            strSql.AppendLine("      , B.DANGA AS SALEUNITPRICE ");
            strSql.AppendLine("      , B.KONGKEP AS SALEPRICE ");
            strSql.AppendLine("      , E.DEALER_NM AS PURC_DEALER_NM ");
            strSql.AppendLine("      , C.DANGA AS PURCHUNITPRICE ");
            strSql.AppendLine("      , C.IKONGKEP AS PURC_AMT ");
            strSql.AppendLine("      , ROUND(((ISNULL(B.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0)) / B.DANJUNG, 0, 1) AS ARRIVEUNITPRICE  ");
            strSql.AppendLine("      , ISNULL(B.CKONGKEP, 0) AS CARRYCOST  ");
            strSql.AppendLine("      , ROUND((ISNULL(B.CKONGKEP, 0) / B.DANJUNG),0, 1) AS CARRY_UNIT_PRICE  ");
            strSql.AppendLine("      , (ISNULL(C.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0) AS ARRVTOTALPRICE  ");
            //strSql.AppendLine("      , ISNULL(B.DANGA, 0) - ISNULL(C.DANGA, 0) AS DIFFUNITPRICE ");
            strSql.AppendLine("      , ISNULL(B.DANGA, 0) - ROUND(((ISNULL(B.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0)) / B.DANJUNG, 0, 1) AS DIFFUNITPRICE ");
            strSql.AppendLine("      , C.HALIN ");
            strSql.AppendLine("      , I.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("      , A.ETC_COST1  ");
            strSql.AppendLine("      , A.ETC_REMARK1  ");
            strSql.AppendLine("      , J.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("      , A.ETC_COST2  ");
            strSql.AppendLine("      , A.ETC_REMARK2  ");
            strSql.AppendLine("   FROM MESURING A ");
            strSql.AppendLine("  LEFT OUTER JOIN INLIST B  ");
            strSql.AppendLine("    ON A.IPCHULGO_MACHULID = B.J_ID ");
            strSql.AppendLine("  LEFT OUTER JOIN INLIST C  ");
            strSql.AppendLine("    ON A.IPCHULGO_MAIPID = C.J_ID ");
            strSql.AppendLine("  LEFT OUTER JOIN ( SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID ");
            strSql.AppendLine("                      FROM ACC_DEALER_CD X1 ");
            strSql.AppendLine("                      LEFT OUTER JOIN HR_EMP_BASIS X2 ");
            strSql.AppendLine("                        ON X1.CHRG_ID = X2.EMP_ID ) D  ");
            strSql.AppendLine("    ON A.J_ASSIGNID = D.DEALER_CD --매출처 ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS XX ");
            strSql.AppendLine("    ON C.CHRG_ID = XX.EMP_ID ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD E --매입처  ");
            strSql.AppendLine("    ON C.J_ID1 = E.DEALER_CD ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE H ");
            strSql.AppendLine("    ON A.J_SERIAL = H.J_SERIAL  ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD1 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD J          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD2 = CONVERT(VARCHAR,J.DEALER_CD)");
            strSql.AppendLine(" WHERE A.KERATYPE = '직송' ");
            strSql.AppendLine("   AND B.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("   AND C.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("   AND B.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            if (sGb.Equals("1"))
            {
                strSql.AppendLine("   AND C.J_ID1 = '" + dicParams["J_ID"] + "' ");
            }
            else
            {
                strSql.AppendLine("   AND B.J_ID1 = '" + dicParams["J_ID"] + "' ");
            }
            strSql.AppendLine("  ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void SetInitValue()
        {
            double dLandedWeight = 0; //입고리스트 대지중량
            double dAcptWeight = 0; //직송리스트 인수량

            double dPurcAmt1 = 0; //입고리스트 매입액
            double dPurcAmt2 = 0; //직송리스트 매입액

            double dCarryCost1 = 0; //입고리스트 운반비
            double dCarryCost2 = 0; //직송리스트 운반비

            double dArrvAmt1 = 0; //입고리스트 도착도금액
            double dArrvAmt2 = 0; //직송리스트 도착도금액

            double.TryParse(GridCol1_2AcceptWeight.SummaryItem.SummaryValue?.ToString(), out dLandedWeight);
            double.TryParse(GridCol1_3OWeight.SummaryItem.SummaryValue?.ToString(), out dAcptWeight);

            double.TryParse(GridCol1_2TotalPrice.SummaryItem.SummaryValue?.ToString(), out dPurcAmt1);
            double.TryParse(GridCol1_3PurcAmt.SummaryItem.SummaryValue?.ToString(), out dPurcAmt2);

            double.TryParse(GridCol1_2CarryCost.SummaryItem.SummaryValue?.ToString(), out dCarryCost1);
            double.TryParse(GridCol1_3CarryCost.SummaryItem.SummaryValue?.ToString(), out dCarryCost2);

            double.TryParse(GridCol1_2ArrvTotalPrice.SummaryItem.SummaryValue?.ToString(), out dArrvAmt1);
            double.TryParse(GridCol1_3ArrvTotalPrice.SummaryItem.SummaryValue?.ToString(), out dArrvAmt2);
            
            string sWeight = string.Format("인수량 : {0}", (dLandedWeight + dAcptWeight).ToString("n0"));
            string sPurcAmt = string.Format("매입액 : {0}", (dPurcAmt1 + dPurcAmt2).ToString("n0"));
            string sCarryCost = string.Format("운반비 : {0}", (dCarryCost1 + dCarryCost2).ToString("n0"));
            string sArrvAmt = string.Format("도착도금액 : {0}", (dArrvAmt1 + dArrvAmt2).ToString("n0"));

            LabelTotal1.Text = string.Format("{0}, {1}, {2}, {3}", sWeight, sPurcAmt, sCarryCost, sArrvAmt);
        }

        private void SetInitValue2()
        {
            double dLandedWeight = 0; //매출리스트 대지중량
            double dAcptWeight = 0; //직송리스트 인수량

            double dSaleAmt1 = 0; //매출리스트 매출액
            double dSaleAmt2 = 0; //직송리스트 매출액

            double dCarryCost1 = 0; //매출리스트 운반비
            double dCarryCost2 = 0; //직송리스트 운반비

            double dArrvAmt1 = 0; //매출리스트 도착도금액
            double dArrvAmt2 = 0; //직송리스트 도착도금액
            
            double.TryParse(GridColR2_1AcceptWeight.SummaryItem.SummaryValue?.ToString(), out dLandedWeight);
            double.TryParse(GridColR2_2OWeight.SummaryItem.SummaryValue?.ToString(), out dAcptWeight);

            double.TryParse(GridColR2_1TotalPrice.SummaryItem.SummaryValue?.ToString(), out dSaleAmt1);
            double.TryParse(GridColR2_2SalePrice.SummaryItem.SummaryValue?.ToString(), out dSaleAmt2);

            double.TryParse(GridColR2_1CarryCost.SummaryItem.SummaryValue?.ToString(), out dCarryCost1);
            double.TryParse(GridColR2_2CarryCost.SummaryItem.SummaryValue?.ToString(), out dCarryCost2);

            double.TryParse(GridColR2_1ArrvTotalPrice.SummaryItem.SummaryValue?.ToString(), out dArrvAmt1);
            double.TryParse(GridColR2_2ArrvtotalPrice.SummaryItem.SummaryValue?.ToString(), out dArrvAmt2);

            string sWeight = string.Format("인수량 : {0}", (dLandedWeight + dAcptWeight).ToString("n0"));
            string sPurcAmt = string.Format("매출액 : {0}", (dSaleAmt1 + dSaleAmt2).ToString("n0"));
            string sCarryCost = string.Format("운반비 : {0}", (dCarryCost1 + dCarryCost2).ToString("n0"));
            string sArrvAmt = string.Format("도착도금액 : {0}", (dArrvAmt1 + dArrvAmt2).ToString("n0"));

            LabelTotal2.Text = string.Format("{0}, {1}, {2}, {3}", sWeight, sPurcAmt, sCarryCost, sArrvAmt);
        }

        private void GridViewRetrL1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetrL1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetrL2_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetrL2_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void IN010F00_TextChanged(object sender, EventArgs e)
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

        private void IN010F00_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}