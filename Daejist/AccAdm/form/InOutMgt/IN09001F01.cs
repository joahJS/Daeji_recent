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
using DevExpress.XtraBars;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using System.IO;
/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-01
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            매입 (기존) - 차액 = 지급단가 - 기준단가
 *                 (수정) - 차액 = 도착도단가 - 기준단가
 *            매출 (기존) - 차액 = 매출단가 - 기준단가
 *                 (수정) - 차액 = 도착도단가 - 기준단가
 *            직송 (기존) - 차액 = 매출단가 - 매입단가
 *                 (수정) - 차액 = 매출단가 - 도착도단가
 *            
 *            해당사항에 대하여 조회쿼리 수정
 *            
 * 수정일자 : 2021-02-08
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            매입처별 마감조회에서 매출입 전체 조회할 수 있도록 탭 및 쿼리, 로직 수정
 * 
 */
namespace AccAdm
{
    public partial class IN09001F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN09001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void IN09001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            InitControls();
            UpdateDropDownButton(BtnDealerInfo);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            //arrGrdView = new GridView[] { GridViewRetr1, BGridViewRetr2, BGridViewRetr3 };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("FIND_IDX", CboFindSbj.SelectedIndex.ToString());
            dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());
            
            GridRetr1.DataSource = null;
            GridRetr1.DataSource = GetDealerInfo(dicParams);
            if(GridViewRetr1.RowCount == 0)
            {
                TxtFindWord.SelectAll();
                TxtFindWord.Focus();
            }
            else
            {
                GridViewRetr1.Focus();
            }
        }

        private DataTable GetDealerInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.J_ID1 AS DEALER_CD ");
            strSql.AppendLine("      , C.DEALER_NM ");
            strSql.AppendLine("      , CONCAT(SUBSTRING(C.IDT_NO, 1, 3), '-', SUBSTRING(C.IDT_NO, 4, 2), '-', SUBSTRING(C.IDT_NO, 6, 5)) AS IDT_NO ");
            strSql.AppendLine("      , C.REP_NM ");
            strSql.AppendLine("      , D.EMP_NM ");
            strSql.AppendLine("      , CASE WHEN ( C.CHRG_HP_NO IS NULL OR C.CHRG_HP_NO = '' ) THEN C.CHRG_TEL_NO ELSE C.CHRG_HP_NO END AS TEL_NO ");
            strSql.AppendLine("      , COUNT(CASE WHEN B.KERATYPE = '입고' THEN B.KERATYPE END) AS IN_CNT ");
            strSql.AppendLine("      , COUNT(CASE WHEN B.KERATYPE = '출고' THEN B.KERATYPE END) AS OUT_CNT ");
            strSql.AppendLine("      , COUNT(CASE WHEN B.KERATYPE = '직송' THEN B.KERATYPE END) AS DIRECT_CNT ");
            strSql.AppendLine("      , SUM(IFNULL(A.IKONGKEP, 0)) AS TOT_IKONGKEP ");
            strSql.AppendLine("      , SUM(IFNULL(A.CKONGKEP, 0)) AS TOT_CKONGKEP ");
            strSql.AppendLine("   FROM INLIST A  ");
            strSql.AppendLine("   LEFT JOIN MESURING B  ");
            strSql.AppendLine("     ON A.J_ID = CASE WHEN B.KERATYPE = '입고' THEN B.IPCHULGO_MAIPID ELSE B.IPCHULGO_MACHULID END  ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C  ");
            strSql.AppendLine("     ON A.J_ID1 = C.DEALER_CD ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS D  ");
            strSql.AppendLine("     ON C.CHRG_ID = D.EMP_ID ");
            strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("    AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 )");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '0' AND (C.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) OR C.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '1' AND CAST(A.J_ID1 AS CHAR) LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '2' AND D.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '3' AND C.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '4' AND C.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' )) ");
            //strSql.AppendLine("  WHERE A.KERATYPE = '매입' ");
            //strSql.AppendLine("    AND ( B.KERATYPE = '입고' OR B.KERATYPE = '직송' ) ");
            //strSql.AppendLine("    AND A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            //strSql.AppendLine("    AND ( ('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 )");
            //strSql.AppendLine("          OR ");
            //strSql.AppendLine("          ('" + dicParams["FIND_IDX"] + "' = '0' AND C.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("          OR ");
            //strSql.AppendLine("          ('" + dicParams["FIND_IDX"] + "' = '1' AND CAST(A.J_ID1 AS CHAR) LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("          OR ");
            //strSql.AppendLine("          ('" + dicParams["FIND_IDX"] + "' = '2' AND D.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("          OR ");
            //strSql.AppendLine("          ('" + dicParams["FIND_IDX"] + "' = '3' AND C.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            //strSql.AppendLine("          OR ");
            //strSql.AppendLine("          ('" + dicParams["FIND_IDX"] + "' = '4' AND C.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ) ");
            strSql.AppendLine("  GROUP BY A.J_ID1 ");
            strSql.AppendLine("  ORDER BY REPLACE(C.DEALER_NM, '(주)', '') ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewRetr1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(e.FocusedRowHandle < 0)
            {
                GridRetr2.DataSource = null;
                return;
            }
                
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("J_ID", GridViewRetr1.GetFocusedRowCellValue(GridColDealerCd)?.ToString());

            GridControl gdPurcSale = new GridControl();
            DataTable dt = new DataTable();
            if(TabControlPurcSale.SelectedTabPage == TabPagePurc)
            {
                dt = GetPurchaseInfo(dicParams);
                gdPurcSale = GridRetr2;
            }
            else
            {
                dt = SalesPerBuyer(dicParams);
                gdPurcSale = GridRetr4;
            }

            gdPurcSale.DataSource = null;
            gdPurcSale.DataSource = dt;
            
            //GridRetr3.DataSource = null;
            //GridRetr3.DataSource = GetDirectSendInfo(dicParams);

            SetInitValue();
        }

        private void TabControlPurcSale_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Substring(0, 10));
            dicParams.Add("J_ID", GridViewRetr1.GetFocusedRowCellValue(GridColDealerCd)?.ToString());

            GridControl gdPurcSale = new GridControl();
            DataTable dt = new DataTable();
            if (TabControlPurcSale.SelectedTabPage == TabPagePurc)
            {
                dt = GetPurchaseInfo(dicParams);
                gdPurcSale = GridRetr2;
            }
            else
            {
                dt = SalesPerBuyer(dicParams);
                gdPurcSale = GridRetr4;
            }

            gdPurcSale.DataSource = null;
            gdPurcSale.DataSource = dt;

            SetInitValue();
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

            //double.TryParse(GridCol2AcceptWeight.SummaryItem.SummaryValue?.ToString(), out dLandedWeight);
            //double.TryParse(GridCol3OWeight.SummaryItem.SummaryValue?.ToString(), out dAcptWeight);

            //double.TryParse(GridCol2TotalPrc.SummaryItem.SummaryValue?.ToString(), out dPurcAmt1);
            //double.TryParse(GridCol3PurcAmt.SummaryItem.SummaryValue?.ToString(), out dPurcAmt2);

            //double.TryParse(GridCol2CarryCost.SummaryItem.SummaryValue?.ToString(), out dCarryCost1);
            //double.TryParse(GridCol3CarryCost.SummaryItem.SummaryValue?.ToString(), out dCarryCost2);

            //double.TryParse(GridCol2ArrvTotalPrc.SummaryItem.SummaryValue?.ToString(), out dArrvAmt1);
            //double.TryParse(GridCol3ArrvTotPrc.SummaryItem.SummaryValue?.ToString(), out dArrvAmt2);

            string sWeight = string.Format("인수량 : {0}", (dLandedWeight + dAcptWeight).ToString("n0"));
            string sPurcAmt = string.Format("매입액 : {0}", (dPurcAmt1 + dPurcAmt2).ToString("n0"));
            string sCarryCost = string.Format("운반비 : {0}", (dCarryCost1 + dCarryCost2).ToString("n0"));
            string sArrvAmt = string.Format("도착도금액 : {0}", (dArrvAmt1 + dArrvAmt2).ToString("n0"));

            LblTotal.Text = string.Format("{0}, {1}, {2}, {3}", sWeight, sPurcAmt, sCarryCost, sArrvAmt);
        }

        private DataTable GetPurchaseInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.J_DATE AS CLOSE_DT ");
            strSql.AppendLine("      , G.J_DATE AS MESURE_DT  ");
            strSql.AppendLine("      , D.DEALER_NM  "); //2020-11-24 현업요청에 따라 계근일자 빼고 매입처로 변경
            strSql.AppendLine("      , B.J_BNUM  ");
            strSql.AppendLine("      , G.WEIGHT AS LANDEDWEIGHT  ");
            strSql.AppendLine("      , A.HALIN AS LOSS  ");
            strSql.AppendLine("      , A.IWEIGHT AS ACCEPTWEIGHT  ");
            strSql.AppendLine("      , IFNULL(A.IWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT  ");
            strSql.AppendLine("      , F.GUBUN1  ");
            strSql.AppendLine("      , A.MIDANGA AS STDDUNITPRICE  ");
            strSql.AppendLine("      , A.DANGA AS PAYEDUNITPRICE  ");
            //strSql.AppendLine("      , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE  ");
            strSql.AppendLine("      , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT - IFNULL(A.MIDANGA, 0) AS DIFFUNITPRICE  ");
            strSql.AppendLine("      , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE  ");
            strSql.AppendLine("      , IFNULL(A.CKONGKEP, 0) AS CARRYCOST  ");
            strSql.AppendLine("      , A.IKONGKEP AS TOTALPRICE  ");
            strSql.AppendLine("      , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE  ");
            strSql.AppendLine("      , G.J_STATE AS LOSSCAUSE  ");
            strSql.AppendLine("      , G.JUNPYOID AS IMAGE  ");
            strSql.AppendLine("      , G.GUMSUBIGO AS INSPECTNOTE  ");
            strSql.AppendLine("     , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE  ");
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
            strSql.AppendLine("    ON E.EMP_ID = D.CHRG_ID  ");
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
            strSql.AppendLine("   AND A.J_ID1 = " +  dicParams["J_ID"] + " ");
            strSql.AppendLine(" ORDER BY A.J_DATE, D.DEALER_NM ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable SalesPerBuyer(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

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
            strSql.AppendLine("               , IFNULL(A.OWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            strSql.AppendLine("               , F.GUBUN1 ");
            strSql.AppendLine("               , A.MIDANGA AS STDDUNITPRICE ");
            strSql.AppendLine("               , A.DANGA AS SALEUNITPRICE ");
            strSql.AppendLine("               , ((IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE ");
            strSql.AppendLine("               , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            strSql.AppendLine("               , A.KONGKEP AS TOTALPRICE ");
            strSql.AppendLine("               , (IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("               , G.JUNPYOID AS IMAGE ");
            strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE ");
            strSql.AppendLine("               , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("               , G.ETC_COST1 ");
            strSql.AppendLine("               , G.ETC_REMARK1 ");
            strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2");
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
            strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            strSql.AppendLine("              ON G.IPCHULGO_MACHULID = A.J_ID ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("           WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("             AND A.KERATYPE = '매출' ");
            strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            strSql.AppendLine("             AND G.J_CHECK = '1' ");
            strSql.AppendLine("             AND A.DANGA != 0    ");
            strSql.AppendLine("             AND A.J_ID1 = " + dicParams["J_ID"] + " ");
            strSql.AppendLine("      ) Y1 ");
            strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");
            
            Cursor = Cursors.Default;
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
        }

        private DataTable GetDirectSendInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT B.J_DATE ");
            strSql.AppendLine("      , D.EMP_NM AS NAME ");
            strSql.AppendLine("      , H.GUBUN1 ");
            strSql.AppendLine("      , A.J_BNUM ");
            strSql.AppendLine("      , D.DEALER_NM AS SALE_DEALER_NM ");
            strSql.AppendLine("      , B.DANJUNG AS OWEIGHT ");
            strSql.AppendLine("      , B.DANGA AS SALEUNITPRICE ");
            strSql.AppendLine("      , B.KONGKEP AS SALEPRICE ");
            strSql.AppendLine("      , E.DEALER_NM AS PURC_DEALER_NM ");
            strSql.AppendLine("      , C.DANGA AS PURCHUNITPRICE ");
            strSql.AppendLine("      , C.IKONGKEP AS PURC_AMT ");
            strSql.AppendLine("      , TRUNCATE(((IFNULL(B.DANJUNG, 0) * IFNULL(C.DANGA, 0)) + IFNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1) AS ARRIVEUNITPRICE  ");
            strSql.AppendLine("      , IFNULL(B.CKONGKEP, 0) AS CARRYCOST  ");
            strSql.AppendLine("      , TRUNCATE((IFNULL(B.CKONGKEP, 0) / B.DANJUNG), 1) AS CARRY_UNIT_PRICE  ");
            strSql.AppendLine("      , (IFNULL(C.DANJUNG, 0) * IFNULL(C.DANGA, 0)) + IFNULL(B.CKONGKEP, 0) AS ARRVTOTALPRICE  ");
            //strSql.AppendLine("      , IFNULL(B.DANGA, 0) - IFNULL(C.DANGA, 0) AS DIFFUNITPRICE ");
            strSql.AppendLine("      , IFNULL(B.DANGA, 0) - TRUNCATE(((IFNULL(B.DANJUNG, 0) * IFNULL(C.DANGA, 0)) + IFNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1) AS DIFFUNITPRICE ");
            strSql.AppendLine("      , C.HALIN ");
            strSql.AppendLine("      , I.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("      , A.ETC_COST1  ");
            strSql.AppendLine("      , A.ETC_REMARK1  ");
            strSql.AppendLine("      , J.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("      , A.ETC_COST2  ");
            strSql.AppendLine("      , A.ETC_REMARK2  ");
            strSql.AppendLine("   FROM MESURING A ");
            strSql.AppendLine("   LEFT OUTER JOIN INLIST B  ");
            strSql.AppendLine("     ON A.IPCHULGO_MACHULID = B.J_ID ");
            strSql.AppendLine("   LEFT OUTER JOIN INLIST C  ");
            strSql.AppendLine("     ON A.IPCHULGO_MAIPID = C.J_ID ");
            strSql.AppendLine("   LEFT OUTER JOIN ( SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID ");
            strSql.AppendLine("                       FROM ACC_DEALER_CD X1 ");
            strSql.AppendLine("                       LEFT OUTER JOIN HR_EMP_BASIS X2 ");
            strSql.AppendLine("                         ON X1.CHRG_ID = X2.EMP_ID ) D  ");
            strSql.AppendLine("     ON A.J_ASSIGNID = D.DEALER_CD #매출처 ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD E #매입처  ");
            strSql.AppendLine("     ON C.J_ID1 = E.DEALER_CD ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE H ");
            strSql.AppendLine("     ON A.J_SERIAL = H.J_SERIAL  ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD1 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD J          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD2 = CONVERT(VARCHAR,J.DEALER_CD)");
            strSql.AppendLine("  WHERE A.KERATYPE = '직송' ");
            strSql.AppendLine("    AND B.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("    AND C.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("    AND B.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("    AND C.J_ID1 = " + dicParams["J_ID"] + " ");
            strSql.AppendLine("  ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void IN09001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                DropBtnExcel.PerformClick();

        }

        private void GridViewRetr1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void CboFindSbj_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void DropBtnExcel_Click(object sender, EventArgs e)
        {
            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "매입처리스트")
            {
                GetDealerInfo();
            }
            else if (tag == "입고리스트")
            {
                GetPurcInfo();
            }
            else if (tag == "직송리스트")
            {
                GetDirectSendInfo();
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnDealerInfo;
        BarButtonItem BtnPurcList;
        BarButtonItem BtnDirectSendList;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnDealerInfo = new BarButtonItem(barManager1, "매입처리스트");
            BtnPurcList = new BarButtonItem(barManager1, "입고리스트");
            BtnDirectSendList = new BarButtonItem(barManager1, "직송리스트");
            popupMenu1.AddItem(BtnDealerInfo);
            popupMenu1.AddItem(BtnPurcList);
            popupMenu1.AddItem(BtnDirectSendList);

            DropBtnExcel.DropDownControl = popupMenu1;

            BtnDealerInfo.Tag = "매입처리스트";
            BtnDealerInfo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnDealerInfo_ItemClick);

            BtnPurcList.Tag = "입고리스트";
            BtnPurcList.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnPurcList_ItemClick);

            BtnDirectSendList.Tag = "직송리스트";
            BtnDirectSendList.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnDirectSendList_ItemClick);
        }

        private void BtnDealerInfo_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            GetDealerInfo();
        }

        private void BtnPurcList_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            GetPurcInfo();
        }

        private void BtnDirectSendList_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            GetDirectSendInfo();
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnExcel.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnExcel.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnExcel.Tag = submenuItem.Tag;
            //DropBtnExcel.Text = string.Format("{0}", submenuItem.Tag);
        }

        private void GetDealerInfo()
        {
            if(GridViewRetr1.RowCount == 0)
            {
                XtraMessageBox.Show("매입처리스트에 자료가 존재하지 않습니다.");
                return;
            }
            
            using (var saveDialog = new SaveFileDialog())
            {
                string sFileNM = string.Format("{0}_{1}", this.Text, DateTime.Now.ToLongDateString().Replace(" ", ""));
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                saveDialog.Filter = "Excel(.xlsx) | *.xlsx";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    var printingSystem = new PrintingSystemBase();
                    printingSystem.XlSheetCreated += PrintingSystem_XlsxDocumentCreated;
                    var compositeLink = new CompositeLinkBase();
                    compositeLink.PrintingSystemBase = printingSystem;

                    var link1 = new PrintableComponentLinkBase();
                    link1.Component = GridRetr2;
                    
                    var link2 = new PrintableComponentLinkBase();
                    //link2.Component = GridRetr3;

                    compositeLink.Links.Add(link1);
                    compositeLink.Links.Add(link2);

                    printingSystem.Document.AutoFitToPagesWidth = 1;

                    var options = new XlsxExportOptions();
                    options.ExportMode = XlsxExportMode.SingleFilePageByPage;



                    compositeLink.CreatePageForEachLink();
                    
                    compositeLink.ExportToXlsx(saveDialog.FileName, options);
                }
            }
            
        }

        void PrintingSystem_XlsxDocumentCreated(object sender, DevExpress.XtraPrinting.XlSheetCreatedEventArgs e)
        {
            if(e.Index == 0)
            {
                e.SheetName = "입고리스트";
            }
            else if(e.Index == 1)
            {
                e.SheetName = "직송리스트";
            }
        }

        /// <summary>
        /// Byte를 File로 변환합니다.
        /// </summary>
        /// <author> Kim Se Hoon </author>
        /// <param name="source"> 대상 byte[] </param>
        /// <param name="filename"> 파일명 (상대 or 절대)</param>
        private void ByteToFile(byte[] source, string filename)
        {
            /// Create Mode로 FileStream을 오픈합니다.
            FileStream file = new FileStream(filename, FileMode.Create);
            
            /// Byte에 있는 내용을 File에 씁니다.
            file.Write(source, 0, source.Length);
            
            /// 파일을 닫습니다.
            file.Close();
        }

        private void GetPurcInfo()
        {
            if (BGridViewRetr2.RowCount == 0)
            {
                XtraMessageBox.Show("입고리스트에 자료가 존재하지 않습니다.");
                return;
            }

            ComnEtcFunc.ExportExcelFile(string.Format("{0}_입고리스트_", this.Text), GridRetr2);
        }

        private void GetDirectSendInfo()
        {
            //if (BGridViewRetr3.RowCount == 0)
            //{
            //    XtraMessageBox.Show("직송리스트에 자료가 존재하지 않습니다.");
            //    return;
            //}

            //ComnEtcFunc.ExportExcelFile(string.Format("{0}_직송리스트_", this.Text), GridRetr3);
        }

        private void IN09001F01_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

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