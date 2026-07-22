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
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.Data;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DevExpress.XtraGrid;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-01-25
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            등급별, 매입처별 리스트는 GROUP BY J_DATE 제외
 *            
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
 *            엑셀 출력 시 같은파일명으로 저장하여 덮어씌워 바로 볼수 있도록 세팅
 *            기존에는 파일명 뒤 생성일자를 붙혀 파일이 여러개 생성되는 현상이어 수정요청
 *            
 * 수정일자 : 2021-02-23
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            INLIST에 이력성 컬럼인 CHRG_ID(담당자)로 업체담당자 표기하도록 쿼리 수정'
 *            
 *            
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2022-06-21
 * 수정자   : 정은영
 * ID       : #001
 * 수정내용 : (현업요청)
 *            1. 부대비용 합계 표시 (기타비용1+기타비용2)
 *            2. 더블클릭시 상세내용 보기
 *            3. (수정) 도착도단가 = (금액+운반비+부대비용)/인수량
 */
namespace AccAdm
{
    public partial class AccFieldPSDailyRpt : DevExpress.XtraEditors.XtraForm
    {
        public AccFieldPSDailyRpt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccFieldPSDailyRpt_Load(object sender, EventArgs e)
        {
            //
            

            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            SetGridLookupEdit(GLkupEditCate, "N", "1");
            GLkupEditCate.Properties.View.PopulateColumns(GLkupEditCate.Properties.DataSource);
            GLkupEditCate.Properties.View.Columns[GLkupEditCate.Properties.ValueMember].Visible = false;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewBuyerPurc, GridViewBuyerSales, GridViewBuyerDtSend, GridViewGradePurc, GridViewGradeSales, GridViewCompany };
            foreach(GridView view in arrGrdView)
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
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }

            ClearFps();

            string sYmd = DateEditFrom.EditValue.ToString();

            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);
            string sDealGB = RdGbDealGB.EditValue.ToString();

            if (sDealGB == "A")
            {
                sDealGB = "매입";
            }
            else if (sDealGB == "B")
            {
                sDealGB = "매출";
            }
            else
            {
                sDealGB = "직송";
            }

            string sBigCate = string.Empty;

            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.ValueMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0)
            {
                sBigCate = sBigCate.Substring(0, sBigCate.Length - 1);
            }
            else
            {
                sBigCate = null;
            }

            if (RdGbDealGB.SelectedIndex == 0)
            {
                PurchacePerBuyer(sYmdFrom, sYmdTo, sDealGB, sBigCate);
                PurchasePerGrade(sYmdFrom, sYmdTo, sDealGB, sBigCate);
            }
            else if (RdGbDealGB.SelectedIndex == 1)
            {
                SalesPerBuyer(sYmdFrom, sYmdTo, sDealGB, sBigCate);
                SalesPerGrade(sYmdFrom, sYmdTo, sDealGB, sBigCate);
            }
            else if (RdGbDealGB.SelectedIndex == 2)
            {
                DirectSendPerBuyer(sYmdFrom, sYmdTo, sDealGB, sBigCate);
                DirectSendPerGrade(sYmdFrom, sYmdTo, sDealGB, sBigCate);
            }

            XtTControl_SelectedPageChanged(null, null);
        }

        private void RdoDealGB_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void PurchacePerBuyer(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            //#001
            strSql.AppendLine(" SELECT *                                                                 ");
            strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE                    ");
            strSql.AppendLine("   FROM(                                                                  ");
            strSql.AppendLine("          SELECT A.J_DATE                                                 ");
            strSql.AppendLine("               , A.JUNPYOID                                               ");
            strSql.AppendLine("               , G.J_DATE AS MESURE_DT                                    ");
            strSql.AppendLine("               , D.DEALER_NM                                              ");
            strSql.AppendLine("               , E.EMP_NM AS INSPECTOR                                    ");
            strSql.AppendLine("               , G.J_BNUM                                                 ");
            strSql.AppendLine("               , G.WEIGHT AS LANDEDWEIGHT                                 ");
            strSql.AppendLine("               , A.HALIN AS LOSS                                          ");
            strSql.AppendLine("               , A.IWEIGHT AS ACCEPTWEIGHT                                ");
            strSql.AppendLine("               , ISNULL(A.IWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            strSql.AppendLine("               , F.GUBUN1                                                 ");
            strSql.AppendLine("               , A.MIDANGA AS STDDUNITPRICE                               ");
            strSql.AppendLine("               , A.DANGA AS PAYEDUNITPRICE                                ");
            strSql.AppendLine("               , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) + (ISNULL(G.ETC_COST1, 0) + ISNULL(G.ETC_COST2, 0))) / A.IWEIGHT AS ARRVUNITPRICE");
            strSql.AppendLine("               , ISNULL(A.CKONGKEP, 0) AS CARRYCOST                                                    ");
            strSql.AppendLine("               , A.IKONGKEP AS TOTALPRICE                                                              ");
            strSql.AppendLine("               , (ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE                                                                ");
            strSql.AppendLine("               , G.JUNPYOID AS IMAGE                                                                   ");
            strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE                                                            ");
            strSql.AppendLine("               , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG), 1, 0) AS CARRY_UNIT_PRICE                  ");
            strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("               , ISNULL(G.ETC_COST1, 0) AS ETC_COST1                                                   ");
            strSql.AppendLine("               , G.ETC_REMARK1                                                                         ");
            strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("               , ISNULL(G.ETC_COST2, 0) AS ETC_COST2                                                   ");
            strSql.AppendLine("               , ISNULL(G.ETC_COST1, 0) + ISNULL(G.ETC_COST1, 0) AS ETC_COST                           ");
            strSql.AppendLine("               , G.ETC_REMARK2                                                                         ");
            strSql.AppendLine("            FROM INLIST A                                                                              ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE C                                                                    ");
            strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL                                                               ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D                                                            ");
            strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1                                                                 ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E                                                             ");
            strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID                                                                  ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE F                                                                    ");
            strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1                                                                   ");
            strSql.AppendLine("            LEFT OUTER JOIN MESURING G                                                                 ");
            strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID                                                            ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("           WHERE A.J_DATE >= '"+ sYmdFrom + "'                                                              ");
            strSql.AppendLine("             AND A.J_DATE <= '"+ sYmdTo + "'                                                              ");
            strSql.AppendLine("             AND A.KERATYPE = '"+ jType + "'                                                                   ");
            strSql.AppendLine("             AND G.KERATYPE <> '직송'                                                                  ");
            strSql.AppendLine("             AND G.J_CHECK = '1'                                                                       ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine("      ) Y1 ");
            strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");

            #region 2022-06-21 이전 부대비용 표기 전
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , A.JUNPYOID ");
            //strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("	              , D.DEALER_NM ");
            //strSql.AppendLine("	              , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("	              , G.J_BNUM ");
            //strSql.AppendLine("	              , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("	              , A.HALIN AS LOSS ");
            //strSql.AppendLine("	              , A.IWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("	              , ISNULL(A.IWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("	              , F.GUBUN1 ");
            //strSql.AppendLine("	              , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("	              , A.DANGA AS PAYEDUNITPRICE ");
            ////strSql.AppendLine("	              , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("	              , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("	              , ISNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("	              , A.IKONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("	              , (ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("	              , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("	              , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("	              , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("               , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG),1,0) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("               , G.ETC_COST2 ");
            //strSql.AppendLine("               , G.ETC_REMARK2 ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("              ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("              ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("             AND G.KERATYPE <> '직송'  ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");
            #endregion

            #region mariaDB
            /*
             * INLIST CHRG_ID 추가됨에 따라 해당 컬럼으로 업체담당자 JOIN
             */
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , A.JUNPYOID ");
            //strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("	              , D.DEALER_NM ");
            //strSql.AppendLine("	              , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("	              , B.J_BNUM ");
            //strSql.AppendLine("	              , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("	              , A.HALIN AS LOSS ");
            //strSql.AppendLine("	              , A.IWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("	              , IFNULL(A.IWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("	              , F.GUBUN1 ");
            //strSql.AppendLine("	              , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("	              , A.DANGA AS PAYEDUNITPRICE ");
            ////strSql.AppendLine("	              , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("	              , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("	              , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("	              , A.IKONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("	              , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("	              , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("	              , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("	              , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("               , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("               , G.ETC_COST2 ");
            //strSql.AppendLine("               , G.ETC_REMARK2 ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("	             ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("              ON E.EMP_ID = A.CHRG_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("              ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("              ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");
            #endregion

            #region[2021-02-23 이전쿼리]

            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , A.JUNPYOID ");
            //strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("	              , D.DEALER_NM ");
            //strSql.AppendLine("	              , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("	              , B.J_BNUM ");
            //strSql.AppendLine("	              , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("	              , A.HALIN AS LOSS ");
            //strSql.AppendLine("	              , A.IWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("	              , IFNULL(A.IWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("	              , F.GUBUN1 ");
            //strSql.AppendLine("	              , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("	              , A.DANGA AS PAYEDUNITPRICE ");
            ////strSql.AppendLine("	              , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("	              , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("	              , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("	              , A.IKONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("	              , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("	              , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("	              , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("	              , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("               , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("               , G.ETC_COST2 ");
            //strSql.AppendLine("               , G.ETC_REMARK2 ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("	             ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("              ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("              ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");

            #endregion[2021-02-23 이전쿼리]

            #region[2021-02-01 이전 쿼리]

            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , A.JUNPYOID ");
            //strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("	              , D.DEALER_NM ");
            //strSql.AppendLine("	              , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("	              , B.J_BNUM ");
            //strSql.AppendLine("	              , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("	              , A.HALIN AS LOSS ");
            //strSql.AppendLine("	              , A.IWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("	              , IFNULL(A.IWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("	              , F.GUBUN1 ");
            //strSql.AppendLine("	              , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("	              , A.DANGA AS PAYEDUNITPRICE ");
            //strSql.AppendLine("	              , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("	              , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("	              , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("	              , A.IKONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("	              , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("	              , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("	              , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("	              , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("               , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("               , G.ETC_COST2 ");
            //strSql.AppendLine("               , G.ETC_REMARK2 ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("	             ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("              ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("              ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            #endregion[2021-02-01 이전 쿼리]

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerPurc;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;

        }

        private void PurchasePerGrade(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            #region mariaDB
            //strSql.AppendLine(" SELECT A.J_DATE ");
            //strSql.AppendLine(" 	  , E.GUBUN1 ");
            //strSql.AppendLine(" 	  , SUM(A.IWEIGHT) AS ACCEPTWEIGHT ");
            //strSql.AppendLine(" 	  , SUM(A.IWEIGHT) - SUM(F.WEIGHT) AS DIFFWEIGHT ");
            //strSql.AppendLine(" 	  , SUM(F.WEIGHT) AS LANDEDWEIGHT ");
            //strSql.AppendLine(" 	  , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine(" 	  , A.DANGA AS PAYEDUNITPRICE ");
            //strSql.AppendLine(" 	  , SUM(A.IKONGKEP) AS TOTALPRICE ");
            //strSql.AppendLine(" 	  , SUM(A.HALIN) AS HALIN ");
            //strSql.AppendLine("   FROM INLIST A ");
            //strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("   	 ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE E ");
            //strSql.AppendLine("     ON E.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN MESURING F ");
            //strSql.AppendLine("     ON F.IPCHULGO_MAIPID = A.J_ID ");
            //strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.KERATYPE = '매입' ");
            //strSql.AppendLine("    AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("    AND F.J_CHECK = '1' ");
            //strSql.AppendLine("    AND A.DANGA != 0    ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  GROUP BY A.J_SERIAL ");
            //strSql.AppendLine("  ORDER BY E.GUBUN1 ");
            ////strSql.AppendLine("  GROUP BY A.J_DATE, E.GUBUN1 ");
            ////strSql.AppendLine("  ORDER BY A.J_DATE, E.GUBUN1 ");
            #endregion

            strSql.AppendLine("SELECT E.GUBUN1                                     ");
 	        strSql.AppendLine("     , SUM(A.IWEIGHT) AS ACCEPTWEIGHT               ");
            strSql.AppendLine("     , SUM(A.IWEIGHT) -SUM(F.WEIGHT) AS DIFFWEIGHT  ");
            strSql.AppendLine("     , SUM(F.WEIGHT) AS LANDEDWEIGHT                ");
            strSql.AppendLine("     , ROUND(AVG(A.MIDANGA), 1, 0) AS STDDUNITPRICE ");
            strSql.AppendLine("     , ROUND(AVG(A.DANGA), 1, 0) AS PAYEDUNITPRICE  ");
            strSql.AppendLine("     , SUM(A.IKONGKEP) AS TOTALPRICE                ");
            strSql.AppendLine("     , SUM(A.HALIN) AS HALIN                        ");
            strSql.AppendLine("  FROM INLIST A                                     ");
            strSql.AppendLine("  LEFT OUTER JOIN IPCHULGO B                        ");
            strSql.AppendLine("    ON B.J_ID = A.J_ID                              ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE C                           ");
            strSql.AppendLine("    ON C.J_SERIAL = A.J_SERIAL                      ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D                   ");
            strSql.AppendLine("    ON D.DEALER_CD = A.J_ID1                        ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE E                           ");
            strSql.AppendLine("    ON E.GUBUN1 = C.GUBUN1                          ");
            strSql.AppendLine("  LEFT OUTER JOIN MESURING F                        ");
            strSql.AppendLine("    ON F.IPCHULGO_MAIPID = A.J_ID                   ");
            strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
            strSql.AppendLine("    AND A.KERATYPE = '매입' ");
            strSql.AppendLine("    AND F.KERATYPE <> '직송' ");
            strSql.AppendLine("    AND F.J_CHECK = '1' ");
            strSql.AppendLine("    AND A.DANGA != 0    ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine(" GROUP BY A.J_SERIAL, E.GUBUN1 ");
            strSql.AppendLine(" ORDER BY E.GUBUN1             ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridGrade.MainView = GridViewGradePurc;
            GridGrade.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void SalesPerBuyer(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
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
            strSql.AppendLine("               , G.J_BNUM ");
            strSql.AppendLine("               , G.WEIGHT AS LANDEDWEIGHT ");
            strSql.AppendLine("               , A.HALIN AS LOSS ");
            strSql.AppendLine("               , A.OWEIGHT AS ACCEPTWEIGHT ");
            strSql.AppendLine("               , ISNULL(A.OWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            strSql.AppendLine("               , F.GUBUN1 ");
            strSql.AppendLine("               , A.MIDANGA AS STDDUNITPRICE ");
            strSql.AppendLine("               , A.DANGA AS SALEUNITPRICE ");
            strSql.AppendLine("               , ((ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE ");
            strSql.AppendLine("               , ISNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            strSql.AppendLine("               , A.KONGKEP AS TOTALPRICE ");
            strSql.AppendLine("               , (ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("               , G.JUNPYOID AS IMAGE ");
            strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE ");
            strSql.AppendLine("               , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG), 1,0) AS CARRY_UNIT_PRICE ");
            strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("               , G.ETC_COST1 ");
            strSql.AppendLine("               , G.ETC_REMARK1 ");
            strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("               , G.ETC_COST2 ");
            strSql.AppendLine("               , G.ETC_REMARK2 ");
            strSql.AppendLine("            FROM INLIST A ");
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
            strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            strSql.AppendLine("             AND G.KERATYPE <> '직송' ");
            strSql.AppendLine("             AND G.J_CHECK = '1' ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine("      ) Y1 ");
            strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");

            #region mariaDB
            /*
             * INLIST CHRG_ID 추가됨에 따라 해당 컬럼으로 업체담당자 JOIN
             */
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
            //strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");
            #endregion

            #region[2021-02-23 이전쿼리]

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
            //strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("              ON G.IPCHULGO_MACHULID = A.J_ID ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("              ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("              ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM ");

            #endregion[2021-02-23]

            #region[2021-02-01 이전 쿼리]

            //strSql.AppendLine(" SELECT A.J_DATE ");
            //strSql.AppendLine("      , A.JUNPYOID ");
            //strSql.AppendLine("      , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("      , D.DEALER_NM ");
            //strSql.AppendLine("      , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("      , B.J_BNUM ");
            //strSql.AppendLine("      , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("      , A.HALIN AS LOSS ");
            //strSql.AppendLine("      , A.OWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("      , IFNULL(A.OWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("      , F.GUBUN1 ");
            //strSql.AppendLine("      , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("      , A.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine("      , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , ((IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("      , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("      , A.KONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("      , (IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("      , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("      , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("      , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("      , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("      , H.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("      , G.ETC_COST1 ");
            //strSql.AppendLine("      , G.ETC_REMARK1 ");
            //strSql.AppendLine("      , I.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("      , G.ETC_COST2 ");
            //strSql.AppendLine("      , G.ETC_REMARK2 ");
            //strSql.AppendLine("   FROM INLIST A ");
            //strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("     ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("     ON E.EMP_ID = D.CHRG_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("     ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("     ON G.IPCHULGO_MACHULID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD H");
            //strSql.AppendLine("     ON H.DEALER_CD = G.ETC_DEALER_CD1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD I");
            //strSql.AppendLine("     ON I.DEALER_CD = G.ETC_DEALER_CD2 ");
            //strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.KERATYPE = '" + jType + "' ");
            //strSql.AppendLine("    AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("    AND G.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            #endregion[2021-02-01 이전 쿼리]

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerSales;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void SalesPerGrade(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            strSql.AppendLine(" SELECT E.GUBUN1 ");
            strSql.AppendLine(" 	  , SUM(A.OWEIGHT) AS ACCEPTWEIGHT ");
            strSql.AppendLine(" 	  , SUM(A.OWEIGHT) - SUM(F.WEIGHT) AS DIFFWEIGHT ");
            strSql.AppendLine(" 	  , SUM(F.WEIGHT) AS LANDEDWEIGHT ");
            strSql.AppendLine(" 	  , ROUND(AVG(A.MIDANGA),1,0) AS STTDUNITPRICE ");
            strSql.AppendLine(" 	  , ROUND(AVG(A.DANGA),1,0) AS SALEUNITPRICE ");
            strSql.AppendLine(" 	  , SUM(A.KONGKEP) AS TOTALPRICE ");
            strSql.AppendLine(" 	  , SUM(A.HALIN) AS HALIN ");
            strSql.AppendLine("   FROM INLIST A ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE E ");
            strSql.AppendLine("     ON E.GUBUN1 = C.GUBUN1 ");
            strSql.AppendLine("   LEFT OUTER JOIN MESURING F ");
            strSql.AppendLine("     ON F.IPCHULGO_MACHULID = A.J_ID ");
            strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            strSql.AppendLine("    AND F.KERATYPE <> '직송' ");
            strSql.AppendLine("    AND F.J_CHECK = '1' ");
            strSql.AppendLine("    AND A.DANGA != 0    ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine("  GROUP BY A.J_SERIAL, E.GUBUN1 ");
            strSql.AppendLine("  ORDER BY E.GUBUN1 ");

            #region mariaDB
            //strSql.AppendLine(" SELECT A.J_DATE ");
            //strSql.AppendLine(" 	  , E.GUBUN1 ");
            //strSql.AppendLine(" 	  , SUM(A.OWEIGHT) AS ACCEPTWEIGHT ");
            //strSql.AppendLine(" 	  , SUM(A.OWEIGHT) - SUM(F.WEIGHT) AS DIFFWEIGHT ");
            //strSql.AppendLine(" 	  , SUM(F.WEIGHT) AS LANDEDWEIGHT ");
            //strSql.AppendLine(" 	  , A.MIDANGA AS STTDUNITPRICE ");
            //strSql.AppendLine(" 	  , A.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine(" 	  , SUM(A.KONGKEP) AS TOTALPRICE ");
            //strSql.AppendLine(" 	  , SUM(A.HALIN) AS HALIN ");
            //strSql.AppendLine("   FROM INLIST A ");
            //strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("   	 ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE E ");
            //strSql.AppendLine("     ON E.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN MESURING F ");
            //strSql.AppendLine("     ON F.IPCHULGO_MACHULID = A.J_ID ");
            //strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            //strSql.AppendLine("    AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("    AND F.J_CHECK = '1' ");
            //strSql.AppendLine("    AND A.DANGA != 0    ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  GROUP BY A.J_SERIAL ");
            //strSql.AppendLine("  ORDER BY E.GUBUN1 ");
            ////strSql.AppendLine("  GROUP BY A.J_DATE, E.GUBUN1 ");
            ////strSql.AppendLine("  ORDER BY A.J_DATE, E.GUBUN1 ");
            #endregion

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());


            GridGrade.MainView = GridViewGradeSales;
            GridGrade.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void DirectSendPerBuyer(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            strSql.AppendLine("SELECT B.J_DATE                       ");
            strSql.AppendLine("     , XX.EMP_NM AS NAME              ");
            strSql.AppendLine("     , H.GUBUN1                       ");
            strSql.AppendLine("     , A.J_BNUM                       ");
            strSql.AppendLine("     , D.DEALER_NM AS SALE_DEALER_NM  ");
            strSql.AppendLine("     , B.DANJUNG AS OWEIGHT           ");
            strSql.AppendLine("     , B.DANGA AS SALEUNITPRICE       ");
            strSql.AppendLine("     , B.KONGKEP AS SALEPRICE         ");
            strSql.AppendLine("     , E.DEALER_NM AS PURC_DEALER_NM  ");
            strSql.AppendLine("     , C.DANGA AS PURCHUNITPRICE      ");
            strSql.AppendLine("     , C.IKONGKEP AS PURC_AMT         ");
            strSql.AppendLine("     , ROUND(((ISNULL(B.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1, 0) AS ARRIVEUNITPRICE");
            strSql.AppendLine("     , ISNULL(B.CKONGKEP, 0) AS CARRYCOST                                                                               ");
            strSql.AppendLine("     , ROUND((ISNULL(B.CKONGKEP, 0) / B.DANJUNG), 1, 0) AS CARRY_UNIT_PRICE                                             ");
            strSql.AppendLine("     , (ISNULL(C.DANJUNG, 0) * ISNULL(C.DANGA, 0)) +ISNULL(B.CKONGKEP, 0) AS ARRVTOTALPRICE                             ");
            strSql.AppendLine("     , ISNULL(B.DANGA, 0) -ROUND(((ISNULL(B.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1, 0) AS DIFFUNITPRICE");
            strSql.AppendLine("     , C.HALIN                      ");
            strSql.AppendLine("     , I.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("     , A.ETC_COST1                  ");
            strSql.AppendLine("     , A.ETC_REMARK1                ");
            strSql.AppendLine("     , J.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("     , A.ETC_COST2                  ");
            strSql.AppendLine("     , A.ETC_REMARK2                ");
            strSql.AppendLine("  FROM MESURING A                   ");
            strSql.AppendLine("  LEFT OUTER JOIN INLIST B          ");
            strSql.AppendLine("    ON A.IPCHULGO_MACHULID = B.J_ID ");
            strSql.AppendLine("  LEFT OUTER JOIN INLIST C          ");
            strSql.AppendLine("    ON A.IPCHULGO_MAIPID = C.J_ID   ");
            strSql.AppendLine("  LEFT OUTER JOIN(SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID ");
            strSql.AppendLine("                    FROM ACC_DEALER_CD X1                                  ");
            strSql.AppendLine("                    LEFT OUTER JOIN HR_EMP_BASIS X2                        ");
            strSql.AppendLine("                      ON X1.CHRG_ID = X2.EMP_ID) D                         ");
            strSql.AppendLine("   ON A.J_ASSIGNID = D.DEALER_CD--매출처                                   ");
            strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD E --매입처                                  ");
            strSql.AppendLine("   ON C.J_ID1 = E.DEALER_CD                                                ");
            strSql.AppendLine(" LEFT OUTER JOIN(SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID  ");
            strSql.AppendLine("                     FROM ACC_DEALER_CD X1            ");
            strSql.AppendLine("                     LEFT OUTER JOIN HR_EMP_BASIS X2  ");
            strSql.AppendLine("                       ON X1.CHRG_ID = X2.EMP_ID) E1  ");
            strSql.AppendLine("   ON A.MAIPCHERID = E1.DEALER_CD--매입처             ");
            strSql.AppendLine(" LEFT OUTER JOIN HR_EMP_BASIS XX                      ");
            strSql.AppendLine("   ON C.CHRG_ID = XX.EMP_ID                           ");
            strSql.AppendLine(" LEFT OUTER JOIN JAJAE H                              ");
            strSql.AppendLine("   ON A.J_SERIAL = H.J_SERIAL                         ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD1 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD J          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD2 = CONVERT(VARCHAR,J.DEALER_CD)");
            strSql.AppendLine("WHERE A.KERATYPE = '직송'                             ");
            strSql.AppendLine("  AND B.JUNPYOID IS NOT NULL                          ");
            strSql.AppendLine("  AND C.JUNPYOID IS NOT NULL                          ");
            strSql.AppendLine("  AND B.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND H.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine("ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");

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
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD E #매입처  ");
            //strSql.AppendLine("    ON C.J_ID1 = E.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ( SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID ");
            //strSql.AppendLine("                      FROM ACC_DEALER_CD X1 ");
            //strSql.AppendLine("                      LEFT OUTER JOIN HR_EMP_BASIS X2 ");
            //strSql.AppendLine("                        ON X1.CHRG_ID = X2.EMP_ID ) E1  ");
            //strSql.AppendLine("    ON A.MAIPCHERID = E1.DEALER_CD #매입처 ");
            //strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS XX ");
            //strSql.AppendLine("    ON C.CHRG_ID = XX.EMP_ID ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD F ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD1 = F.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD G ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD2 = G.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN JAJAE H ");
            //strSql.AppendLine("    ON A.J_SERIAL = H.J_SERIAL  ");
            //strSql.AppendLine(" WHERE A.KERATYPE = '직송' ");
            //strSql.AppendLine("   AND B.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("   AND C.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("   AND B.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND H.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");
            #endregion

            #region[2021-02-23 이전쿼리]

            //2020-07-21 수정 쿼리 (고혜성작성)
            //strSql.AppendLine(" SELECT B.J_DATE ");
            //strSql.AppendLine("      , E1.EMP_NM AS NAME ");
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
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD E #매입처  ");
            //strSql.AppendLine("    ON C.J_ID1 = E.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ( SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID ");
            //strSql.AppendLine("                      FROM ACC_DEALER_CD X1 ");
            //strSql.AppendLine("                      LEFT OUTER JOIN HR_EMP_BASIS X2 ");
            //strSql.AppendLine("                        ON X1.CHRG_ID = X2.EMP_ID ) E1  ");
            //strSql.AppendLine("    ON A.MAIPCHERID = E1.DEALER_CD #매입처 ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD F ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD1 = F.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD G ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD2 = G.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN JAJAE H ");
            //strSql.AppendLine("    ON A.J_SERIAL = H.J_SERIAL  ");
            //strSql.AppendLine(" WHERE A.KERATYPE = '직송' ");
            //strSql.AppendLine("   AND B.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("   AND C.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("   AND B.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND H.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");

            ////기존 직송 조회 쿼리
            //strSql.AppendLine(" SELECT A.J_DATE ");
            //strSql.AppendLine(" 	 , C.GUBUN1 ");
            //strSql.AppendLine(" 	 , I.EMP_NM AS NAME ");
            //strSql.AppendLine(" 	 , B.J_BNUM ");
            //strSql.AppendLine(" 	 , D.DEALER_NM ");
            //strSql.AppendLine("      , A.OWEIGHT");
            //strSql.AppendLine("	     , A.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine(" 	 , A.KONGKEP AS SALEPRICE ");
            //strSql.AppendLine(" 	 , A.J_BOOKING ");
            //strSql.AppendLine("      , G.DANGA AS PURCHUNITPRICE ");
            //strSql.AppendLine("	     , F.IKONGKEP ");
            ////strSql.AppendLine("	     , G.IKONGKEP ");
            //strSql.AppendLine(" 	 , IFNULL(A.DANGA, 0) - IFNULL(G.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine(" 	 , F.JUNPYOID ");
            //strSql.AppendLine(" 	 , ((IFNULL(A.OWEIGHT, 0) * IFNULL(G.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRIVEUNITPRICE ");
            //strSql.AppendLine(" 	 , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine(" 	 , (IFNULL(A.OWEIGHT, 0) * IFNULL(G.DANGA, 0)) + IFNULL(A.KONGKEP, 0) ARRVTOTALPRICE ");
            //strSql.AppendLine(" 	 , G.HALIN ");
            //strSql.AppendLine("      , TRUNCATE((IFNULL(A.CKONGKEP, 0) / G.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("      , J.DEALER_NM AS ETC_DEALER_NM1 ");
            //strSql.AppendLine("      , F.ETC_COST1 ");
            //strSql.AppendLine("      , F.ETC_REMARK1 ");
            //strSql.AppendLine("      , K.DEALER_NM AS ETC_DEALER_NM2 ");
            //strSql.AppendLine("      , F.ETC_COST2 ");
            //strSql.AppendLine("      , F.ETC_REMARK2 ");
            //strSql.AppendLine("   FROM INLIST A ");
            //strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("     ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE E ");
            //strSql.AppendLine("     ON E.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN MESURING F ");
            //strSql.AppendLine("     ON F.IPCHULGO_MACHULID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN INLIST G ");
            //strSql.AppendLine("     ON G.JUNPYOID = A.JUNPYOID ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD H ");
            //strSql.AppendLine("    	ON H.DEALER_CD = G.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS I ");
            //strSql.AppendLine("    	ON I.EMP_ID = H.CHRG_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD J");
            //strSql.AppendLine("     ON J.DEALER_CD = F.ETC_DEALER_CD1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD K");
            //strSql.AppendLine("     ON K.DEALER_CD = F.ETC_DEALER_CD2 ");
            //strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            //strSql.AppendLine("    AND F.KERATYPE = '직송' ");
            //strSql.AppendLine("    AND F.J_CHECK = '1' ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            //수정 조회쿼리 (정은영)
            //strSql.AppendLine("   SELECT DISTINCT A.J_DATE ");
            //strSql.AppendLine("  	   , A.GUBUN1  ");
            //strSql.AppendLine("  	   , F.CHRG_NM AS INSPECTOR ");
            //strSql.AppendLine("  	   , A.J_BNUM     ");
            //strSql.AppendLine("  	   , A.J_COMPANY AS DEALER_NM ");
            //strSql.AppendLine("        , D.OWEIGHT                 ");
            //strSql.AppendLine("  	   , D.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine("  	   , D.KONGKEP AS SALEPRICE   ");
            //strSql.AppendLine("  	   , A.MAIPCHER AS J_BOOKING  ");
            //strSql.AppendLine("        , C.DANGA AS PURCHUNITPRICE ");
            //strSql.AppendLine("  	   , C.IKONGKEP               ");
            //strSql.AppendLine("  	   , IFNULL(D.DANGA, 0) - IFNULL(C.DANGA, 0) AS DIFFUNITPRICE  ");
            //strSql.AppendLine("  	   , A.JUNPYOID  ");
            //strSql.AppendLine("  	   , ((IFNULL(D.OWEIGHT, 0) * IFNULL(C.DANGA, 0)) + IFNULL(D.CKONGKEP, 0)) / D.OWEIGHT AS ARRIVEUNITPRICE  ");
            //strSql.AppendLine("  	   , IFNULL(D.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("  	   , (IFNULL(D.OWEIGHT, 0) * IFNULL(C.DANGA, 0)) + IFNULL(D.KONGKEP, 0) ARRVTOTALPRICE ");
            //strSql.AppendLine("  	   , C.HALIN ");
            //strSql.AppendLine("        , TRUNCATE((IFNULL(D.CKONGKEP, 0) / C.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("        , G.DEALER_NM AS ETC_DEALER_NM1");
            //strSql.AppendLine("        , A.ETC_COST1                  ");
            //strSql.AppendLine("        , A.ETC_REMARK1                ");
            //strSql.AppendLine("        , H.DEALER_NM AS ETC_DEALER_NM2");
            //strSql.AppendLine("        , A.ETC_COST2                  ");
            //strSql.AppendLine("        , A.ETC_REMARK2                ");
            //strSql.AppendLine("     FROM MESURING A                   ");
            //strSql.AppendLine("     LEFT OUTER JOIN IPCHULGO B        ");
            //strSql.AppendLine("       ON A.IPCHULGO_MAIPID = B.J_ID   ");
            //strSql.AppendLine("       OR A.IPCHULGO_MACHULID = B.J_ID ");
            //strSql.AppendLine("     LEFT OUTER JOIN INLIST C          ");
            //strSql.AppendLine("       ON A.IPCHULGO_MAIPID = C.J_ID   ");
            //strSql.AppendLine("     LEFT OUTER JOIN INLIST D          ");
            //strSql.AppendLine("       ON A.IPCHULGO_MACHULID = D.J_ID ");
            //strSql.AppendLine("     LEFT OUTER JOIN JAJAE E           ");
            //strSql.AppendLine("       ON A.J_SERIAL = E.J_SERIAL      ");
            //strSql.AppendLine("     LEFT OUTER JOIN ACC_DEALER_CD F   ");
            //strSql.AppendLine("       ON A.MAIPCHERID = F.DEALER_CD   ");
            //strSql.AppendLine("     LEFT OUTER JOIN ACC_DEALER_CD G   ");
            //strSql.AppendLine("       ON A.ETC_DEALER_CD1 = G.DEALER_CD ");
            //strSql.AppendLine("     LEFT OUTER JOIN ACC_DEALER_CD H     ");
            //strSql.AppendLine("       ON A.ETC_DEALER_CD2 = H.DEALER_CD ");
            //strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("      AND A.KERATYPE = '직송'            ");
            //strSql.AppendLine("      AND A.J_CHECK = '1'                ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("    ORDER BY A.J_DATE, A.MAIPCHER        ");

            #endregion[2021-02-23 이전쿼리]

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerDtSend;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void DirectSendPerGrade(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            /*
             * #0001
             * 속도이슈로 쿼리변경
             * 2022-02-17
             */
            strSql.Clear();
            strSql.AppendLine(" SELECT A1.J_SERIAL                                                ");
            strSql.AppendLine("      , B1.GUBUN1                                                  ");
            strSql.AppendLine("      , SUM(A1.OWEIGHT) AS ACCEPTWEIGHT                            ");
            strSql.AppendLine("      , SUM(A2.WEIGHT) AS LANDEDWEIGHT                             ");
            strSql.AppendLine("      , SUM(A1.OWEIGHT) -SUM(A2.WEIGHT) AS DIFFWEIGHT              ");
            strSql.AppendLine("      , ROUND(AVG(A1.MIDANGA), 1, 1) AS STTDUNITPRICE              ");
            strSql.AppendLine("      , ROUND(AVG(A1.DANGA), 1, 1) AS SALEUNITPRICE                ");
            strSql.AppendLine("      , SUM(A1.KONGKEP) AS TOTALPRICE                              ");
            strSql.AppendLine("      , MAX(A3.HALIN) AS HALIN                                     ");
            strSql.AppendLine("   FROM INLIST A1                                                  ");
            strSql.AppendLine("   LEFT JOIN MESURING A2                                           ");
            strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MACHULID                             ");
            strSql.AppendLine("   LEFT JOIN(SELECT A1.J_SERIAL                                    ");
            strSql.AppendLine("                  , B1.GUBUN1                                      ");
            strSql.AppendLine("                  , SUM(A1.HALIN) AS HALIN                         ");
            strSql.AppendLine("               FROM INLIST A1                                      ");
            strSql.AppendLine("               LEFT JOIN MESURING A2                               ");
            strSql.AppendLine("                 ON A1.J_ID = A2.IPCHULGO_MAIPID                   ");
            strSql.AppendLine("               LEFT JOIN JAJAE B1                                  ");
            strSql.AppendLine("                 ON A1.J_SERIAL = B1.J_SERIAL                      ");
            strSql.AppendLine("              WHERE A1.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
            strSql.AppendLine("                AND A1.KERATYPE = '매입'                           ");
            strSql.AppendLine("                AND A2.KERATYPE = '직송'                           ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine("             GROUP BY A1.J_SERIAL, B1.GUBUN1 )A3                   ");
            strSql.AppendLine("     ON A1.J_SERIAL = A3.J_SERIAL                                  ");
            strSql.AppendLine("   LEFT JOIN JAJAE B1                                              ");
            strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                                  ");
            strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'            ");
            strSql.AppendLine("    AND A1.KERATYPE = '매출'                                       ");
            strSql.AppendLine("    AND A2.KERATYPE = '직송'                                       ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
            strSql.AppendLine("  GROUP BY A1.J_SERIAL, B1.GUBUN1                                  ");
            strSql.AppendLine("  ORDER BY B1.GUBUN1                                               ");

            #region 2022-02-17 이전코드
            //  strSql.AppendLine("SELECT E.GUBUN1                                    ");
            //strSql.AppendLine("    , SUM(A.OWEIGHT) AS ACCEPTWEIGHT               ");
            //  strSql.AppendLine("    , SUM(A.OWEIGHT) -SUM(F.WEIGHT) DIFFWEIGHT     ");
            //strSql.AppendLine("    , SUM(F.WEIGHT) AS LANDEDWEIGHT                ");
            //  strSql.AppendLine("    , ROUND(AVG(A.MIDANGA), 1, 0) AS STTDUNITPRICE ");
            //  strSql.AppendLine("    , ROUND(AVG(A.DANGA), 1, 0) AS SALEUNITPRICE   ");
            //  strSql.AppendLine("    , SUM(A.KONGKEP) AS TOTALPRICE                 ");
            //  strSql.AppendLine("    , SUM(G.HALIN) AS HALIN                        ");
            //  strSql.AppendLine(" FROM INLIST A                                     ");
            //  strSql.AppendLine(" LEFT OUTER JOIN IPCHULGO B                        ");
            //  strSql.AppendLine("   ON B.J_ID = A.J_ID                              ");
            //  strSql.AppendLine(" LEFT OUTER JOIN JAJAE C                           ");
            //  strSql.AppendLine("   ON C.J_SERIAL = A.J_SERIAL                      ");
            //  strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD D                   ");
            //  strSql.AppendLine("   ON D.DEALER_CD = A.J_ID1                        ");
            //  strSql.AppendLine(" LEFT OUTER JOIN JAJAE E                           ");
            //  strSql.AppendLine("   ON E.GUBUN1 = C.GUBUN1                          ");
            //  strSql.AppendLine(" LEFT OUTER JOIN MESURING F                        ");
            //  strSql.AppendLine("   ON F.IPCHULGO_MACHULID = A.J_ID                 ");
            //  strSql.AppendLine(" LEFT OUTER JOIN INLIST G                          ");
            //  strSql.AppendLine("   ON G.J_LOTNO = A.J_LOTNO                        ");
            //  strSql.AppendLine("  AND G.KERATYPE = '매입'                          ");
            //  strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //  strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            //  strSql.AppendLine("    AND B.KERAGUBUN = '직송' ");
            //  strSql.AppendLine("    AND F.J_CHECK = '1' ");
            //  if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND E.DAEGUBUN IN (" + sBigCate + ") ");
            //  strSql.AppendLine("  GROUP BY A.J_SERIAL,E.GUBUN1 ");
            //  strSql.AppendLine("  ORDER BY E.GUBUN1 ");
            #endregion

            #region mariaDB
            /*
             * #0001
             * 속도이슈로 쿼리변경
             * 2022-02-17
             */
            //strSql.Clear();
            //strSql.AppendLine(" SELECT A1.J_SERIAL                                               ");
            //strSql.AppendLine("      , B1.GUBUN1                                                  ");
            //strSql.AppendLine("      , SUM(A1.OWEIGHT) AS ACCEPTWEIGHT                            ");
            //strSql.AppendLine("      , SUM(A2.WEIGHT) AS LANDEDWEIGHT                             ");
            //strSql.AppendLine("      , SUM(A1.OWEIGHT) -SUM(A2.WEIGHT) AS DIFFWEIGHT              ");
            //strSql.AppendLine("      , TRUNCATE(AVG(A1.MIDANGA), 1) AS STTDUNITPRICE              ");
            //strSql.AppendLine("      , TRUNCATE(AVG(A1.DANGA), 1) AS SALEUNITPRICE                ");
            //strSql.AppendLine("      , SUM(A1.KONGKEP) AS TOTALPRICE                              ");
            //strSql.AppendLine("      , A3.HALIN                                                   ");
            //strSql.AppendLine("   FROM INLIST A1                                               ");
            //strSql.AppendLine("   LEFT JOIN MESURING A2                                            ");
            //strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MACHULID                            ");
            //strSql.AppendLine("   LEFT JOIN(SELECT A1.J_SERIAL                                     ");
            //strSql.AppendLine("                  , B1.GUBUN1                                      ");
            //strSql.AppendLine("                  , SUM(A1.HALIN) AS HALIN                         ");
            //strSql.AppendLine("               FROM INLIST A1                                   ");
            //strSql.AppendLine("               LEFT JOIN MESURING A2                            ");
            //strSql.AppendLine("                 ON A1.J_ID = A2.IPCHULGO_MAIPID                  ");
            //strSql.AppendLine("               LEFT JOIN JAJAE B1                               ");
            //strSql.AppendLine("                 ON A1.J_SERIAL = B1.J_SERIAL                     ");
            //strSql.AppendLine("              WHERE A1.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
            //strSql.AppendLine("                AND A1.KERATYPE = '매입'                         ");
            //strSql.AppendLine("                AND A2.KERATYPE = '직송'                         ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("             GROUP BY A1.J_SERIAL )A3                             ");
            //strSql.AppendLine("     ON A1.J_SERIAL = A3.J_SERIAL                                 ");
            //strSql.AppendLine("   LEFT JOIN JAJAE B1                                               ");
            //strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                                 ");
            //strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'            ");
            //strSql.AppendLine("    AND A1.KERATYPE = '매출'                                     ");
            //strSql.AppendLine("    AND A2.KERATYPE = '직송'                                     ");
            //if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
            //strSql.AppendLine("  GROUP BY A1.J_SERIAL                                             ");
            //strSql.AppendLine("  ORDER BY B1.GUBUN1                                               ");
            #endregion

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridGrade.MainView = GridViewGradeSales;
            GridGrade.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void ClearFps()
        {
            GridBuyer.DataSource = null;
            GridGrade.DataSource = null;
        }

        public enum WEIGHT_GUBUN { Purc, Sale, None }
        public WEIGHT_GUBUN _Weight_Gubun = WEIGHT_GUBUN.None;
        private void XtTControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0,10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);
            string sDealGB = RdGbDealGB.EditValue.ToString();

            if (sDealGB == "A")
            {
                sDealGB = "매입";
            }
            else if (sDealGB == "B")
            {
                sDealGB = "매출";
            }
            else
            {
                sDealGB = "직송";
            }

            string sBigCate = string.Empty;

            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.ValueMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0)
            {
                sBigCate = sBigCate.Substring(0, sBigCate.Length - 1);
            }
            else
            {
                sBigCate = null;
            }

            if (XtTControl.SelectedTabPage.Name.Equals("XtTCCompany") && (sDealGB == "매입" || sDealGB == "매출"))
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                #region mariaDB
                //strSql.AppendLine(" SELECT A.J_DATE ");
                //strSql.AppendLine("	     , D.DEALER_NM AS DEALER_NM");
                //strSql.AppendLine("	     , E.EMP_NM AS CHRG_NM ");
                //strSql.AppendLine("	     , F.GUBUN1 AS GRADE ");

                //if (sDealGB == "매입")
                //{
                //    strSql.AppendLine("	     , SUM(A.IWEIGHT) AS ACCEPT_WEIGHT ");
                //    strSql.AppendLine("	     , SUM(A.IKONGKEP) AS PRICE ");
                //}
                //else if (sDealGB == "매출")
                //{
                //    strSql.AppendLine("	     , SUM(A.OWEIGHT) AS ACCEPT_WEIGHT ");
                //    strSql.AppendLine("      , SUM(A.KONGKEP) AS PRICE ");
                //}

                //strSql.AppendLine("	     , AVG(A.DANGA) AS AVG_UNIT_PRICE ");
                //strSql.AppendLine("      , D.J_REGION ");
                //strSql.AppendLine("   FROM INLIST A ");
                //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
                //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
                //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
                //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
                //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS E ");
                //strSql.AppendLine("     ON E.EMP_ID = D.CHRG_ID ");
                //strSql.AppendLine("   LEFT OUTER JOIN JAJAE F ");
                //strSql.AppendLine("     ON F.GUBUN1 = C.GUBUN1 ");
                //strSql.AppendLine("   LEFT OUTER JOIN MESURING G ");
                //if (sDealGB == "매입")
                //{
                //    strSql.AppendLine("     ON G.IPCHULGO_MAIPID = A.J_ID ");
                //}
                //else if (sDealGB == "매출")
                //{
                //    strSql.AppendLine("     ON G.IPCHULGO_MACHULID = A.J_ID ");
                //}
                //strSql.AppendLine("  WHERE G.AGREE_DATE >= '" + sYmdFrom + "' ");
                //strSql.AppendLine("    AND G.AGREE_DATE <= '" + sYmdTo + "' ");
                //strSql.AppendLine("    AND A.KERATYPE = '" + sDealGB + "' ");
                //strSql.AppendLine("    AND G.KERATYPE <> '직송' ");
                //strSql.AppendLine("    AND A.DANGA != 0    ");
                //if (sBigCate != null) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ") ");
                //strSql.AppendLine("  GROUP BY A.J_ID1 ");
                //strSql.AppendLine("  ORDER BY REPLACE(D.DEALER_NM, '(주)', '') ");
                ////strSql.AppendLine("  GROUP BY DEALER_NM, GRADE ");
                ////strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");
                #endregion

                if (sDealGB == "매입")
                {
                    strSql.AppendLine(" SELECT C1.DEALER_NM                                   ");
                    strSql.AppendLine("      , C2.EMP_NM AS CHRG_NM                           ");
                    strSql.AppendLine("      , SUM(A1.DANJUNG) AS ACCEPT_WEIGHT               ");
                    strSql.AppendLine("      , SUM(A1.IKONGKEP) AS PRICE                      ");
                    strSql.AppendLine("      , ROUND(AVG(A1.DANGA), 1, 1) AS AVG_UNIT_PRICE   ");
                    strSql.AppendLine("      , C1.J_REGION                                    ");
                    strSql.AppendLine("   FROM INLIST A1                                      ");
                    strSql.AppendLine("   LEFT JOIN MESURING A2                               ");
                    strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MAIPID                   ");
                    strSql.AppendLine("   LEFT JOIN JAJAE B1                                  ");
                    strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                      ");
                    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C1                          ");
                    strSql.AppendLine("     ON A1.J_ID1 = C1.DEALER_CD                        ");
                    strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C2                           ");
                    strSql.AppendLine("     ON C1.CHRG_ID = C2.EMP_ID                         ");
                    strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'");
                    strSql.AppendLine("    AND A2.KERATYPE <> '직송'                          ");
                    strSql.AppendLine("    AND A1.KERATYPE = '매입'                           ");
                    strSql.AppendLine("    AND A2.J_CHECK = '1'                               ");
                    strSql.AppendLine("    AND A1.DANGA != 0                                  ");
                    if (sBigCate != null) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
                    strSql.AppendLine("  GROUP BY C1.DEALER_NM, C2.EMP_NM, C1.J_REGION        ");
                    strSql.AppendLine("  ORDER BY REPLACE(C1.DEALER_NM, '(주)', '')           ");
                }
                else if (sDealGB == "매출")
                {
                    strSql.AppendLine(" SELECT C1.DEALER_NM                                   ");
                    strSql.AppendLine("      , C2.EMP_NM AS CHRG_NM                           ");
                    strSql.AppendLine("      , SUM(A1.DANJUNG) AS ACCEPT_WEIGHT               ");
                    strSql.AppendLine("      , SUM(A1.KONGKEP) AS PRICE                      ");
                    strSql.AppendLine("      , ROUND(AVG(A1.DANGA), 1, 1) AS AVG_UNIT_PRICE   ");
                    strSql.AppendLine("      , C1.J_REGION                                    ");
                    strSql.AppendLine("   FROM INLIST A1                                      ");
                    strSql.AppendLine("   LEFT JOIN MESURING A2                               ");
                    strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MACHULID                 ");
                    strSql.AppendLine("   LEFT JOIN JAJAE B1                                  ");
                    strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                      ");
                    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C1                          ");
                    strSql.AppendLine("     ON A1.J_ID1 = C1.DEALER_CD                        ");
                    strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C2                           ");
                    strSql.AppendLine("     ON C1.CHRG_ID = C2.EMP_ID                         ");
                    strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'");
                    strSql.AppendLine("    AND A2.KERATYPE <> '직송'                          ");
                    strSql.AppendLine("    AND A1.KERATYPE = '매출'                           ");
                    strSql.AppendLine("    AND A2.J_CHECK = '1'                               ");
                    strSql.AppendLine("    AND A1.DANGA != 0                                  ");
                    if (sBigCate != null) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
                    strSql.AppendLine("  GROUP BY C1.DEALER_NM, C2.EMP_NM, C1.J_REGION        ");
                    strSql.AppendLine("  ORDER BY REPLACE(C1.DEALER_NM, '(주)', '')           ");
                }

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                GridCompany.DataSource = dt;

            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTCCompany") && sDealGB == "직송")
            {
                IN010F01 frm = new IN010F01();
                frm._ParentForm = this;
                if (frm.ShowDialog() == DialogResult.OK && _Weight_Gubun != WEIGHT_GUBUN.None)
                {
                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    if (_Weight_Gubun == WEIGHT_GUBUN.Purc)
                    {
                        strSql.AppendLine(" SELECT C1.DEALER_NM                                    ");
                        strSql.AppendLine("      , C2.EMP_NM AS CHRG_NM                            ");
                        strSql.AppendLine("      , SUM(A1.DANJUNG) AS ACCEPT_WEIGHT                ");
                        strSql.AppendLine("      , SUM(A1.IKONGKEP) AS PRICE                       ");
                        strSql.AppendLine("      , ROUND(AVG(A1.DANGA), 1, 1) AS AVG_UNIT_PRICE    ");
                        strSql.AppendLine("      , C1.J_REGION                                     ");
                        strSql.AppendLine("   FROM INLIST A1                                       ");
                        strSql.AppendLine("   LEFT JOIN MESURING A2                                ");
                        strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MAIPID                    ");
                        strSql.AppendLine("   LEFT JOIN JAJAE B1                                   ");
                        strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                       ");
                        strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C1                           ");
                        strSql.AppendLine("     ON A1.J_ID1 = C1.DEALER_CD                         ");
                        strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C2                            ");
                        strSql.AppendLine("     ON C1.CHRG_ID = C2.EMP_ID                          ");
                        strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "' ");
                        strSql.AppendLine("    AND A2.KERATYPE = '직송'                            ");
                        strSql.AppendLine("    AND A1.KERATYPE = '매입'                            ");
                        strSql.AppendLine("    AND A2.J_CHECK = '1'                                ");
                        strSql.AppendLine("    AND A1.DANGA != 0                                   ");
                        if (sBigCate != null) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
                        strSql.AppendLine("  GROUP BY C1.DEALER_NM, C2.EMP_NM, C1.J_REGION         ");
                        strSql.AppendLine("  ORDER BY REPLACE(C1.DEALER_NM, '(주)', '')            ");
                    }
                    else if (_Weight_Gubun == WEIGHT_GUBUN.Sale)
                    {
                        strSql.AppendLine(" SELECT C1.DEALER_NM                                   ");
                        strSql.AppendLine("      , C2.EMP_NM AS CHRG_NM                           ");
                        strSql.AppendLine("      , SUM(A1.DANJUNG) AS ACCEPT_WEIGHT               ");
                        strSql.AppendLine("      , SUM(A1.KONGKEP) AS PRICE                       ");
                        strSql.AppendLine("      , ROUND(AVG(A1.DANGA), 1, 1) AS AVG_UNIT_PRICE   ");
                        strSql.AppendLine("      , C1.J_REGION                                    ");
                        strSql.AppendLine("   FROM INLIST A1                                      ");
                        strSql.AppendLine("   LEFT JOIN MESURING A2                               ");
                        strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MACHULID                 ");
                        strSql.AppendLine("   LEFT JOIN JAJAE B1                                  ");
                        strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                      ");
                        strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C1                          ");
                        strSql.AppendLine("     ON A1.J_ID1 = C1.DEALER_CD                        ");
                        strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C2                           ");
                        strSql.AppendLine("     ON C1.CHRG_ID = C2.EMP_ID                         ");
                        strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '"+ sYmdFrom + "' AND '"+ sYmdTo + "'");
                        strSql.AppendLine("    AND A2.KERATYPE = '직송'                           ");
                        strSql.AppendLine("    AND A1.KERATYPE = '매출'                           ");
                        strSql.AppendLine("    AND A2.J_CHECK = '1'                               ");
                        strSql.AppendLine("    AND A1.DANGA != 0                                  ");
                        if (sBigCate != null) strSql.AppendLine("    AND B1.DAEGUBUN IN (" + sBigCate + ") ");
                        strSql.AppendLine("  GROUP BY C1.DEALER_NM, C2.EMP_NM, C1.J_REGION        ");
                        strSql.AppendLine("  ORDER BY REPLACE(C1.DEALER_NM, '(주)', '')           ");
                    }

                    #region mariaDB
                    //strSql.Clear();
                    //if (_Weight_Gubun == WEIGHT_GUBUN.Purc)
                    //{
                    //    strSql.AppendLine(" SELECT C1.DEALER_NM                                   ");
                    //    strSql.AppendLine("      , C2.EMP_NM AS CHRG_NM                           ");
                    //    strSql.AppendLine("      , B1.GUBUN1 AS GRADE                             ");
                    //    strSql.AppendLine("      , SUM(A1.DANJUNG) AS ACCEPT_WEIGHT               ");
                    //    strSql.AppendLine("      , SUM(A1.IKONGKEP) AS PRICE                      ");
                    //    strSql.AppendLine("      , TRUNCATE(AVG(A1.DANGA), 1) AS AVG_UNIT_PRICE   ");
                    //    strSql.AppendLine("      , C1.J_REGION                                    ");
                    //    strSql.AppendLine("   FROM INLIST A1                                      ");
                    //    strSql.AppendLine("   LEFT JOIN MESURING A2                               ");
                    //    strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MAIPID                   ");
                    //    strSql.AppendLine("   LEFT JOIN JAJAE B1                                  ");
                    //    strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                      ");
                    //    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C1                          ");
                    //    strSql.AppendLine("     ON A1.J_ID1 = C1.DEALER_CD                        ");
                    //    strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C2                           ");
                    //    strSql.AppendLine("     ON C1.CHRG_ID = C2.EMP_ID                         ");
                    //    strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
                    //    strSql.AppendLine("    AND A2.KERATYPE = '직송'                           ");
                    //    strSql.AppendLine("    AND A1.KERATYPE = '매입'                           ");
                    //    strSql.AppendLine("    AND A2.J_CHECK = '1'                               ");
                    //    strSql.AppendLine("    AND A1.DANGA != 0                                  ");
                    //    strSql.AppendLine("  GROUP BY A1.J_ID1                                    ");
                    //    strSql.AppendLine("  ORDER BY REPLACE(C1.DEALER_NM, '(주)', '')           ");
                    //}
                    //else if (_Weight_Gubun == WEIGHT_GUBUN.Sale)
                    //{
                    //    strSql.AppendLine(" SELECT C1.DEALER_NM                                ");
                    //    strSql.AppendLine("      , C2.EMP_NM AS CHRG_NM                        ");
                    //    strSql.AppendLine("      , B1.GUBUN1 AS GRADE                          ");
                    //    strSql.AppendLine("      , SUM(A1.DANJUNG) AS ACCEPT_WEIGHT            ");
                    //    strSql.AppendLine("      , SUM(A1.KONGKEP) AS PRICE                    ");
                    //    strSql.AppendLine("      , TRUNCATE(AVG(A1.DANGA), 1) AS AVG_UNIT_PRICE");
                    //    strSql.AppendLine("      , C1.J_REGION                                 ");
                    //    strSql.AppendLine("   FROM INLIST A1                                   ");
                    //    strSql.AppendLine("   LEFT JOIN MESURING A2                            ");
                    //    strSql.AppendLine("     ON A1.J_ID = A2.IPCHULGO_MACHULID              ");
                    //    strSql.AppendLine("   LEFT JOIN JAJAE B1                               ");
                    //    strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                   ");
                    //    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C1                       ");
                    //    strSql.AppendLine("     ON A1.J_ID1 = C1.DEALER_CD                     ");
                    //    strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C2                        ");
                    //    strSql.AppendLine("     ON C1.CHRG_ID = C2.EMP_ID                      ");
                    //    strSql.AppendLine("  WHERE A1.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
                    //    strSql.AppendLine("    AND A2.KERATYPE = '직송'       ");
                    //    strSql.AppendLine("    AND A1.KERATYPE = '매출'                        ");
                    //    strSql.AppendLine("    AND A2.J_CHECK = '1'                            ");
                    //    strSql.AppendLine("    AND A1.DANGA != 0                               ");
                    //    strSql.AppendLine("  GROUP BY A1.J_ID1                                 ");
                    //    strSql.AppendLine("  ORDER BY REPLACE(C1.DEALER_NM, '(주)', '')        ");
                    //}
                    #endregion

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    GridCompany.DataSource = dt;
                }

                _Weight_Gubun = WEIGHT_GUBUN.None;
            }
            Cursor = Cursors.Default;
        }

        private void RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {

        }

        private void SetGridLookupEdit(DevExpress.XtraEditors.GridLookUpEdit gLkup, string sNullYn, string sGb)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '전체' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }
            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT DAEGUBUN AS CD ");
                strSql.AppendLine("      , DAEGUBUN AS NM ");
                strSql.AppendLine("   FROM JAJAE ");
                strSql.AppendLine("  WHERE DAEGUBUN IN ('고철A', '고철B', '슈레더') ");
                strSql.AppendLine("  GROUP BY DAEGUBUN  ");

                //strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                //strSql.AppendLine("      , A.COM_NM AS NM ");
                //strSql.AppendLine("   FROM COM_BASE_CD A ");
                //strSql.AppendLine("  WHERE A.CD_GB = '0002' ");
                //strSql.AppendLine("    AND A.COM_CD NOT IN('E', 'F')");
            }
            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gLkup.Properties.DataSource = dt;
            gLkup.Properties.DisplayMember = "NM";
            gLkup.Properties.ValueMember = "CD";
        }

        private void PerBuyerRowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;

            if (GridBuyer.MainView == GridViewBuyerPurc)
            {
                if (e.Column.FieldName == "PAYEDUNITPRICE")
                {
                    string sStdd = view.GetRowCellDisplayText(e.RowHandle, view.Columns["STDDUNITPRICE"]).ToString();
                    string dPay = view.GetRowCellDisplayText(e.RowHandle, view.Columns["PAYEDUNITPRICE"]).ToString();

                    double dStddUtPrc = String.IsNullOrEmpty(sStdd) ? 0 : Convert.ToDouble(sStdd);
                    double dPayUtPrc = String.IsNullOrEmpty(dPay) ? 0 : Convert.ToDouble(dPay);

                    if (dStddUtPrc > dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.RoyalBlue;
                    }
                    else if (dStddUtPrc < dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.IndianRed;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            if (GridBuyer.MainView == GridViewBuyerSales)
            {
                if (e.Column.FieldName == "SALEUNITPRICE")
                {
                    string sStdd = view.GetRowCellDisplayText(e.RowHandle, view.Columns["STDDUNITPRICE"]).ToString();
                    string sSale = view.GetRowCellDisplayText(e.RowHandle, view.Columns["SALEUNITPRICE"]).ToString();

                    double dStddUtPrc = String.IsNullOrEmpty(sStdd) ? 0 : Convert.ToDouble(sStdd);
                    double dPayUtPrc = String.IsNullOrEmpty(sSale) ? 0 : Convert.ToDouble(sSale);

                    if (dStddUtPrc > dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.RoyalBlue;
                    }
                    else if (dStddUtPrc < dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.IndianRed;
                    }
                    else
                    {
                        return;
                    }
                }
                if (e.Column.FieldName == "DIFFWEIGHT")
                {
                    string sDiff = view.GetRowCellDisplayText(e.RowHandle, view.Columns["DIFFWEIGHT"]).ToString();

                    double dDiffUtPrc = String.IsNullOrEmpty(sDiff) ? 0 : Convert.ToDouble(sDiff);

                    if (dDiffUtPrc > 0) e.Appearance.BackColor = Color.IndianRed;
                }
            }

            if (GridGrade.MainView == GridViewGradePurc || GridGrade.MainView == GridViewGradeSales)
            {
                if (e.Column.FieldName == "DIFFWEIGHT")
                {
                    string sDiff = view.GetRowCellDisplayText(e.RowHandle, view.Columns["DIFFWEIGHT"]).ToString();
                    double dDiffWeight = String.IsNullOrEmpty(sDiff) ? 0 : Convert.ToDouble(sDiff);

                    if (dDiffWeight != 0)
                    {
                        e.Appearance.BackColor = Color.IndianRed;
                    }
                }
            }
        }
        private void PerGradeRowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;

            if (e.Column.FieldName == "DIFFWEIGHT")
            {
                string sDiff = view.GetRowCellDisplayText(e.RowHandle, view.Columns["DIFFWEIGHT"]).ToString();
                double dDiffWeight = String.IsNullOrEmpty(sDiff) ? 0 : Convert.ToDouble(sDiff);

                if (dDiffWeight != 0)
                {
                    e.Appearance.BackColor = Color.IndianRed;
                }
            }
        }


        private void GLkupEditCate_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            string sBigCate = string.Empty;

            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.DisplayMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0) e.DisplayText = sBigCate.Substring(0, sBigCate.Length - 1);
        }

        private void GLkupEditCate_CloseUp(object sender, DevExpress.XtraEditors.Controls.CloseUpEventArgs e)
        {
            BtnRetr.Focus();
        }

        private void DateEditFrom_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void DateEditTo_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                string sDealType = RdGbDealGB.EditValue?.ToString();

                if (sDealType.Equals("A"))
                {
                    sDealType = "매입";
                }
                else if (sDealType.Equals("B"))
                {
                    sDealType = "매출";
                }
                else if (sDealType.Equals("C"))
                {
                    sDealType = "직송";
                }

                //string sFileNM = "현장매입출일보 " + sYmdFrom + "~" + sYmdTo + " " + sDealType;
                string sFileNM = this.Text;
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    if (XtTControl.SelectedTabPage.Name.ToString().Equals("XtTCRetr1"))
                    {
                        XlsxExportOptions options = new XlsxExportOptions();
                        options.ExportMode = XlsxExportMode.SingleFilePageByPage;
                        //CompositeLink link = new CompositeLink();

                        GridBuyer.ExportToXlsx(FileName, options);
                        Process.Start(FileName);
                    }
                    else if (XtTControl.SelectedTabPage.Name.ToString().Equals("XtTCRetr2"))
                    {
                        GridGrade.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else
                    {
                        GridCompany.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                }
                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        private void SaveMultiGridViewToExcel(GridView[] ArrGridViews)
        {
            //int num = 0;
            
            //object missingType = Type.Missing;
            //Excel.Application objApp;
            //Excel._Workbook objBook;
            //Excel.Workbooks objBooks;
            //Excel.Sheets objSheets;
            //Excel._Worksheet objSheet;
            //Excel.Range range;
            
            //string[] headers = new string[myDataGridViews[0].ColumnCount];
            //string[] columns = new string[myDataGridViews[0].ColumnCount];
            //for (int c = 0; c < myDataGridViews[0].ColumnCount; c++)
            //{
            //    headers[c] = myDataGridViews[0].Rows[0].Cells[c].OwningColumn.HeaderText.ToString();
            //    if (c <= 25)
            //    {
            //        num = c + 65;
            //        columns[c] = Convert.ToString((char)num);
            //    }
            //    else
            //    {
            //        columns[c] = Convert.ToString((char)(Convert.ToInt32(c / 26) - 1 + 65)) + Convert.ToString((char)(c % 26 + 65));
            //    }
            //}

            //try
            //{
            //    objApp = new Excel.Application();
            //    objBooks = objApp.Workbooks;
            //    objBook = objBooks.Add(Missing.Value);
            //    objSheets = objBook.Worksheets;
            //    objSheet = (Excel._Worksheet)objSheets.get_Item(1);
                
            //    if (captions)
            //    {
            //        for (int c = 0; c < myDataGridViews[0].ColumnCount; c++)
            //        {
            //            range = objSheet.get_Range(columns[c] + "1", Missing.Value);
            //            range.set_Value(Missing.Value, headers[c]);
            //        }
            //    }
                
            //    int iGridViewNum = 0;
            //    foreach (DataGridView myDataGridView in myDataGridViews)
            //    {
            //        columns = new string[myDataGridView.ColumnCount];
            //        for (int i = 0; i < myDataGridView.RowCount; i++)
            //        {
            //            for (int j = 0; j < myDataGridView.ColumnCount; j++)
            //            {
            //                if (j <= 25)
            //                {
            //                    num = j + 65;
            //                    columns[j] = Convert.ToString((char)num);
            //                }
            //                else
            //                {
            //                    columns[j] = Convert.ToString((char)(Convert.ToInt32(j / 26) - 1 + 65)) + Convert.ToString((char)(j % 26 + 65));
            //                }

            //                range = objSheet.get_Range(columns[j] + Convert.ToString(iGridViewNum + i + 2), Missing.Value);
            //                range.set_Value(Missing.Value, myDataGridView.Rows[i].Cells[j].Value.ToString());
            //            }
            //        }
            //        iGridViewNum++;
            //    }

            //    objApp.Visible = false;
            //    objApp.UserControl = false;
            //    objBook.SaveAs(@saveFileDialog.FileName,
            //        Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
            //        missingType, missingType, missingType, missingType,
            //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            //        missingType, missingType, missingType, missingType, missingType);

            //    objBook.Close(false, missingType, missingType);
            //    Cursor.Current = Cursors.Default;

            //    MessageBox.Show("Save Success!!!");
            //}
            //catch (Exception theException)
            //{
            //    String errorMessage;
            //    errorMessage = "Error: ";
            //    errorMessage = String.Concat(errorMessage, theException.Message);
            //    errorMessage = String.Concat(errorMessage, " Line: ");
            //    errorMessage = String.Concat(errorMessage, theException.Source);
            //    MessageBox.Show(errorMessage, "Error");
            //}
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AccFieldPSDailyRpt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {

            }
            else if (e.KeyCode == Keys.F1)
            {

            }
            else if (e.KeyCode == Keys.F3)
            {

            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExport_Click(null, null);
            }
        }

        private void GridBuyer_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridGrade_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void XtTControl_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewBuyerPurc_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccFieldPSDailyRpt_FormClosed(object sender, FormClosedEventArgs e)
        {

        }


        #region [정렬기능(2020-06-02 정은영)]
        private ColumnSortOrder GetNextSortOrder(ColumnSortOrder order)
        {
            switch (order)
            {
                case ColumnSortOrder.None: return ColumnSortOrder.Descending;
                case ColumnSortOrder.Descending: return ColumnSortOrder.Ascending;
                case ColumnSortOrder.Ascending: return ColumnSortOrder.None;
            }

            return ColumnSortOrder.None;
        }

        private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        {

            if (sender.GetType() == GridViewBuyerPurc.GetType())
            {
                BandedGridView view = (BandedGridView)sender;
                BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

                if (hitInfo.InBandPanel)
                {
                    if (hitInfo.Band.Name.Equals("gridBand11") || hitInfo.Band.Name.Equals("gridBand30")) //상호
                    {
                        ColumnSortOrder order = view.Columns["DEALER_NM"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "상호↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "상호↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "상호";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand12") || hitInfo.Band.Name.Equals("gridBand31") || hitInfo.Band.Name.Equals("gridBand49")) //담당자
                    {
                        ColumnSortOrder order = view.Columns["INSPECTOR"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "담당자↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "담당자↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "담당자";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand18") || hitInfo.Band.Name.Equals("gridBand37") || hitInfo.Band.Name.Equals("gridBand50")) //등급
                    {
                        ColumnSortOrder order = view.Columns["GUBUN1"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "등급↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "등급↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "등급";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand52")) //매출처
                    {
                        ColumnSortOrder order = view.Columns["DEALER_NM"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "매출처↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "매출처↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "매출처";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand56")) //매입처
                    {
                        ColumnSortOrder order = view.Columns["J_BOOKING"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "매입처↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "매입처↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "매입처";
                    }
                }

            }
            else
            {
                GridView view = (GridView)sender;
                GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

                if (hitInfo.InColumn)
                {
                    ColumnSortOrder order = hitInfo.Column.SortOrder;

                    order = GetNextSortOrder(order);

                    hitInfo.Column.SortOrder = order;
                    view.FocusedRowHandle = 0;
                }
            }
        }
        #endregion

        private void AccFieldPSDailyRpt_TextChanged(object sender, EventArgs e)
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

        private void GridViewBuyerPurc_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewBuyerPurc_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColPurcPayUtPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColPurcPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColPurcArrvPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
        }

        private void GridViewBuyerSales_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColSaleUtPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColSalePrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColSaleArrvPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
        }

        private void GridViewBuyerPurc_RowClick(object sender, RowClickEventArgs e)
        {

        }

        private void GridViewBuyerPurc_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                if(e.Column == GridColPurcImg || e.Column == GridColSaleImg)
                {
                    try
                    {
                        Cursor = Cursors.WaitCursor;

                        GridView view = (GridView)sender;
                        AccFieldPSRetrImage frm = new AccFieldPSRetrImage();
                        frm._JunpyoID = view.GetFocusedRowCellValue(e.Column)?.ToString();
                        frm._JDate = view.GetFocusedRowCellValue("MESURE_DT")?.ToString();
                        frm.ShowDialog();

                        Cursor = Cursors.Default;
                    }
                    catch (Exception ex)
                    {

                    }
                }

                if(e.Column == GridColEtcCost)
                {
                    string cost = GridViewBuyerPurc.GetFocusedRowCellValue(GridColEtcCost)?.ToString();

                    double.TryParse(cost, out double result);

                    if(result != 0)
                    {
                        PopEtcCostDtl frm = new PopEtcCostDtl();

                        frm.Owner = this;
                        frm.rowInfo = GridViewBuyerPurc.GetDataRow(e.RowHandle);
                        frm.ShowDialog();
                    }
                }
            }
        }

        private void GLkupEditCate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewBuyerPurc_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e)
        {
            GridColumnSummaryItem item = e.Item as GridColumnSummaryItem;
            if(item.Tag.ToString() == "F01")
            {
                double kk = (Convert.ToDouble(GridColPurcArrvPrc.SummaryItem.SummaryValue) + Convert.ToDouble(bandedGridColumn3.SummaryItem.SummaryValue) + Convert.ToDouble(bandedGridColumn6.SummaryItem.SummaryValue))
                          / Convert.ToDouble(GridColPurcInWeight.SummaryItem.SummaryValue);
                string yy = string.Format("{0:0.#}", kk);
                e.TotalValue = yy;
            }
        }

        private void GridViewGradePurc_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e)
        {
            GridColumnSummaryItem item = e.Item as GridColumnSummaryItem;
            if (item.Tag.ToString() == "F02")
            {
                double kk = Convert.ToDouble(GridColGPurcPrice.SummaryItem.SummaryValue)
                          / Convert.ToDouble(GridColGPurcAcctAmt.SummaryItem.SummaryValue);
                string yy = string.Format("{0:0.#}", kk);
                e.TotalValue = yy;
            }
        }

        private void GridViewCompany_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e)
        {
            GridColumnSummaryItem item = e.Item as GridColumnSummaryItem;
            if (item.Tag.ToString() == "F03")
            {
                double kk = Convert.ToDouble(GridColCompPrice.SummaryItem.SummaryValue)
                          / Convert.ToDouble(GridColCompAcptWeight.SummaryItem.SummaryValue);
                string yy = string.Format("{0:0.#}", kk);
                e.TotalValue = yy;
            }
        }
    }
}