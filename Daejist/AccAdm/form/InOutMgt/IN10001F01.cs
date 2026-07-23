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
using ComLib;
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Columns;
using System.Data.SqlClient;
/*
* 작성일자 : 2021-03-15
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
*/
namespace AccAdm
{
    public partial class IN10001F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN10001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void IN10001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewBuyerPurc, GridViewBuyerSales };

            DataTable dtEmp = SetLookupEdit();
            LkupChrgId.Properties.DataSource = dtEmp;
            LkupChrgId.Properties.ValueMember = "CD";
            LkupChrgId.Properties.DisplayMember = "NM";
            LkupChrgId.EditValue = string.Empty;

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

            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                /*
                 * 0 : 마감일자
                 * 1 : 계근일자
                 */
                int iDateGb = CboDateGb.SelectedIndex;
                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                /*
                 * 0 : 매입
                 * 1 : 매출
                 */
                int iKeraType = CboKeraType.SelectedIndex;
                string sChrgId = LkupChrgId.EditValue?.ToString();

                /*
                 * 0 : 거래처명
                 * 1 : 등급
                 * 2 : 차번
                 */
                string sFindIdx = CboFindIdx.SelectedIndex.ToString();
                string sFindWord = TxtFindWord.EditValue?.ToString().Trim();

                if (iDateGb < 0)
                {
                    XtraMessageBox.Show("마감일자/계근일자 중 하나를 선택하세요.");
                    CboDateGb.Focus();
                    return;
                }
                else if (string.IsNullOrEmpty(sYmdFrom))
                {
                    XtraMessageBox.Show("일자를 올바르게 입력하세요.");
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                    return;
                }
                else if (string.IsNullOrEmpty(sYmdTo))
                {
                    XtraMessageBox.Show("일자를 올바르게 입력하세요.");
                    DateEditTo.SelectAll();
                    DateEditTo.Focus();
                    return;
                }
                else if (iKeraType < 0)
                {
                    XtraMessageBox.Show("매입/매출 중 하나를 선택하세요.");
                    CboKeraType.Focus();
                    return;
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("DATE_IDX", iDateGb.ToString());
                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);
                dicParams.Add("KERATYPE", iKeraType == 0 ? "매입" : "매출");
                dicParams.Add("CHRG_ID", sChrgId);
                dicParams.Add("FIND_IDX", sFindIdx);
                dicParams.Add("FIND_WORD", sFindWord);

                DataTable dt = null;
                if (iKeraType == 0)
                {
                    GridRetr.MainView = GridViewBuyerPurc;
                    dt = PurchacePerBuyer(dicParams);
                }
                else if (iKeraType == 1)
                {
                    GridRetr.MainView = GridViewBuyerSales;
                    dt = SalesPerBuyer(dicParams);
                }
                GridRetr.DataSource = null;
                GridRetr.DataSource = dt;

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("[Retr Error] " + ex.Message);
            }
            
        }

        private DataTable PurchacePerBuyer(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            #region mariaDB
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , '' AS REMARK ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , IFNULL(YY.CNT, 0) AS CNT ");
            //strSql.AppendLine("               , A.TAXNO ");
            //strSql.AppendLine("               , A.KERATYPE ");
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
            //strSql.AppendLine("	              , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("	              , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("	              , A.IKONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("	              , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("	              , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("	              , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("	              , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("               , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("               , G.ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , G.ETC_DEALER_NM2 ");
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
            //strSql.AppendLine("            LEFT JOIN ( SELECT COUNT(*) AS CNT ");
            //strSql.AppendLine("                             , X1.REF1 AS REF1 ");
            //strSql.AppendLine("                          FROM ACTRAN X1 ");
            //strSql.AppendLine("                         GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO ) YY ");
            //strSql.AppendLine("              ON A.JUNPYOID = YY.REF1 ");
            //strSql.AppendLine("           WHERE (('" + dicParams["DATE_IDX"] + "' = '0' AND (A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["DATE_IDX"] + "' = '1' AND (G.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')))");
            //strSql.AppendLine("             AND A.KERATYPE = '" + dicParams["KERATYPE"] + "' ");
            //strSql.AppendLine("             AND (('" + dicParams["CHRG_ID"] + "' = '' AND 1 = 1) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["CHRG_ID"] + "' <> '' AND D.CHRG_ID LIKE '%" + dicParams["CHRG_ID"] + "%')) ");
            //strSql.AppendLine("             AND (('" +dicParams["FIND_WORD"]+ "' = '' AND 1 = 1) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '0' AND D.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND C.GUBUN1 LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND G.J_BNUM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND A.DANGA > 0 ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY CASE WHEN '" + dicParams["DATE_IDX"] + "' = '0' THEN Y1.J_DATE ELSE Y1.MESURE_DT END, Y1.DEALER_NM ");
            #endregion

            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            strSql.AppendLine("      , '' AS REMARK ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT A.J_DATE ");
            strSql.AppendLine("               , ISNULL(YY.CNT, 0) AS CNT ");
            strSql.AppendLine("               , A.TAXNO ");
            strSql.AppendLine("               , A.KERATYPE ");
            strSql.AppendLine("               , A.JUNPYOID ");
            strSql.AppendLine("               , G.J_DATE AS MESURE_DT ");
            strSql.AppendLine("	              , D.DEALER_NM ");
            strSql.AppendLine("	              , E.EMP_NM AS INSPECTOR ");
            strSql.AppendLine("	              , B.J_BNUM ");
            strSql.AppendLine("	              , G.WEIGHT AS LANDEDWEIGHT ");
            strSql.AppendLine("	              , A.HALIN AS LOSS ");
            strSql.AppendLine("	              , A.IWEIGHT AS ACCEPTWEIGHT ");
            strSql.AppendLine("	              , ISNULL(A.IWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            strSql.AppendLine("	              , F.GUBUN1 ");
            strSql.AppendLine("	              , A.MIDANGA AS STDDUNITPRICE ");
            strSql.AppendLine("	              , A.DANGA AS PAYEDUNITPRICE ");
            strSql.AppendLine("	              , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            strSql.AppendLine("	              , ISNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            strSql.AppendLine("	              , A.IKONGKEP AS TOTALPRICE ");
            strSql.AppendLine("	              , (ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("	              , G.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("	              , G.JUNPYOID AS IMAGE ");
            strSql.AppendLine("	              , G.GUMSUBIGO AS INSPECTNOTE ");
            strSql.AppendLine("               , FLOOR((ISNULL(A.CKONGKEP, 0) / A.DANJUNG)) AS CARRY_UNIT_PRICE ");
            strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("               , G.ETC_COST1 ");
            strSql.AppendLine("               , G.ETC_REMARK1 ");
            strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("               , G.ETC_COST2 ");
            strSql.AppendLine("               , G.ETC_REMARK2 ");
            strSql.AppendLine("            FROM INLIST A ");
            strSql.AppendLine("            LEFT OUTER JOIN IPCHULGO B ");
            strSql.AppendLine("	             ON B.J_ID = A.J_ID ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE C ");
            strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D ");
            strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E ");
            strSql.AppendLine("              ON E.EMP_ID = A.CHRG_ID ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ( SELECT COUNT(*) AS CNT ");
            strSql.AppendLine("                             , X1.REF1 AS REF1 ");
            strSql.AppendLine("                          FROM ACTRAN X1 ");
            strSql.AppendLine("                         GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO, X1.REF1 ) YY ");
            strSql.AppendLine("              ON CONVERT(VARCHAR,A.JUNPYOID) = YY.REF1 ");
            strSql.AppendLine("           WHERE (('" + dicParams["DATE_IDX"] + "' = '0' AND (A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')) ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["DATE_IDX"] + "' = '1' AND (G.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')))");
            strSql.AppendLine("             AND A.KERATYPE = '" + dicParams["KERATYPE"] + "' ");
            strSql.AppendLine("             AND (('" + dicParams["CHRG_ID"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["CHRG_ID"] + "' <> '' AND D.CHRG_ID LIKE '%" + dicParams["CHRG_ID"] + "%')) ");
            strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '0' AND D.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND C.GUBUN1 LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND G.J_BNUM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            strSql.AppendLine("             AND A.DANGA > 0 ");
            strSql.AppendLine("             AND G.J_CHECK = '1' ");
            strSql.AppendLine("      ) Y1 ");
            strSql.AppendLine("  ORDER BY CASE WHEN '" + dicParams["DATE_IDX"] + "' = '0' THEN Y1.J_DATE ELSE Y1.MESURE_DT END, Y1.DEALER_NM ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable SalesPerBuyer(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            /*
             * INLIST CHRG_ID 추가됨에 따라 해당 컬럼으로 업체담당자 JOIN
             */
            #region mariaDB
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , '' AS REMARK ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_DATE ");
            //strSql.AppendLine("               , IFNULL(YY.CNT, 0) AS CNT ");
            //strSql.AppendLine("               , A.TAXNO ");
            //strSql.AppendLine("               , A.KERATYPE ");
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
            //strSql.AppendLine("               , G.ETC_DEALER_NM1 ");
            //strSql.AppendLine("               , G.ETC_COST1 ");
            //strSql.AppendLine("               , G.ETC_REMARK1 ");
            //strSql.AppendLine("               , G.ETC_DEALER_NM2 ");
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
            //strSql.AppendLine("            LEFT JOIN ( SELECT COUNT(*) AS CNT ");
            //strSql.AppendLine("                             , X1.REF1 AS REF1 ");
            //strSql.AppendLine("                          FROM ACTRAN X1 ");
            //strSql.AppendLine("                         GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO ) YY ");
            //strSql.AppendLine("              ON A.JUNPYOID = YY.REF1 ");
            //strSql.AppendLine("           WHERE (('" + dicParams["DATE_IDX"] + "' = '0' AND (A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["DATE_IDX"] + "' = '1' AND (G.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')))");
            //strSql.AppendLine("             AND A.KERATYPE = '" + dicParams["KERATYPE"] + "' ");
            //strSql.AppendLine("             AND (('" + dicParams["CHRG_ID"] + "' = '' AND 1 = 1) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["CHRG_ID"] + "' <> '' AND D.CHRG_ID LIKE '%" + dicParams["CHRG_ID"] + "%')) ");
            //strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '0' AND D.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND C.GUBUN1 LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND G.J_BNUM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            //strSql.AppendLine("             AND A.DANGA > 0 ");
            //strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            //strSql.AppendLine("             AND G.J_CHECK = '1' ");
            //strSql.AppendLine("      ) Y1 ");
            //strSql.AppendLine("  ORDER BY CASE WHEN '" + dicParams["DATE_IDX"] + "' = '0' THEN Y1.J_DATE ELSE Y1.MESURE_DT END, Y1.DEALER_NM ");
            #endregion

            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE ");
            strSql.AppendLine("      , '' AS REMARK ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT A.J_DATE ");
            strSql.AppendLine("               , ISNULL(YY.CNT, 0) AS CNT ");
            strSql.AppendLine("               , A.TAXNO ");
            strSql.AppendLine("               , A.KERATYPE ");
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
            strSql.AppendLine("               , ((ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE ");
            strSql.AppendLine("               , ISNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            strSql.AppendLine("               , A.KONGKEP AS TOTALPRICE ");
            strSql.AppendLine("               , (ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("               , G.JUNPYOID AS IMAGE ");
            strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE ");
            strSql.AppendLine("               , FLOOR((ISNULL(A.CKONGKEP, 0) / A.DANJUNG)) AS CARRY_UNIT_PRICE ");
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
            strSql.AppendLine("              ON E.EMP_ID = A.CHRG_ID ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE F ");
            strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1 ");
            strSql.AppendLine("            LEFT OUTER JOIN MESURING G ");
            strSql.AppendLine("              ON G.IPCHULGO_MACHULID = A.J_ID ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("            LEFT JOIN ( SELECT COUNT(*) AS CNT ");
            strSql.AppendLine("                             , X1.REF1 AS REF1 ");
            strSql.AppendLine("                          FROM ACTRAN X1 ");
            strSql.AppendLine("                         GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO, X1.REF1 ) YY ");
            strSql.AppendLine("              ON CONVERT(VARCHAR,A.JUNPYOID) = YY.REF1 ");
            strSql.AppendLine("           WHERE (('" + dicParams["DATE_IDX"] + "' = '0' AND (A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')) ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["DATE_IDX"] + "' = '1' AND (G.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "')))");
            strSql.AppendLine("             AND A.KERATYPE = '" + dicParams["KERATYPE"] + "' ");
            strSql.AppendLine("             AND (('" + dicParams["CHRG_ID"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["CHRG_ID"] + "' <> '' AND D.CHRG_ID LIKE '%" + dicParams["CHRG_ID"] + "%')) ");
            strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '0' AND D.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND C.GUBUN1 LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                  OR ");
            strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND G.J_BNUM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("             AND A.DANGA > 0 ");
            strSql.AppendLine("             AND B.KERAGUBUN <> '직송' ");
            strSql.AppendLine("             AND G.J_CHECK = '1' ");
            strSql.AppendLine("      ) Y1 ");
            strSql.AppendLine("  ORDER BY CASE WHEN '" + dicParams["DATE_IDX"] + "' = '0' THEN Y1.J_DATE ELSE Y1.MESURE_DT END, Y1.DEALER_NM ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                GridView view = (GridView)GridRetr.MainView;
                GridColumn gridColJunpyo = null;
                if(view == GridViewBuyerPurc)
                {
                    gridColJunpyo = GridColPurcJunpyoId;
                }
                else
                {
                    gridColJunpyo = GridColSaleJunpyoID;
                }

                if (view.RowCount == 0)
                {
                    XtraMessageBox.Show("마감리스트에 데이터가 존재하지 않습니다.");
                    return;
                }
                
                for(int i = 0; i < view.RowCount; i++)
                {
                    double dAmt = 0;
                    double.TryParse(view.GetRowCellValue(i, "TOTALPRICE")?.ToString(), out dAmt);
                    if(dAmt == 0)
                    {
                        XtraMessageBox.Show("금액을 0 이상 입력하세요.");
                        if(GridRetr.MainView == GridViewBuyerPurc)
                        {
                            view.FocusedRowHandle = i;
                            view.FocusedColumn = GridColPurcPrc;
                        }
                        else if (GridRetr.MainView == GridViewBuyerSales)
                        {
                            view.FocusedRowHandle = i;
                            view.FocusedColumn = GridColSalePrc;
                        }
                        return;
                    }
                }

                Cursor = Cursors.WaitCursor;

                DataTable dt = (DataTable)GridRetr.DataSource;
                DataTable dtChanged = dt.GetChanges(DataRowState.Modified);
                if(dtChanged == null)
                {
                    XtraMessageBox.Show("수정한 내역이 존재하지 않습니다.");
                    return;
                }
                StringBuilder strSql = new StringBuilder();

                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sJunpyoId_Find = string.Empty;
                int iCnt = 0;
                foreach (DataRow row in dtChanged.Rows)
                {
                    string sJunpyoId = row["JUNPYOID"]?.ToString();
                    if(iCnt == 0)
                    {
                        sJunpyoId_Find = sJunpyoId;
                    }
                    iCnt++;

                    double dBeforeAmt = 0;
                    double dAfterAmt = 0;

                    double.TryParse(row["TOTALPRICE", DataRowVersion.Original]?.ToString(), out dBeforeAmt);
                    double.TryParse(row["TOTALPRICE", DataRowVersion.Current]?.ToString(), out dAfterAmt);

                    string sRemark = row["REMARK"]?.ToString();

                    string sKeraType = row["KERATYPE"]?.ToString();

                    strSql.Clear();
                    strSql.AppendFormat(" ");
                    strSql.AppendFormat(" UPDATE INLIST ");
                    if (sKeraType.Equals("매입"))
                    {
                        strSql.AppendFormat("      SET IKONGKEP = {0}", dAfterAmt);
                    }
                    else
                    {
                        strSql.AppendFormat("      SET KONGKEP = {0}", dAfterAmt);
                    }
                    strSql.AppendFormat("      , BUGASE = {0}", Math.Round((dAfterAmt * 0.1), 0));
                    strSql.AppendFormat("  WHERE JUNPYOID = {0}", sJunpyoId);

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                    cmd.Parameters.AddWithValue("@CHG_AMT_DT", DateTime.Now.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@PRV_CLO_AMT", dBeforeAmt);
                    cmd.Parameters.AddWithValue("@CUR_CLO_AMT", dAfterAmt);
                    cmd.Parameters.AddWithValue("@CHG_REMARK", sRemark);
                    cmd.Parameters.AddWithValue("@USER", FmMainToolBar2.UserID);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    strSql.Clear();
                    strSql.AppendLine("           INSERT INTO INLIST_LOG                                     ");
                    strSql.AppendLine("                (JUNPYOID                                             ");
                    strSql.AppendLine("                , JUNPYO_SEQ                                          ");
                    strSql.AppendLine("                , CHG_AMT_DT                                          ");
                    strSql.AppendLine("                , PRV_CLO_AMT                                         ");
                    strSql.AppendLine("                , CUR_CLO_AMT                                         ");
                    strSql.AppendLine("                , CHG_REMARK                                          ");
                    strSql.AppendLine("                , CDATE                                               ");
                    strSql.AppendLine("                , CUSER)                                              ");
                    strSql.AppendLine("          VALUES(@JUNPYOID                                            ");
                    strSql.AppendLine("                , (SELECT ISNULL(MAX(JUNPYO_SEQ), 0) + 1 FROM INLIST_LOG X1 WHERE JUNPYOID = @JUNPYOID)");
                    strSql.AppendLine(" 	           , @CHG_AMT_DT                         ");
                    strSql.AppendLine(" 	           , @PRV_CLO_AMT                        ");
                    strSql.AppendLine(" 	           , @CUR_CLO_AMT                        ");
                    strSql.AppendLine(" 	           , @CHG_REMARK                         ");
                    strSql.AppendLine(" 	           , CONVERT(VARCHAR(20), GETDATE(), 20) ");
                    strSql.AppendLine(" 	           , @USER )                             ");

                    #region MariaDB
                    //strSql.AppendLine(" INSERT INTO INLIST_LOG ");
                    //strSql.AppendLine("           ( JUNPYOID ");
                    //strSql.AppendLine("           , JUNPYO_SEQ ");
                    //strSql.AppendLine("           , CHG_AMT_DT ");
                    //strSql.AppendLine("           , PRV_CLO_AMT ");
                    //strSql.AppendLine("           , CUR_CLO_AMT ");
                    //strSql.AppendLine("           , CHG_REMARK ");
                    //strSql.AppendLine("           , CDATE ");
                    //strSql.AppendLine("           , CUSER ) ");
                    //strSql.AppendLine("     VALUES( @JUNPYOID ");
                    //strSql.AppendLine("           , ( SELECT ISNULL(MAX(JUNPYO_SEQ), 0) + 1 FROM INLIST_LOG X1 WHERE JUNPYOID = @JUNPYOID ) ");
                    //strSql.AppendLine("           , @CHG_AMT_DT ");
                    //strSql.AppendLine("           , @PRV_CLO_AMT ");
                    //strSql.AppendLine("           , @CUR_CLO_AMT ");
                    //strSql.AppendLine("           , @CHG_REMARK ");
                    //strSql.AppendLine("           , NOW() ");
                    //strSql.AppendLine("           , @USER ) ");
                    //strSql.AppendLine("     ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("           CHG_AMT_DT = @CHG_AMT_DT ");
                    //strSql.AppendLine("         , PRV_CLO_AMT = @PRV_CLO_AMT ");
                    //strSql.AppendLine("         , CUR_CLO_AMT = @CUR_CLO_AMT ");
                    //strSql.AppendLine("         , CHG_REMARK = @CHG_REMARK ");
                    //strSql.AppendLine("         , MDATE = NOW() ");
                    //strSql.AppendLine("         , MUSER = @USER ");
                    #endregion

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                    cmd.Parameters.AddWithValue("@CHG_AMT_DT", DateTime.Now.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@PRV_CLO_AMT", dBeforeAmt);
                    cmd.Parameters.AddWithValue("@CUR_CLO_AMT", dAfterAmt);
                    cmd.Parameters.AddWithValue("@CHG_REMARK", sRemark);
                    cmd.Parameters.AddWithValue("@USER", FmMainToolBar2.UserID);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    /*
                     * LOG 작업진행
                     */
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT B.J_DATE ");
                    strSql.AppendLine("      , B.SUN ");
                    strSql.AppendLine("      , CASE WHEN B.KERATYPE = '입고' THEN B.MAIPCHER ELSE B.J_COMPANY END AS DEALER_NM ");
                    strSql.AppendLine("      , A.GUBUN1 ");
                    strSql.AppendLine("      , B.J_BNUM ");
                    strSql.AppendLine("   FROM INLIST A ");
                    strSql.AppendLine("   LEFT JOIN MESURING B ");
                    strSql.AppendLine("     ON A.J_ID = CASE WHEN B.KERATYPE = '입고' THEN B.IPCHULGO_MAIPID ELSE B.IPCHULGO_MACHULID END ");
                    strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + "");

                    DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    if(dtPrv.Rows.Count > 0)
                    {
                        string sLogMsg = string.Format("[마감금액조정]계근일자 : {0}, 순번{1}, 업체 : {2}, 등급 : {3}, 차번 {4}"
                        , dtPrv.Rows[0]["J_DATE"]?.ToString()
                        , dtPrv.Rows[0]["SUN"]?.ToString()
                        , dtPrv.Rows[0]["DEALER_NM"]?.ToString()
                        , dtPrv.Rows[0]["GUBUN1"]?.ToString()
                        , dtPrv.Rows[0]["J_BNUM"]?.ToString());

                        sLogMsg += string.Format(" l 금액 : {0} -> {1}", dBeforeAmt, dAfterAmt);
                        if (!string.IsNullOrEmpty(sRemark))
                        {
                            sLogMsg += string.Format(" l 수정사유 : {0}", sRemark);
                        }
                        
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                        strSql.AppendLine("           ( OCCUR_DATE ");
                        strSql.AppendLine("           , USRCD ");
                        strSql.AppendLine("           , LOG_SEQ ");
                        strSql.AppendLine("           , EDIT_KIND ");
                        strSql.AppendLine("           , PGM_ID ");
                        strSql.AppendLine("           , ACS_IP ");
                        strSql.AppendLine("           , EDIT_RMK ) ");
                        strSql.AppendLine("     VALUES( @OCCUR_DATE ");
                        strSql.AppendLine("           , @USRCD ");
                        strSql.AppendLine("           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                        strSql.AppendLine("           , 'U' ");
                        strSql.AppendLine("           , @PGM_ID ");
                        strSql.AppendLine("           , @ACS_IP ");
                        strSql.AppendLine("           , @EDIT_RMK ) ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                        cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                        cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                        cmd.Parameters.AddWithValue("@EDIT_RMK", sLogMsg);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                Cursor = Cursors.Default;

                XtraMessageBox.Show("저장을 완료하였습니다.");
                BtnRetr.PerformClick();
                view.FocusedRowHandle = view.LocateByDisplayText(0, gridColJunpyo, sJunpyoId_Find);
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show("[Save Error] " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewBuyerPurc_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            GridView view = (GridView)sender;
            if(view.RowCount == 0)
            {
                GridDet.DataSource = null;
            }

            string sJunpyoId = view.GetFocusedRowCellValue("JUNPYOID")?.ToString();
            if (string.IsNullOrEmpty(sJunpyoId))
            {
                GridDet.DataSource = null;
            }
            else
            {
                GridDet.DataSource = GetDetInfo(sJunpyoId);
            }
        }

        private void GridViewBuyerSales_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            GridView view = (GridView)sender;
            if (view.RowCount == 0)
            {
                GridDet.DataSource = null;
            }

            string sJunpyoId = view.GetFocusedRowCellValue("JUNPYOID")?.ToString();
            if (string.IsNullOrEmpty(sJunpyoId))
            {
                GridDet.DataSource = null;
            }
            else
            {
                GridDet.DataSource = GetDetInfo(sJunpyoId);
            }
        }

        private DataTable GetDetInfo(string sJunpyoId)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.JUNPYOID ");
            strSql.AppendLine("      , A.JUNPYO_SEQ ");
            strSql.AppendLine("      , A.CHG_AMT_DT ");
            strSql.AppendLine("      , A.PRV_CLO_AMT ");
            strSql.AppendLine("      , A.CUR_CLO_AMT ");
            strSql.AppendLine("      , A.CHG_REMARK ");
            strSql.AppendLine("      , A.CDATE ");
            strSql.AppendLine("      , B1.EMP_NM AS CUSER ");
            strSql.AppendLine("      , A.MDATE ");
            strSql.AppendLine("      , B2.EMP_NM AS MUSER ");
            strSql.AppendLine("   FROM INLIST_LOG A ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1 ");
            strSql.AppendLine("     ON A.CUSER = B1.EMP_ID ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B2 ");
            strSql.AppendLine("     ON A.MUSER = B2.EMP_ID ");
            strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewBuyerPurc_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        /*
         * 전표발행 및 세금계산서발행 구분에 따라 색 다르게 지정
         */
        private void GridViewBuyerPurc_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            int iSlipCnt = 0;
            int.TryParse(GridViewBuyerPurc.GetRowCellValue(e.RowHandle, GridColPurcCnt)?.ToString(), out iSlipCnt);
            string sTaxNo = GridViewBuyerPurc.GetRowCellValue(e.RowHandle, GridColPurcTaxNo)?.ToString().Trim();

            if(iSlipCnt > 0)
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (!string.IsNullOrEmpty(sTaxNo))
            {
                e.Appearance.BackColor = Color.Yellow;
            }
            else
            {
                e.Appearance.BackColor = Color.Azure;
            }
        }

        /*
         * 전표발행 및 세금계산서발행 구분에 따라 색 다르게 지정
         */
        private void GridViewBuyerPurc_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColPurcPrc || e.Column == GridColPurcRemark)
            {
                int iSlipCnt = 0;
                int.TryParse(GridViewBuyerPurc.GetRowCellValue(e.RowHandle, GridColPurcCnt)?.ToString(), out iSlipCnt);
                string sTaxNo = GridViewBuyerPurc.GetRowCellValue(e.RowHandle, GridColPurcTaxNo)?.ToString().Trim();

                if (iSlipCnt > 0)
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else if (!string.IsNullOrEmpty(sTaxNo))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else
                {
                    e.Appearance.BackColor = Color.Azure;
                }
            }
        }

        /*
         * 전표발행 및 세금계산서발행 구분에 따라 색 다르게 지정
         */
        private void GridViewBuyerSales_RowStyle(object sender, RowStyleEventArgs e)
        {
            int iSlipCnt = 0;
            int.TryParse(GridViewBuyerSales.GetRowCellValue(e.RowHandle, GridColSaleCnt)?.ToString(), out iSlipCnt);
            string sTaxNo = GridViewBuyerSales.GetRowCellValue(e.RowHandle, GridColSaleTaxNo)?.ToString().Trim();

            if (iSlipCnt > 0)
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (!string.IsNullOrEmpty(sTaxNo))
            {
                e.Appearance.BackColor = Color.Yellow;
            }
            else
            {
                e.Appearance.BackColor = Color.Azure;
            }
        }

        /*
         * 전표발행 및 세금계산서발행 구분에 따라 색 다르게 지정
         */
        private void GridViewBuyerSales_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColSalePrc || e.Column == GridColSaleRemark)
            {
                int iSlipCnt = 0;
                int.TryParse(GridViewBuyerSales.GetRowCellValue(e.RowHandle, GridColSaleCnt)?.ToString(), out iSlipCnt);
                string sTaxNo = GridViewBuyerSales.GetRowCellValue(e.RowHandle, GridColSaleTaxNo)?.ToString().Trim();

                if (iSlipCnt > 0)
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else if (!string.IsNullOrEmpty(sTaxNo))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else
                {
                    e.Appearance.BackColor = Color.Azure;
                }
            }
        }

        private void CboKeraType_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void IN10001F01_TextChanged(object sender, EventArgs e)
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

        private void CboDateGb_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private DataTable SetLookupEdit()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" WITH INFO AS ( ");
            strSql.AppendLine("      SELECT X1.EMP_ID ");
            strSql.AppendLine("           , X1.EMP_NM ");
            strSql.AppendLine("           , X1.DEPT_CD ");
            strSql.AppendLine("        FROM HR_EMP_BASIS X1 ");
            strSql.AppendLine("       WHERE DEPT_CD IN ('5000', '3000') ");
            strSql.AppendLine("    )  ");
            strSql.AppendLine(" SELECT Y1.CD ");
            strSql.AppendLine("      , Y1.NM ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT '' AS CD ");
            strSql.AppendLine("               , '전체' AS NM ");
            strSql.AppendLine("               , '9999' AS DEPT_CD ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT EMP_ID ");
            strSql.AppendLine("               , EMP_NM ");
            strSql.AppendLine("               , DEPT_CD ");
            strSql.AppendLine("            FROM INFO A1 ");
            strSql.AppendLine("        ) Y1 ");
            strSql.AppendLine("  ORDER BY Y1.DEPT_CD DESC, Y1.CD  ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        /*
         * 전표발행 및 세금계산서는 Edit할 수 없도록 세팅
         */
        private void GridViewBuyerPurc_ShowingEditor(object sender, CancelEventArgs e)
        {
            int iSlipCnt = 0;
            int.TryParse(GridViewBuyerPurc.GetRowCellValue(GridViewBuyerPurc.FocusedRowHandle, GridColPurcCnt)?.ToString(), out iSlipCnt);
            string sTaxNo = GridViewBuyerPurc.GetRowCellValue(GridViewBuyerPurc.FocusedRowHandle, GridColPurcTaxNo)?.ToString().Trim();

            if (iSlipCnt > 0)
            {
                e.Cancel = true;
            }
            else if (!string.IsNullOrEmpty(sTaxNo))
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }

        /*
         * 전표발행 및 세금계산서는 Edit할 수 없도록 세팅
         */
        private void GridViewBuyerSales_ShowingEditor(object sender, CancelEventArgs e)
        {
            int iSlipCnt = 0;
            int.TryParse(GridViewBuyerSales.GetRowCellValue(GridViewBuyerSales.FocusedRowHandle, GridColSaleCnt)?.ToString(), out iSlipCnt);
            string sTaxNo = GridViewBuyerSales.GetRowCellValue(GridViewBuyerSales.FocusedRowHandle, GridColSaleTaxNo)?.ToString().Trim();

            if (iSlipCnt > 0)
            {
                e.Cancel = true;
            }
            else if (!string.IsNullOrEmpty(sTaxNo))
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }

        private void GridViewDet_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewDet_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void IN10001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if(e.KeyCode == Keys.F3)
            {
                BtnSave.PerformClick();
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}