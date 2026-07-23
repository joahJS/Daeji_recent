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
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class AC12001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC12001F01()
        {
            InitializeComponent();
        }
        public DataRow rowUserInfo { get; set; }
        private void AC12001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Now;
            DateEditTo.EditValue = DateTime.Now;
            DateEditFrom1.EditValue = DateTime.Now;
            DateEditTo1.EditValue = DateTime.Now;
        }

        private void GridViewDtlRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewDtlRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewDtlRetr1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewDtlRetr1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void BtnRetr1_Click(object sender, EventArgs e)
        {
            ShowViewM();
            ShowViewD();
        }

        //계산서발행 마스터
        private void ShowViewM()
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("YMD_F", DateEditFrom1.EditValue?.ToString().Replace("-", "").Substring(0, 8));
            dicParams.Add("YMD_T", DateEditTo1.EditValue?.ToString().Replace("-", "").Substring(0, 8));
            dicParams.Add("FIND_IDX", CboFindSbj1.SelectedIndex.ToString());
            dicParams.Add("FIND_WORD", TxtFindWord1.EditValue?.ToString().Trim());

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT X1.* ");
            strSql.AppendLine("      , X1.AMT + X1.VAT AS TOT_AMT #합계금액 ");
            strSql.AppendLine("      , 'N' AS CHK #전체선택 ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT A.J_ID1 #거래처코드 ");
            strSql.AppendLine("               , A.KeraType #거래구분 ");
            strSql.AppendLine("               , A.J_Date #작성일자 ");
            strSql.AppendLine("               , A.JUNPYOID #전표ID ");
            strSql.AppendLine("               , B.DEALER_NM #거래처명 ");
            strSql.AppendLine("               , B.IDT_NO #사업자번호 ");
            strSql.AppendLine("               , B.REP_NM #대표자명 ");
            strSql.AppendLine("               , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE #연락처 ");
            strSql.AppendLine("               , SUM(IFNULL(A.DANJUNG, 0)) AS TOT_WGT #총중량 ");
            strSql.AppendLine("               , SUM(IFNULL(A.KONGKEP, 0)) AS AMT #공급가액 ");
            strSql.AppendLine("               , SUM(CASE C.SEAK_POHAM WHEN 'Y' THEN 0 ELSE IFNULL(A.KONGKEP, 0) END) * 0.1 AS VAT #SEAK_POHAM : Y일 경우 부가세 미포함, N은 포함 ");
            strSql.AppendLine("            FROM INLIST A ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B ");
            strSql.AppendLine("              ON A.J_ID1 = B.DEALER_CD ");
            strSql.AppendLine("            LEFT JOIN MESURING C ");
            strSql.AppendLine("              ON A.J_ID = C.IPCHULGO_MACHULID ");
            strSql.AppendLine("            LEFT JOIN ACC_TAX_MGT D ");
            strSql.AppendLine("              ON A.JUNPYOID = D.BILL_KEY ");
            strSql.AppendLine("           WHERE A.J_DATE BETWEEN  '" + dicParams["YMD_F"] + "'   ");
            strSql.AppendLine("             AND '" + dicParams["YMD_T"] + "'                             ");
            strSql.AppendLine("             AND D.BILL_KEY IS NULL ");
            strSql.AppendLine("             AND A.KERATYPE = '매출' ");
            strSql.AppendLine("             AND C.KERATYPE <> '직송' ");
            strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '0' AND A.J_ID1 LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '1' AND B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '2' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '3' AND B.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ) ");
            strSql.AppendLine("           GROUP BY A.J_ID1 ");
            strSql.AppendLine("        ) X1        ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr1.DataSource = dt;
        }
        //계산서발행 디테일
        private void ShowViewD()
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("YMD_F", DateEditFrom1.EditValue?.ToString().Replace("-", "").Substring(0, 8));
            dicParams.Add("YMD_T", DateEditTo1.EditValue?.ToString().Replace("-", "").Substring(0, 8));
            dicParams.Add("J_ID1", GridViewRetr1.GetFocusedRowCellValue("J_ID1")?.ToString());
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT X1.* ");
            strSql.AppendLine("      , X1.KONGKEP + X1.VAT AS TOT_AMT ");
            strSql.AppendLine("      , 'N' AS CHK ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT A.JUNPYOID #PK값 ");
            strSql.AppendLine("               , A.KeraType #거래구분 ");
            strSql.AppendLine("               , A.J_DATE #마감일자 ");
            strSql.AppendLine("               , B.J_BNUM #차량번호 ");
            strSql.AppendLine("               , A.J_SERIAL #등급코드 컬럼 Visible False 할 것 ");
            strSql.AppendLine("               , C.GUBUN1 #등급 ");
            strSql.AppendLine("               , A.DANJUNG #인수량 ");
            strSql.AppendLine("               , A.DANGA #매출단가 ");
            strSql.AppendLine("               , A.KONGKEP #공급가액 ");
            strSql.AppendLine("               , CASE B.SEAK_POHAM WHEN 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END AS VAT #부가세액 ");
            strSql.AppendLine("            FROM INLIST A ");
            strSql.AppendLine("            LEFT JOIN MESURING B ");
            strSql.AppendLine("              ON A.J_ID = B.IPCHULGO_MACHULID ");
            strSql.AppendLine("            LEFT JOIN JAJAE C ");
            strSql.AppendLine("              ON A.J_SERIAL = C.J_SERIAL ");
            strSql.AppendLine("            LEFT JOIN ACC_TAX_MGT D ");
            strSql.AppendLine("              ON A.JUNPYOID = D.BILL_KEY ");
            strSql.AppendLine("           WHERE A.J_DATE BETWEEN  '" + dicParams["YMD_F"] + "'   ");
            strSql.AppendLine("             AND '" + dicParams["YMD_T"] + "'                     ");
            strSql.AppendLine("             AND D.BILL_KEY IS NULL ");
            strSql.AppendLine("             AND A.KERATYPE = '매출' ");
            strSql.AppendLine("             AND B.KERATYPE = '출고' ");
            strSql.AppendLine("             AND A.J_ID1 = '" + dicParams["J_ID1"] + "' ");
            strSql.AppendLine("        ) X1 ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridDtlRetr1.DataSource = dt;
        }
        private void GridViewRetr1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            ShowViewD();
        }
        //단축키 설정
        private void AC12001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                if (xtraTabControl1.SelectedTabPageIndex == 0)
                {
                    BtnRetr1_Click(null, null);// 세금계산서(F5)
                }
                else
                {
                    BtnRetr1_Click(null, null);// 계산서발행(F5)
                }
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnPub1_Click(null, null);// 발행(F3)
            }
            else if (e.KeyCode == Keys.Escape)
            {
                BtnClose_Click(null, null);// 닫기(ESC)
            }
        }

        private string MakeBillSeq(string sValue)
        {
            if (sValue.Length == 1)
                sValue = "0000000" + sValue;
            else if (sValue.Length == 2)
                sValue = "000000" + sValue;
            else if (sValue.Length == 3)
                sValue = "00000" + sValue;
            else if (sValue.Length == 4)
                sValue = "0000" + sValue;
            else if (sValue.Length == 5)
                sValue = "000" + sValue;
            else if (sValue.Length == 6)
                sValue = "00" + sValue;
            else if (sValue.Length == 7)
                sValue = "0" + sValue;

            return sValue;
        }
        //발행버튼
        private void BtnPub1_Click(object sender, EventArgs e)
        {
            //if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            //{
            //    XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
            //    return;
            //}

            DataTable dt = (DataTable)GridDtlRetr1.DataSource;

            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("승인리스트의 데이터가 존재하지 않습니다.");
                DateEditFrom1.Focus();
                DateEditFrom1.SelectAll();
                return;
            }

            int iCnt = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["CHK"].Equals("Y"))
                {
                    iCnt++;
                }
            }

            if (iCnt == 0)
            {
                XtraMessageBox.Show("세금계산서 발행을 하고자 하는 데이터의 체크박스에 체크하세요.");
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Cursor = Cursors.WaitCursor;

                    string sChkYn = dt.Rows[i]["CHK"].ToString();
                    if (sChkYn.Equals("Y"))
                    {
                        //BILL_SEQ 채번(ACC_TAX_MST)
                        strSql.Clear();
                        strSql.AppendLine(" SELECT CASE WHEN MAX(A.BILL_SEQ) IS NULL THEN '00000001' ELSE MAX(A.BILL_SEQ) END AS MAX_VALUE ");
                        strSql.AppendLine("   FROM ACC_TAX_MGT A    ");
                        strSql.AppendLine("  WHERE A.BILL_GB = '1' ");
                        strSql.AppendLine("    AND A.ISSUE_YY = YEAR(NOW()) ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        Dictionary<string, string> dicParams = new Dictionary<string, string>();

                        dicParams.Add("JUNPYOID", dt.Rows[i]["JUNPYOID"].ToString());
                        dicParams.Add("KERATYPE", dt.Rows[i]["KERATYPE"].ToString());
                        
                        dicParams.Add("J_ID1", GridViewRetr1.GetFocusedRowCellValue("J_ID1")?.ToString());

                        string sMaxValue = cmd.ExecuteScalar()?.ToString();
                        double dMaxValue = Convert.ToDouble(sMaxValue) + 1;
                        sMaxValue = MakeBillSeq(Convert.ToString(dMaxValue));

                        dicParams.Add("sMaxValue", MakeBillSeq(Convert.ToString(dMaxValue)));
                        string sJunpyoId = dt.Rows[i]["JUNPYOID"].ToString();
                        string sKeraType = dt.Rows[i]["KERATYPE"].ToString();

                        #region[ACC_TAX_MGT]

                        if (sKeraType.Equals("매입"))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" INSERT INTO ACC_TAX_MGT ");
                            strSql.AppendLine("           ( BILL_GB, PURCH_SALE_GB, ISSUE_YY, BILL_SEQ, BILL_ISSUE_YMD ");
                            strSql.AppendLine("           , JURIDICAL_GB, DEALER_CD, IDT_NO, DEALER_NM, BIZ_NM, TYPE_NM, REP_NM, DEALER_ADDR, CHRG_NM ");
                            strSql.AppendLine("           , SUPL_AMT, VAT_AMT, TOT_AMT, TAX_GB, ASK_RECPT_GB, VAT_INCLUDE_YN, BILL_PMNT_GB, BILL_KEY ");
                            strSql.AppendLine("           , BILL_ISSUE_DEPT, DEALER_ISSUE_YN, NOTE, ENT_DT, ENT_ID) ");
                            strSql.AppendLine("      SELECT '1' ");
                            strSql.AppendLine("      	  , 'P' ");
                            strSql.AppendLine("      	  , YEAR(NOW()) ");
                            strSql.AppendLine("      	  , '" + dicParams["sMaxValue"] + "' ");
                            strSql.AppendLine("      	  , DATE_FORMAT(NOW(), '%Y%m%d') -- BILL_ISSUE_YMD ");
                            strSql.AppendLine("      	  , 'C' -- 법인사업자 ");
                            strSql.AppendLine("      	  , A.J_ID1 ");
                            strSql.AppendLine("      	  , B.IDT_NO ");
                            strSql.AppendLine("      	  , B.DEALER_NM ");
                            strSql.AppendLine("      	  , B.BIZ_NM ");
                            strSql.AppendLine("      	  , B.TYPE_NM ");
                            strSql.AppendLine("      	  , B.REP_NM ");
                            strSql.AppendLine("      	  , B.ADDR ");
                            strSql.AppendLine("      	  , B.CHRG_NM -- CHRG_NM ");
                            strSql.AppendLine("      	  , A.IKONGKEP ");
                            strSql.AppendLine("      	  , A.BUGASE ");
                            strSql.AppendLine("      	  , (A.IKONGKEP + A.BUGASE) AS TOT_AMT ");
                            strSql.AppendLine("      	  , '21' -- TAX_GB 과세매입 ");
                            strSql.AppendLine("      	  , 'A' -- ASK_RECPT_GB, 청구영수구분 ");
                            strSql.AppendLine("      	  , 'Y' -- VAR_INCLUDE_YN ");
                            strSql.AppendLine("      	  , '4' -- BILL_PMNT_GB, 외상미수금 ");
                            strSql.AppendLine("      	  , A.JUNPYOID ");
                            strSql.AppendLine("      	  , '2000' -- 집행부서 ");
                            strSql.AppendLine("      	  , 'N' -- DEALER_ISSUE_YN ");
                            strSql.AppendLine("      	  , '' -- NOTE ");
                            strSql.AppendLine("      	  , NOW() ");
                            strSql.AppendLine("      	  , '" + FmMainToolBar2.UserID + "' -- ENT_ID ");
                            strSql.AppendLine("        FROM INLIST A ");
                            strSql.AppendLine("        LEFT OUTER JOIN ACC_DEALER_CD B ");
                            strSql.AppendLine("          ON A.J_ID1 = B.DEALER_CD ");
                            strSql.AppendLine("       WHERE A.JUNPYOID = '" + dicParams["JUNPYOID"] + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                            
                            //전표 업데이트(정은영)
                            strSql.Clear();
                            strSql.AppendLine(" UPDATE ACC_TAX_MGT A        ");
                            strSql.AppendLine("   LEFT OUTER JOIN ACTRAN B  ");
                            strSql.AppendLine(" 	ON B.REF1 = '" + dicParams["JUNPYOID"] + "'    ");
                            strSql.AppendLine("    AND B.LINNO = 0          ");
                            strSql.AppendLine("    SET A.SLIP_NO = B.SEQNO  ");
                            strSql.AppendLine(" 	 , A.SLIP_YMD = B.TDATE ");
                            strSql.AppendLine(" 	 , A.SLIP_ISSUE_YN = 'Y'");
                            strSql.AppendLine("  WHERE A.BILL_GB = '1'      ");
                            strSql.AppendLine("    AND A.PURCH_SALE_GB = 'P'");
                            strSql.AppendLine("   AND A.ISSUE_YY = YEAR(NOW())");
                            strSql.AppendLine("   AND A.BILL_SEQ = '" + dicParams["sMaxValue"] + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                        }
                        else if (sKeraType.Equals("매출"))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" INSERT INTO ACC_TAX_MGT ");
                            strSql.AppendLine("           ( BILL_GB, PURCH_SALE_GB, ISSUE_YY, BILL_SEQ, BILL_ISSUE_YMD ");
                            strSql.AppendLine("           , JURIDICAL_GB, DEALER_CD, IDT_NO, DEALER_NM, BIZ_NM, TYPE_NM, REP_NM, DEALER_ADDR, CHRG_NM ");
                            strSql.AppendLine("           , SUPL_AMT, VAT_AMT, TOT_AMT, TAX_GB, ASK_RECPT_GB, VAT_INCLUDE_YN, BILL_PMNT_GB, BILL_KEY ");
                            strSql.AppendLine("           , BILL_ISSUE_DEPT, DEALER_ISSUE_YN, NOTE, ENT_DT, ENT_ID) ");
                            strSql.AppendLine("      SELECT '1' ");
                            strSql.AppendLine("      	  , 'S' ");
                            strSql.AppendLine("      	  , YEAR(NOW()) ");
                            strSql.AppendLine("      	  , '" + dicParams["sMaxValue"] + "' ");
                            strSql.AppendLine("      	  , DATE_FORMAT(NOW(), '%Y%m%d') -- BILL_ISSUE_YMD ");
                            strSql.AppendLine("      	  , 'C' -- 법인사업자 ");
                            strSql.AppendLine("      	  , A.J_ID1 ");
                            strSql.AppendLine("      	  , B.IDT_NO ");
                            strSql.AppendLine("      	  , B.DEALER_NM ");
                            strSql.AppendLine("      	  , B.BIZ_NM ");
                            strSql.AppendLine("      	  , B.TYPE_NM ");
                            strSql.AppendLine("      	  , B.REP_NM ");
                            strSql.AppendLine("      	  , B.ADDR ");
                            strSql.AppendLine("      	  , B.CHRG_NM -- CHRG_NM ");
                            strSql.AppendLine("      	  , A.KONGKEP ");
                            strSql.AppendLine("      	  , A.BUGASE ");
                            strSql.AppendLine("      	  , (A.KONGKEP + A.BUGASE) AS TOT_AMT ");
                            strSql.AppendLine("      	  , '11' -- TAX_GB 과세매입 ");
                            strSql.AppendLine("      	  , 'A' -- ASK_RECPT_GB, 청구영수구분 ");
                            strSql.AppendLine("      	  , 'Y' -- VAR_INCLUDE_YN ");
                            strSql.AppendLine("      	  , '4' -- BILL_PMNT_GB, 외상미수금 ");
                            strSql.AppendLine("      	  , A.JUNPYOID ");
                            strSql.AppendLine("      	  , '2000' -- 집행부서 ");
                            strSql.AppendLine("      	  , 'N' -- DEALER_ISSUE_YN ");
                            strSql.AppendLine("      	  , '' -- NOTE ");
                            strSql.AppendLine("      	  , NOW() ");
                            strSql.AppendLine("      	  , '" + FmMainToolBar2.UserID + "' -- ENT_ID ");
                            strSql.AppendLine("        FROM INLIST A ");
                            strSql.AppendLine("        LEFT OUTER JOIN ACC_DEALER_CD B ");
                            strSql.AppendLine("          ON A.J_ID1 = B.DEALER_CD ");
                            strSql.AppendLine("       WHERE A.JUNPYOID = '" + dicParams["JUNPYOID"] + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //전표 업데이트(정은영)
                            strSql.Clear();
                            strSql.AppendLine(" UPDATE ACC_TAX_MGT A        ");
                            strSql.AppendLine("   LEFT OUTER JOIN ACTRAN B  ");
                            strSql.AppendLine(" 	ON B.REF1 = '" + dicParams["JUNPYOID"] + "'    ");
                            strSql.AppendLine("    AND B.LINNO = 0          ");
                            strSql.AppendLine("    SET A.SLIP_NO = B.SEQNO  ");
                            strSql.AppendLine(" 	 , A.SLIP_YMD = B.TDATE ");
                            strSql.AppendLine(" 	 , A.SLIP_ISSUE_YN = 'Y'");
                            strSql.AppendLine("  WHERE A.BILL_GB = '1'      ");
                            strSql.AppendLine("    AND A.PURCH_SALE_GB = 'S'");
                            strSql.AppendLine("   AND A.ISSUE_YY = YEAR(NOW())");
                            strSql.AppendLine("   AND A.BILL_SEQ = '" + dicParams["sMaxValue"] + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }

                        #endregion[ACC_TAX_MGT]

                        #region[ACC_TAX_DTL]

                        strSql.Clear();
                        strSql.AppendLine(" INSERT INTO ACC_TAX_DTL ");
                        strSql.AppendLine("           ( BILL_GB, PURCH_SALE_GB, ISSUE_YY, BILL_SEQ, DTL_SEQ ");
                        strSql.AppendLine("           , BILL_ISSUE_YMD, ITEM_NM, ITEM_SPEC, ITEM_AMOUNT, ITEM_UNPR ");
                        strSql.AppendLine("           , SUPL_AMT, VAT_AMT, ITEM_NOTE, ENT_DT, ENT_ID ) ");
                        strSql.AppendLine("      SELECT A.BILL_GB ");
                        strSql.AppendLine("      	  , A.PURCH_SALE_GB ");
                        strSql.AppendLine("      	  , A.ISSUE_YY ");
                        strSql.AppendLine("      	  , A.BILL_SEQ ");
                        strSql.AppendLine("      	  , (SELECT CASE WHEN MAX(X1.DTL_SEQ) IS NULL THEN 1 ELSE MAX(X1.DTL_SEQ) + 1 END AS MAX_VALUE ");
                        strSql.AppendLine("      	       FROM ACC_TAX_DTL X1 ");
                        strSql.AppendLine("      	      WHERE X1.BILL_GB = A.BILL_GB ");
                        strSql.AppendLine("      	        AND X1.PURCH_SALE_GB = A.PURCH_SALE_GB ");
                        strSql.AppendLine("      	        AND X1.ISSUE_YY = A.ISSUE_YY ");
                        strSql.AppendLine("      	        AND X1.BILL_SEQ = A.BILL_SEQ ) AS MAX_VALUE ");
                        strSql.AppendLine("      	  , A.BILL_ISSUE_YMD ");
                        strSql.AppendLine("      	  , C.GUBUN1 AS ITEM_NM ");
                        strSql.AppendLine("      	  , '' AS ITEM_SPEC ");
                        strSql.AppendLine("      	  , B.DANJUNG AS ITEM_AMOUNT ");
                        strSql.AppendLine("      	  , B.DANGA AS ITEM_UNPR  ");
                        strSql.AppendLine("      	  , CASE WHEN B.KERATYPE = '매입' THEN B.IKONGKEP ELSE B.KONGKEP END AS SUPL_AMT ");
                        strSql.AppendLine("      	  , B.BUGASE AS VAT_AMT ");
                        strSql.AppendLine("      	  , D.GUMSUBIGO ");
                        strSql.AppendLine("      	  , NOW() ");
                        strSql.AppendLine("      	  , '" + FmMainToolBar2.UserID + "' AS ENT_ID ");
                        strSql.AppendLine("        FROM ACC_TAX_MGT A ");
                        strSql.AppendLine("        LEFT OUTER JOIN INLIST B ");
                        strSql.AppendLine("          ON A.BILL_KEY = B.JUNPYOID ");
                        strSql.AppendLine("        LEFT OUTER JOIN JAJAE C ");
                        strSql.AppendLine("          ON B.J_SERIAL = C.J_SERIAL ");
                        strSql.AppendLine("        LEFT OUTER JOIN MESURING D ");
                        strSql.AppendLine("          ON B.J_RID = D.JUNPYOID ");
                        strSql.AppendLine("       WHERE A.BILL_KEY = '" + dicParams["JUNPYOID"] + "' ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        #endregion[ACC_TAX_DTL]
                    }
                    Cursor = Cursors.Default;
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                Cursor = Cursors.Default;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                BtnRetr1_Click(null, null);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }
        //마스터 선택버튼 미완성
        private void RepoChk_CheckedChanged(object sender, EventArgs e)
        {
            string chkVal = GridViewRetr1.GetFocusedRowCellValue(GridColCheck).ToString();
            DataTable dt = (DataTable)GridDtlRetr1.DataSource;

            if (chkVal.Equals("Y") || chkVal.Equals("True"))
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    GridViewDtlRetr1.SetRowCellValue(i, GridColCheck, "N");
                }
            }
            else if (chkVal.Equals("N") || chkVal.Equals("False"))
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    GridViewDtlRetr1.SetRowCellValue(i, GridColCheck, "Y");
                }
            }
            GridViewDtlRetr1.UpdateCurrentRow();
        }
        //디테일 선택버튼
        private void RepoDChk_CheckedChanged(object sender, EventArgs e)
        {
            string chkVal = GridViewDtlRetr1.GetFocusedRowCellValue(GridColChk1).ToString();
            
            if (chkVal.Equals("Y") || chkVal.Equals("True"))
            {
                GridViewDtlRetr1.SetFocusedRowCellValue(GridColChk1, "N");
            }
            else if (chkVal.Equals("N") || chkVal.Equals("False"))
            {
                GridViewDtlRetr1.SetFocusedRowCellValue(GridColChk1, "Y");
            }
            GridViewDtlRetr1.UpdateCurrentRow();
        }
        //닫기버튼
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}
