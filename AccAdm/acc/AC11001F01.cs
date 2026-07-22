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

/*
 * 
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-01
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            매입 (기존) - 차액 = 지급단가 - 기준단가
 *                 (수정) - 차액 = 도착도단가 - 기준단가
 *            매출 (기존) - 차액 = 매출단가 - 기준단가
 *                 (수정) - 차액 = 도착도단가 - 기준단가
 *            직송 (기존) - 차액 = 매출단가 - 매입단가
 *                 (수정) - 차액 = 매출단가 - 도착도단가
 *            해당사항에 대하여 조회쿼리 수정
 *            
 * 수정일자 : 2021-02-23
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            INLIST에 이력성 컬럼인 CHRG_ID(담당자)로 업체담당자 표기하도록 쿼리 수정
 *            
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2021-03-17
 * 수정자   : 고혜성
 * Reference Key = #0001
 * 수정내용 : (현업요청)
 *            1. 로그수정
 *               1-1) 기본사항/변경사항을 나누어 입력
 */
namespace AccAdm
{
    public partial class AC11001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC11001F01()
        {
            InitializeComponent();
        }

        public bool APRV_YN;

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC11001F01_Load(object sender, EventArgs e)
        {
            //APRV_YN = GetApprovalInfo();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            ComnEtcFunc.SetDateFromToValue(DateEditFrom, DateEditTo);
            ComnGridFunc.SetInitGridRowColor(GridViewBuyerPurc);
            ComnGridFunc.SetInitGridRowColor(GridViewBuyerSales);
            ComnGridFunc.SetInitGridRowColor(GridViewBuyerDtSend);
            SetButtonVisible();

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewBuyerPurc, GridViewBuyerSales, GridViewBuyerDtSend };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);

            BtnRetr_Click(null, null);

            DataTable dt = GetLookUpData();
            ComGrid.SetLookUpEdit(LkupDealerNm, dt, "CD", "CD", "Y");
            LkupDealerNm.Properties.PopulateColumns();
            LkupDealerNm.Properties.Columns[1].Visible = false;
        }

        private DataTable GetLookUpData()
        {
            string sCboKeraType = CboKeraType.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");
            strSql.AppendLine(" SELECT '' AS CD");
            strSql.AppendLine("     , '' AS NM");
            strSql.AppendLine(" UNION ALL");

            if (sCboKeraType.Equals("직송"))
            {
                strSql.AppendLine(" SELECT C.DEALER_NM AS CD       ");
                strSql.AppendLine("       ,'' AS NM");
                strSql.AppendLine(" FROM INLIST A                  ");
                strSql.AppendLine(" LEFT OUTER JOIN MESURING B     ");
                strSql.AppendLine(" ON B.IPCHULGO_MACHULID = A.J_ID");
                strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD C");
                strSql.AppendLine(" ON A.J_ID1 = C.DEALER_CD       ");
                strSql.AppendLine(" WHERE B.KERATYPE = '직송'      ");
                strSql.AppendLine(" AND A.KERATYPE = '매출'        ");
                strSql.AppendLine(" GROUP BY C.DEALER_NM                    ");

            }
            else
            {
                strSql.AppendLine(" SELECT A.DEALER_NM AS CD");
                strSql.AppendLine(" 	  , '' AS NM        ");
                strSql.AppendLine(" FROM ACC_DEALER_CD A    ");
                strSql.AppendLine(" LEFT OUTER JOIN INLIST B");
                strSql.AppendLine(" ON B.J_ID1 = A.DEALER_CD");
                strSql.AppendLine(" WHERE B.KERATYPE = '" + sCboKeraType + "'");
                strSql.AppendLine(" GROUP BY A.DEALER_NM             ");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            ComnEtcFunc.YmdFromToValuesCheck(DateEditFrom, DateEditTo);
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            string sDealerNm = LkupDealerNm.EditValue?.ToString();
            string sKeraType = CboKeraType.EditValue?.ToString();
            string sAprvGb = RdgbAprvGb.EditValue == null ? string.Empty : RdgbAprvGb.EditValue.ToString();

            if (sKeraType.Equals("매입"))
            {
                PurchacePerBuyer(sYmdFrom, sYmdTo, sDealerNm, sAprvGb);
            }
            else if (sKeraType.Equals("매출"))
            {
                SalesPerBuyer(sYmdFrom, sYmdTo, sDealerNm, sAprvGb);
            }
            else if (sKeraType.Equals("직송"))
            {
                DirectSendPerBuyer(sYmdFrom, sYmdTo, sDealerNm, sAprvGb);   
            }

        }

        private void BtnAprvClose_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (CheckApproval())
                return;

            DataTable dt = (DataTable)GridBuyer.DataSource;
            int[] selectedRows = GridViewBuyerPurc.GetSelectedRows();

            if (GridBuyer.FocusedView == GridViewBuyerPurc)
                selectedRows = GridViewBuyerPurc.GetSelectedRows();
            else if (GridBuyer.FocusedView == GridViewBuyerSales)
                selectedRows = GridViewBuyerSales.GetSelectedRows();
            else if (GridBuyer.FocusedView == GridViewBuyerDtSend)
                selectedRows = GridViewBuyerDtSend.GetSelectedRows();

            if(selectedRows.Length == 0)
            {
                XtraMessageBox.Show("승인하려는 데이터에 체크하세요.");
                return;
            }
            
            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("조회내역이 존재하지 않습니다.");
                DateEditFrom.Focus();
                DateEditFrom.SelectAll();
                return;
            }

            if (GridBuyer.FocusedView == GridViewBuyerPurc)
            {
                selectedRows = GridViewBuyerPurc.GetSelectedRows();
                IssueingSlipOfPurcAndSale(selectedRows, dt);
            }
            else if (GridBuyer.FocusedView == GridViewBuyerSales)
            {
                selectedRows = GridViewBuyerSales.GetSelectedRows();
                IssueingSlipOfPurcAndSale(selectedRows, dt);
            }
            else if (GridBuyer.FocusedView == GridViewBuyerDtSend)
            {
                selectedRows = GridViewBuyerDtSend.GetSelectedRows();
                IssueingSlipOfDirectSend(selectedRows, dt);
            }
            //int iCnt = 0;
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    string sChk = dt.Rows[i]["CHK"].ToString();
            //    if (sChk.Equals("Y"))
            //    {
            //        Cursor = Cursors.WaitCursor;
            //        iCnt++;
            //        Cursor = Cursors.Default;
            //    }
            //}

            //if (iCnt == 0)
            //{
            //    XtraMessageBox.Show("승인하고자 하는 데이터의 체크를 진행하세요.");
            //    return;
            //}


        }

        /// <summary>
        ///     매입출 관련 전표생성
        /// </summary>
        /// <param name="selectedRows">GridView RowHandle(index)</param>
        /// <param name="dt">GridControl</param>
        private void IssueingSlipOfPurcAndSale(int[] selectedRows, DataTable dt)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selectedRows.Length; i++)
                {
                    DataRow row = dt.NewRow();

                    if (GridBuyer.FocusedView == GridViewBuyerPurc)
                        row = GridViewBuyerPurc.GetDataRow(selectedRows[i]);
                    else if (GridBuyer.FocusedView == GridViewBuyerSales)
                        row = GridViewBuyerSales.GetDataRow(selectedRows[i]);
                    else if (GridBuyer.FocusedView == GridViewBuyerDtSend)
                        row = GridViewBuyerDtSend.GetDataRow(selectedRows[i]);

                    string sJunpyoId = row["JUNPYOID"].ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE INLIST ");
                    strSql.AppendLine("    SET APRV_YN = NULL ");
                    strSql.AppendLine("      , APRV_DATE = NULL ");
                    strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                    cmd.ExecuteNonQuery();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE INLIST ");
                    strSql.AppendLine("    SET APRV_YN = 'Y' ");
                    strSql.AppendLine("      , APRV_DATE = CONVERT(VARCHAR(10),GETDATE(),23) ");
                    strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                    cmd.ExecuteNonQuery();

                    #region [마감승인시 전표생성되도록 변경 2020-06-05 정은영]

                    string sCboKeraType = CboKeraType.EditValue.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" "); //Y : 부가세 미포함, N : 부가세 포함
                    strSql.AppendLine(" SELECT CASE WHEN SEAKPOHAM = 'Y' THEN SEAKPOHAM ELSE 'N' END AS YN ");
                    strSql.AppendLine("   FROM INLIST ");
                    strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + "");

                    DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    string Yn = dtChk.Rows[0]["YN"]?.ToString();

                    //매출
                    if (sCboKeraType.Equals("매출"))
                    {
                        if (Yn.Equals("Y"))
                        {
                            #region[부가세 미포함 시 수행되는 로직]

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            #region mariaDB
                            //strSql.AppendLine(" -- 전표 매출 경우 ");
                            //strSql.AppendLine(" INSERT INTO ACTRAN ");
                            //strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO, ACCOD ");
                            //strSql.AppendLine("           , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT ");
                            //strSql.AppendLine("           , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER ");
                            //strSql.AppendLine("           , REF1, REF2, REF3, RK, CUSER, CDATE ) ");
                            //strSql.AppendLine(" SELECT DATE_FORMAT(B.J_DATE, '%Y%m%d') AS TDATE ");
                            //strSql.AppendLine(" 	 , '3' AS ATGUB  ");
                            //strSql.AppendLine(" 	 , (  ");
                            //strSql.AppendLine("          SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE LPAD(CAST((CAST(MAX(X1.SEQNO) AS INT) + 1) AS CHAR), 4, '0') END AS SEQNO ");
                            //strSql.AppendLine("            FROM ACTRAN X1 ");
                            //strSql.AppendLine("           WHERE TDATE = DATE_FORMAT(B.J_DATE, '%Y%m%d') ");
                            //strSql.AppendLine("             AND ATGUB = '3' -- HARDCODING, '대체' ");
                            //strSql.AppendLine("             AND SEQNO >= 5000 ");
                            //strSql.AppendLine("        ) AS SEQNO ");
                            //strSql.AppendLine("      , A.SEQNO AS LINNO ");
                            //strSql.AppendLine("      , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD ");
                            //strSql.AppendLine("      , C.ACNAM ");
                            //strSql.AppendLine("      , B.J_ID1 AS CVCOD ");
                            //strSql.AppendLine("      , D.DEALER_NM AS CVNAM ");
                            //strSql.AppendLine("      , CONCAT(D.DEALER_NM, ',', E.GUBUN1 , ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT ");
                            //strSql.AppendLine("      , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP ");
                            //strSql.AppendLine("             WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO = 2 THEN B.BUGASE END AS ADAMT  ");
                            //strSql.AppendLine("      , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP END AS ACAMT ");
                            //strSql.AppendLine("      , F2.DEPT_CD AS ADPCD ");
                            //strSql.AppendLine("      , G.DEPT_NM AS ADPNM ");
                            //strSql.AppendLine("      , @AAUTO AS AAUTO "); //HARDCODING 자동전표구분 A01
                            //strSql.AppendLine("      , 'Y' AS APVYN ");
                            //strSql.AppendLine("      , NOW() AS ADATE ");
                            //strSql.AppendLine("      , F2.EMP_ID AS AUSER ");
                            //strSql.AppendLine("      , B.JUNPYOID AS REF1 "); //INLIST JUNPYOID
                            //strSql.AppendLine("      , B.J_ID AS REF2 "); //INLIST J_ID
                            //strSql.AppendLine("      , E.JUNPYOID AS REF3 "); //MESURING JUNPYOID 
                            //strSql.AppendLine("      , '' AS RK ");
                            //strSql.AppendLine("      , F2.EMP_ID AS CUSER ");
                            //strSql.AppendLine("      , NOW() AS CDATE ");
                            //strSql.AppendLine("   FROM ACBUNF A ");
                            //strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
                            //strSql.AppendLine("     ON B.JUNPYOID = @JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF C ");
                            //strSql.AppendLine("     ON A.CACOD = C.ACCOD ");
                            //strSql.AppendLine("     OR A.DACOD = C.ACCOD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
                            //strSql.AppendLine("     ON B.J_ID1 = D.DEALER_CD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN MESURING E ");
                            //strSql.AppendLine("     ON B.J_RID = E.JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST F1 ");
                            //strSql.AppendLine("     ON F1.USRCD = @USRCD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS F2 ");
                            //strSql.AppendLine("     ON F2.EMP_ID = F1.INSANO ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD G ");
                            //strSql.AppendLine("     ON F2.DEPT_CD = G.DEPT_CD ");
                            //strSql.AppendLine("  WHERE A.BUNCD = '12' "); //SEAKPOHAM(부가세여부 Y 일 경우 미포함으로 12 수행)
                            #endregion

                            strSql.AppendLine("INSERT INTO ACTRAN                                          ");
                            strSql.AppendLine("         (TDATE, ATGUB, SEQNO, LINNO, ACCOD                 ");
                            strSql.AppendLine("         , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT                ");
                            strSql.AppendLine("         , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER  ");
                            strSql.AppendLine("         , REF1, REF2, REF3, RK, CUSER, CDATE)              ");
                            strSql.AppendLine("SELECT CONVERT(VARCHAR(8),CONVERT(DATE,B.J_DATE),112) AS TDATE            ");
                            strSql.AppendLine("     , '3' AS ATGUB                                         ");
                            strSql.AppendLine("     , (                                                    ");
                            strSql.AppendLine("         SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE REPLICATE(0, 4 - LEN(MAX(X1.SEQNO))) + (MAX(X1.SEQNO)+1) END AS SEQNO");
                            strSql.AppendLine("           FROM ACTRAN X1                                    ");
                            strSql.AppendLine("          WHERE TDATE = CONVERT(VARCHAR(8), CONVERT(DATE,B.J_DATE), 112)   ");
                            strSql.AppendLine("            AND ATGUB = '3'-- HARDCODING, '대체'             ");
                            strSql.AppendLine("            AND SEQNO >= 5000                                ");
                            strSql.AppendLine("       ) AS SEQNO                                            ");
                            strSql.AppendLine("     , A.SEQNO AS LINNO                                      ");
                            strSql.AppendLine("     , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD");
                            strSql.AppendLine("     , C.ACNAM                ");
                            strSql.AppendLine("     , B.J_ID1 AS CVCOD       ");
                            strSql.AppendLine("     , D.DEALER_NM AS CVNAM   ");
                            strSql.AppendLine("     , CONCAT(D.DEALER_NM, ',', E.GUBUN1, ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT");
                            strSql.AppendLine("    , CASE WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP               ");
                            strSql.AppendLine("            WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO = 2 THEN B.BUGASE END AS ADAMT   ");
                            strSql.AppendLine("    , CASE WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP END AS ACAMT  ");
                            strSql.AppendLine("   , F2.DEPT_CD AS ADPCD              ");
                            strSql.AppendLine("     , G.DEPT_NM AS ADPNM             ");
                            strSql.AppendLine("     , @AAUTO AS AAUTO                 ");
                            strSql.AppendLine("     , 'Y' AS APVYN                   ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS ADATE             ");
                            strSql.AppendLine("     , F2.EMP_ID AS AUSER             ");
                            strSql.AppendLine("     , B.JUNPYOID AS REF1             ");
                            strSql.AppendLine("     , CONVERT(VARCHAR,CONVERT(DECIMAL, B.J_ID)) AS REF2                 ");
                            strSql.AppendLine("     , E.JUNPYOID AS REF3             ");
                            strSql.AppendLine("     , '' AS RK                       ");
                            strSql.AppendLine("     , F2.EMP_ID AS CUSER             ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE             ");
                            strSql.AppendLine("  FROM ACBUNF A                       ");
                            strSql.AppendLine("  LEFT OUTER JOIN INLIST B            ");
                            strSql.AppendLine("    ON B.JUNPYOID = @JUNPYOID         ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACMSTF C            ");
                            strSql.AppendLine("    ON A.CACOD = C.ACCOD              ");
                            strSql.AppendLine("    OR A.DACOD = C.ACCOD              ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D     ");
                            strSql.AppendLine("    ON B.J_ID1 = D.DEALER_CD          ");
                            strSql.AppendLine("  LEFT OUTER JOIN MESURING E          ");
                            strSql.AppendLine("    ON B.J_RID = E.JUNPYOID           ");
                            strSql.AppendLine("  LEFT OUTER JOIN ZUSRLST F1          ");
                            strSql.AppendLine("    ON F1.USRCD = @USRCD              ");
                            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS F2     ");
                            strSql.AppendLine("    ON F2.EMP_ID = F1.INSANO          ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEPT_CD G       ");
                            strSql.AppendLine("    ON F2.DEPT_CD = G.DEPT_CD         ");
                            strSql.AppendLine(" WHERE A.BUNCD = '12'                 ");

                            cmd.Parameters.Clear();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.Parameters.AddWithValue("@AAUTO", "A01");
                            cmd.Parameters.AddWithValue("@JUNPYOID", Convert.ToDouble(sJunpyoId));
                            cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                            cmd.ExecuteNonQuery();

                            #endregion[부가세 미포함 시 수행되는 로직]
                        }
                        else
                        {
                            #region[부가세 포함 시 수행되는 로직]

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            #region mariaDB
                            //strSql.AppendLine(" -- 전표 매출 경우 ");
                            //strSql.AppendLine(" INSERT INTO ACTRAN ");
                            //strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO, ACCOD ");
                            //strSql.AppendLine("           , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT ");
                            //strSql.AppendLine("           , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER ");
                            //strSql.AppendLine("           , REF1, REF2, REF3, RK, CUSER, CDATE ) ");
                            //strSql.AppendLine(" SELECT DATE_FORMAT(B.J_DATE, '%Y%m%d') AS TDATE ");
                            //strSql.AppendLine(" 	 , '3' AS ATGUB  ");
                            //strSql.AppendLine(" 	 , (  ");
                            //strSql.AppendLine("          SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE LPAD(CAST((CAST(MAX(X1.SEQNO) AS INT) + 1) AS CHAR), 4, '0') END AS SEQNO ");
                            //strSql.AppendLine("            FROM ACTRAN X1 ");
                            //strSql.AppendLine("           WHERE TDATE = DATE_FORMAT(B.J_DATE, '%Y%m%d') ");
                            //strSql.AppendLine("             AND ATGUB = '3' -- HARDCODING, '대체' ");
                            //strSql.AppendLine("             AND SEQNO >= 5000 ");
                            //strSql.AppendLine("        ) AS SEQNO ");
                            //strSql.AppendLine("      , A.SEQNO AS LINNO ");
                            //strSql.AppendLine("      , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD ");
                            //strSql.AppendLine("      , C.ACNAM ");
                            //strSql.AppendLine("      , B.J_ID1 AS CVCOD ");
                            //strSql.AppendLine("      , D.DEALER_NM AS CVNAM ");
                            //strSql.AppendLine("      , CONCAT(D.DEALER_NM, ',', E.GUBUN1 , ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT ");
                            //strSql.AppendLine("      , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP ");
                            //strSql.AppendLine("             WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO = 2 THEN B.BUGASE END AS ADAMT  ");
                            //strSql.AppendLine("      , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP + B.BUGASE END AS ACAMT ");
                            //strSql.AppendLine("      , F2.DEPT_CD AS ADPCD ");
                            //strSql.AppendLine("      , G.DEPT_NM AS ADPNM ");
                            //strSql.AppendLine("      , @AAUTO AS AAUTO "); //HARDCODING 자동전표구분 A01
                            //strSql.AppendLine("      , 'Y' AS APVYN ");
                            //strSql.AppendLine("      , NOW() AS ADATE ");
                            //strSql.AppendLine("      , F2.EMP_ID AS AUSER ");
                            //strSql.AppendLine("      , B.JUNPYOID AS REF1 "); //INLIST JUNPYOID
                            //strSql.AppendLine("      , B.J_ID AS REF2 "); //INLIST J_ID
                            //strSql.AppendLine("      , E.JUNPYOID AS REF3 "); //MESURING JUNPYOID 
                            //strSql.AppendLine("      , '' AS RK ");
                            //strSql.AppendLine("      , F2.EMP_ID AS CUSER ");
                            //strSql.AppendLine("      , NOW() AS CDATE ");
                            //strSql.AppendLine("   FROM ACBUNF A ");
                            //strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
                            //strSql.AppendLine("     ON B.JUNPYOID = @JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF C ");
                            //strSql.AppendLine("     ON A.CACOD = C.ACCOD ");
                            //strSql.AppendLine("     OR A.DACOD = C.ACCOD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
                            //strSql.AppendLine("     ON B.J_ID1 = D.DEALER_CD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN MESURING E ");
                            //strSql.AppendLine("     ON B.J_RID = E.JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST F1 ");
                            //strSql.AppendLine("     ON F1.USRCD = @USRCD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS F2 ");
                            //strSql.AppendLine("     ON F2.EMP_ID = F1.INSANO ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD G ");
                            //strSql.AppendLine("     ON F2.DEPT_CD = G.DEPT_CD ");
                            //strSql.AppendLine("  WHERE A.BUNCD = '11' "); //SEAKPOHAM(부가세여부 Y 일 경우 미포함으로 12 수행)
                            #endregion

                            strSql.AppendLine("INSERT INTO ACTRAN                                                                                                                ");
                            strSql.AppendLine("          (TDATE, ATGUB, SEQNO, LINNO, ACCOD                                                                                      ");
                            strSql.AppendLine("          , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT                                                                                     ");
                            strSql.AppendLine("          , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER                                                                       ");
                            strSql.AppendLine("          , REF1, REF2, REF3, RK, CUSER, CDATE)                                                                                   ");
                            strSql.AppendLine("SELECT CONVERT(VARCHAR(8),CONVERT(DATE,B.J_DATE),112) AS TDATE                                                                                  ");
                            strSql.AppendLine("     , '3' AS ATGUB                                                                                                               ");
                            strSql.AppendLine("     , (                                                                                                                          ");
                            strSql.AppendLine("         SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE REPLICATE(0, 4 - LEN(MAX(X1.SEQNO))) + (MAX(X1.SEQNO)+1) END AS SEQNO   ");
                            strSql.AppendLine("           FROM ACTRAN X1                                                                                                         ");
                            strSql.AppendLine("          WHERE TDATE = CONVERT(VARCHAR(8), CONVERT(DATE,B.J_DATE), 112)                                                                        ");
                            strSql.AppendLine("            AND ATGUB = '3'-- HARDCODING, '대체'                                                                                  ");
                            strSql.AppendLine("            AND SEQNO >= 5000                                                                                                     ");
                            strSql.AppendLine("       ) AS SEQNO                                                                                                                 ");
                            strSql.AppendLine("     , A.SEQNO AS LINNO                                                                                                           ");
                            strSql.AppendLine("     , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD                                           ");
                            strSql.AppendLine("     , C.ACNAM                                                                                                                    ");
                            strSql.AppendLine("     , B.J_ID1 AS CVCOD                                                                                                           ");
                            strSql.AppendLine("     , D.DEALER_NM AS CVNAM                                                                                                       ");
                            strSql.AppendLine("     , CONCAT(D.DEALER_NM, ',', E.GUBUN1, ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT                                   ");
                            strSql.AppendLine("     , CASE WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP                                                  ");
                            strSql.AppendLine("            WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO = 2 THEN B.BUGASE END AS ADAMT                                      ");
                            strSql.AppendLine("     , CASE WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP + B.BUGASE END AS ACAMT                          ");
                            strSql.AppendLine("     , F2.DEPT_CD AS ADPCD                                                                                                        ");
                            strSql.AppendLine("     , G.DEPT_NM AS ADPNM                 ");
                            strSql.AppendLine("     , @AAUTO AS AAUTO                     ");
                            strSql.AppendLine("     , 'Y' AS APVYN                       ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS ADATE                 ");
                            strSql.AppendLine("     , F2.EMP_ID AS AUSER                 ");
                            strSql.AppendLine("     , B.JUNPYOID AS REF1                 ");
                            strSql.AppendLine("     , CONVERT(VARCHAR,CONVERT(DECIMAL, B.J_ID)) AS REF2                   ");
                            strSql.AppendLine("     , E.JUNPYOID AS REF3                 ");
                            strSql.AppendLine("     , '' AS RK                           ");
                            strSql.AppendLine("     , F2.EMP_ID AS CUSER                 ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE                 ");
                            strSql.AppendLine("  FROM ACBUNF A                           ");
                            strSql.AppendLine("  LEFT OUTER JOIN INLIST B                ");
                            strSql.AppendLine("    ON B.JUNPYOID = @JUNPYOID                    ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACMSTF C                ");
                            strSql.AppendLine("    ON A.CACOD = C.ACCOD                  ");
                            strSql.AppendLine("    OR A.DACOD = C.ACCOD                  ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D         ");
                            strSql.AppendLine("    ON B.J_ID1 = D.DEALER_CD              ");
                            strSql.AppendLine("  LEFT OUTER JOIN MESURING E              ");
                            strSql.AppendLine("    ON B.J_RID = E.JUNPYOID               ");
                            strSql.AppendLine("  LEFT OUTER JOIN ZUSRLST F1              ");
                            strSql.AppendLine("    ON F1.USRCD = @USRCD                       ");
                            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS F2         ");
                            strSql.AppendLine("    ON F2.EMP_ID = F1.INSANO              ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEPT_CD G           ");
                            strSql.AppendLine("    ON F2.DEPT_CD = G.DEPT_CD             ");
                            strSql.AppendLine(" WHERE A.BUNCD = '11'                                                                                                             ");
                                                                                                                                                                                
                            cmd.Parameters.Clear();                                                                                                                             
                            cmd.CommandType = CommandType.Text;                                                                                                                 
                            cmd.CommandText = strSql.ToString();                                                                                                                
                            cmd.Parameters.AddWithValue("@AAUTO", "A01");
                            cmd.Parameters.AddWithValue("@JUNPYOID", Convert.ToDouble(sJunpyoId));
                            cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                            cmd.ExecuteNonQuery();

                            #endregion[부가세 포함 시 수행되는 로직]
                        }
                    }
                    else if (sCboKeraType.Equals("매입"))    //매입
                    {
                        if (Yn.Equals("Y"))
                        {
                            #region[부가세 미포함 시 수행되는 로직]

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" -- 전표 매입 경우 ");
                            #region mariaDB
                            //strSql.AppendLine(" INSERT INTO ACTRAN ");
                            //strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO, ACCOD ");
                            //strSql.AppendLine("           , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT ");
                            //strSql.AppendLine("           , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER ");
                            //strSql.AppendLine("           , REF1, REF2, REF3, RK, CUSER, CDATE ) ");
                            //strSql.AppendLine(" SELECT DATE_FORMAT(B.J_DATE, '%Y%m%d') AS TDATE ");
                            //strSql.AppendLine(" 	 , '3' AS ATGUB  ");
                            //strSql.AppendLine(" 	 , (  ");
                            //strSql.AppendLine("          SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE LPAD(CAST((CAST(MAX(X1.SEQNO) AS INT) + 1) AS CHAR), 4, '0') END AS SEQNO ");
                            //strSql.AppendLine("            FROM ACTRAN X1 ");
                            //strSql.AppendLine("           WHERE TDATE = DATE_FORMAT(B.J_DATE, '%Y%m%d') ");
                            //strSql.AppendLine("             AND ATGUB = '3' -- HARDCODING, '대체' ");
                            //strSql.AppendLine("             AND SEQNO >= 5000 ");
                            //strSql.AppendLine("        ) AS SEQNO ");
                            //strSql.AppendLine("      , A.SEQNO AS LINNO ");
                            //strSql.AppendLine("      , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD ");
                            //strSql.AppendLine("      , C.ACNAM ");
                            //strSql.AppendLine("      , B.J_ID1 AS CVCOD ");
                            //strSql.AppendLine("      , D.DEALER_NM AS CVNAM ");
                            //strSql.AppendLine("      , CONCAT(D.DEALER_NM, ',', E.GUBUN1 , ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT ");
                            //strSql.AppendLine("      , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP END AS ADAMT -- 주의  ");
                            //strSql.AppendLine("      , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP  ");
                            //strSql.AppendLine("             WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO = 1 THEN B.BUGASE END AS ACAMT -- 주의 ");
                            //strSql.AppendLine("      , F2.DEPT_CD AS ADPCD ");
                            //strSql.AppendLine("      , G.DEPT_NM AS ADPNM ");
                            //strSql.AppendLine("      , @AAUTO AS AAUTO ");
                            //strSql.AppendLine("      , 'Y' AS APVYN ");
                            //strSql.AppendLine("      , NOW() AS ADATE ");
                            //strSql.AppendLine("      , F2.EMP_ID AS AUSER ");
                            //strSql.AppendLine("      , B.JUNPYOID AS REF1 ");
                            //strSql.AppendLine("      , B.J_ID AS REF2 ");
                            //strSql.AppendLine("      , E.JUNPYOID AS REF3 ");
                            //strSql.AppendLine("      , '' AS RK ");
                            //strSql.AppendLine("      , F2.EMP_ID AS CUSER ");
                            //strSql.AppendLine("      , NOW() AS CDATE ");
                            //strSql.AppendLine("   FROM ACBUNF A ");
                            //strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
                            //strSql.AppendLine("     ON B.JUNPYOID = @JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF C ");
                            //strSql.AppendLine("     ON A.CACOD = C.ACCOD ");
                            //strSql.AppendLine("     OR A.DACOD = C.ACCOD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
                            //strSql.AppendLine("     ON B.J_ID1 = D.DEALER_CD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN MESURING E ");
                            //strSql.AppendLine("     ON B.J_RID = E.JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST F1 ");
                            //strSql.AppendLine("     ON F1.USRCD = @USRCD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS F2 ");
                            //strSql.AppendLine("     ON F2.EMP_ID = F1.INSANO ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD G ");
                            //strSql.AppendLine("     ON F2.DEPT_CD = G.DEPT_CD ");
                            //strSql.AppendLine("  WHERE A.BUNCD = '22' "); //SEAKPOHAM(부가세여부 Y 일 경우 미포함으로 12 수행)
                            #endregion
                                                                                                              
                            strSql.AppendLine("INSERT INTO ACTRAN                                             ");
                            strSql.AppendLine("         (TDATE, ATGUB, SEQNO, LINNO, ACCOD                    ");
                            strSql.AppendLine("         , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT                   ");
                            strSql.AppendLine("         , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER     ");
                            strSql.AppendLine("         , REF1, REF2, REF3, RK, CUSER, CDATE)                 ");
                            strSql.AppendLine("SELECT CONVERT(VARCHAR(8),CONVERT(DATE,B.J_DATE),112) AS TDATE               ");
                            strSql.AppendLine("     , '3' AS ATGUB                                            ");
                            strSql.AppendLine("     , (                                                       ");
                            strSql.AppendLine("         SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE REPLICATE(0, 4 - LEN(MAX(X1.SEQNO))) + (MAX(X1.SEQNO)+1) END AS SEQNO");
                            strSql.AppendLine("           FROM ACTRAN X1                                  ");
                            strSql.AppendLine("          WHERE TDATE = CONVERT(VARCHAR(8), CONVERT(DATE,B.J_DATE), 112) ");
                            strSql.AppendLine("            AND ATGUB = '3'-- HARDCODING, '대체'           ");
                            strSql.AppendLine("            AND SEQNO >= 5000                              ");
                            strSql.AppendLine("       ) AS SEQNO                                          ");
                            strSql.AppendLine("     , A.SEQNO AS LINNO                                    ");
                            strSql.AppendLine("     , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD");
                            strSql.AppendLine("     , C.ACNAM               ");
                            strSql.AppendLine("     , B.J_ID1 AS CVCOD      ");
                            strSql.AppendLine("     , D.DEALER_NM AS CVNAM  ");
                            strSql.AppendLine("     , CONCAT(D.DEALER_NM, ',', E.GUBUN1, ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT ");
                            strSql.AppendLine("     , CASE WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP END AS ADAMT --주의");
                            strSql.AppendLine("     , CASE WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP");
                            strSql.AppendLine("            WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO = 1 THEN B.BUGASE END AS ACAMT --주의");
                            strSql.AppendLine("     , F2.DEPT_CD AS ADPCD        ");
                            strSql.AppendLine("     , G.DEPT_NM AS ADPNM         ");
                            strSql.AppendLine("     , @AAUTO AS AAUTO             ");
                            strSql.AppendLine("     , 'Y' AS APVYN               ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS ADATE         ");
                            strSql.AppendLine("     , F2.EMP_ID AS AUSER         ");
                            strSql.AppendLine("     , B.JUNPYOID AS REF1         ");
                            strSql.AppendLine("     , CONVERT(VARCHAR,CONVERT(DECIMAL, B.J_ID)) AS REF2             ");
                            strSql.AppendLine("     , E.JUNPYOID AS REF3         ");
                            strSql.AppendLine("     , '' AS RK                   ");
                            strSql.AppendLine("     , F2.EMP_ID AS CUSER         ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE         ");
                            strSql.AppendLine("  FROM ACBUNF A                   ");
                            strSql.AppendLine("  LEFT OUTER JOIN INLIST B        ");
                            strSql.AppendLine("    ON B.JUNPYOID = @JUNPYOID            ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACMSTF C        ");
                            strSql.AppendLine("    ON A.CACOD = C.ACCOD          ");
                            strSql.AppendLine("    OR A.DACOD = C.ACCOD          ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D ");
                            strSql.AppendLine("    ON B.J_ID1 = D.DEALER_CD      ");
                            strSql.AppendLine("  LEFT OUTER JOIN MESURING E      ");
                            strSql.AppendLine("    ON B.J_RID = E.JUNPYOID       ");
                            strSql.AppendLine("  LEFT OUTER JOIN ZUSRLST F1      ");
                            strSql.AppendLine("    ON F1.USRCD = @USRCD               ");
                            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS F2 ");
                            strSql.AppendLine("    ON F2.EMP_ID = F1.INSANO      ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEPT_CD G   ");
                            strSql.AppendLine("    ON F2.DEPT_CD = G.DEPT_CD     ");
                            strSql.AppendLine(" WHERE A.BUNCD = '22'             ");

                            cmd.Parameters.Clear();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.Parameters.AddWithValue("@AAUTO", "A01");
                            cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                            cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                            cmd.ExecuteNonQuery();

                            #endregion[부가세 미포함 시 수행되는 로직]
                        }
                        else
                        {
                            #region[부가세 포함 시 수행되는 로직]

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" -- 전표 매입 경우 ");
                            #region mariaDB 
                            //strSql.AppendLine(" INSERT INTO ACTRAN ");
                            //strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO, ACCOD ");
                            //strSql.AppendLine("           , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT ");
                            //strSql.AppendLine("           , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER ");
                            //strSql.AppendLine("           , REF1, REF2, REF3, RK, CUSER, CDATE ) ");
                            //strSql.AppendLine(" SELECT DATE_FORMAT(B.J_DATE, '%Y%m%d') AS TDATE ");
                            //strSql.AppendLine(" 	 , '3' AS ATGUB  ");
                            //strSql.AppendLine(" 	 , (  ");
                            //strSql.AppendLine("          SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE LPAD(CAST((CAST(MAX(X1.SEQNO) AS INT) + 1) AS CHAR), 4, '0') END AS SEQNO ");
                            //strSql.AppendLine("            FROM ACTRAN X1 ");
                            //strSql.AppendLine("           WHERE TDATE = DATE_FORMAT(B.J_DATE, '%Y%m%d') ");
                            //strSql.AppendLine("             AND ATGUB = '3' -- HARDCODING, '대체' ");
                            //strSql.AppendLine("             AND SEQNO >= 5000 ");
                            //strSql.AppendLine("        ) AS SEQNO ");
                            //strSql.AppendLine("      , A.SEQNO AS LINNO ");
                            //strSql.AppendLine("      , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD ");
                            //strSql.AppendLine("      , C.ACNAM ");
                            //strSql.AppendLine("      , B.J_ID1 AS CVCOD ");
                            //strSql.AppendLine("      , D.DEALER_NM AS CVNAM ");
                            //strSql.AppendLine("      , CONCAT(D.DEALER_NM, ',', E.GUBUN1 , ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT ");
                            //strSql.AppendLine("      , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP + B.BUGASE END AS ADAMT -- 주의  ");
                            //strSql.AppendLine("      , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP  ");
                            //strSql.AppendLine("             WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO = 1 THEN B.BUGASE END AS ACAMT -- 주의 ");
                            //strSql.AppendLine("      , F2.DEPT_CD AS ADPCD ");
                            //strSql.AppendLine("      , G.DEPT_NM AS ADPNM ");
                            //strSql.AppendLine("      , @AAUTO AS AAUTO ");
                            //strSql.AppendLine("      , 'Y' AS APVYN ");
                            //strSql.AppendLine("      , NOW() AS ADATE ");
                            //strSql.AppendLine("      , F2.EMP_ID AS AUSER ");
                            //strSql.AppendLine("      , B.JUNPYOID AS REF1 ");
                            //strSql.AppendLine("      , B.J_ID AS REF2 ");
                            //strSql.AppendLine("      , E.JUNPYOID AS REF3 ");
                            //strSql.AppendLine("      , '' AS RK ");
                            //strSql.AppendLine("      , F2.EMP_ID AS CUSER ");
                            //strSql.AppendLine("      , NOW() AS CDATE ");
                            //strSql.AppendLine("   FROM ACBUNF A ");
                            //strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
                            //strSql.AppendLine("     ON B.JUNPYOID = @JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF C ");
                            //strSql.AppendLine("     ON A.CACOD = C.ACCOD ");
                            //strSql.AppendLine("     OR A.DACOD = C.ACCOD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
                            //strSql.AppendLine("     ON B.J_ID1 = D.DEALER_CD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN MESURING E ");
                            //strSql.AppendLine("     ON B.J_RID = E.JUNPYOID ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST F1 ");
                            //strSql.AppendLine("     ON F1.USRCD = @USRCD ");
                            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS F2 ");
                            //strSql.AppendLine("     ON F2.EMP_ID = F1.INSANO ");
                            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD G ");
                            //strSql.AppendLine("     ON F2.DEPT_CD = G.DEPT_CD ");
                            //strSql.AppendLine("  WHERE A.BUNCD = '21' "); //SEAKPOHAM(부가세여부 Y 일 경우 미포함으로 12 수행)
                            #endregion

                            strSql.AppendLine("INSERT INTO ACTRAN                                          ");
                            strSql.AppendLine("          ( TDATE, ATGUB, SEQNO, LINNO, ACCOD               ");
                            strSql.AppendLine("          , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT               ");
                            strSql.AppendLine("          , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER ");
                            strSql.AppendLine("          , REF1, REF2, REF3, RK, CUSER, CDATE )            ");
                            strSql.AppendLine("SELECT CONVERT(VARCHAR(8),CONVERT(DATE,B.J_DATE),112) AS TDATE            ");
                            strSql.AppendLine("	 , '3' AS ATGUB                                            ");
                            strSql.AppendLine("	 , (                                                       ");
                            strSql.AppendLine("         SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE REPLICATE(0,4-LEN(MAX(X1.SEQNO))) + (MAX(X1.SEQNO)+1) END AS SEQNO ");
                            strSql.AppendLine("           FROM ACTRAN X1                                ");
                            strSql.AppendLine("          WHERE TDATE = CONVERT(VARCHAR(8),CONVERT(DATE,B.J_DATE),112) ");
                            strSql.AppendLine("            AND ATGUB = '3' -- HARDCODING, '대체'        ");
                            strSql.AppendLine("            AND SEQNO >= 5000                            ");
                            strSql.AppendLine("       ) AS SEQNO                                        ");
                            strSql.AppendLine("     , A.SEQNO AS LINNO                                  ");
                            strSql.AppendLine("     , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD ");
                            strSql.AppendLine("     , C.ACNAM              ");
                            strSql.AppendLine("     , B.J_ID1 AS CVCOD     ");
                            strSql.AppendLine("     , D.DEALER_NM AS CVNAM ");
                            strSql.AppendLine("     , CONCAT(D.DEALER_NM, ',', E.GUBUN1 , ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG) AS ATEXT ");
                            strSql.AppendLine("     , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP + B.BUGASE END AS ADAMT -- 주의 ");
                            strSql.AppendLine("     , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP  ");
                            strSql.AppendLine("            WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO = 1 THEN B.BUGASE END AS ACAMT -- 주의 ");
                            strSql.AppendLine("     , F2.DEPT_CD AS ADPCD ");
                            strSql.AppendLine("     , G.DEPT_NM AS ADPNM          ");
                            strSql.AppendLine("     , @AAUTO AS AAUTO              ");
                            strSql.AppendLine("     , 'Y' AS APVYN                ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS ADATE          ");
                            strSql.AppendLine("     , F2.EMP_ID AS AUSER          ");
                            strSql.AppendLine("     , B.JUNPYOID AS REF1          ");
                            strSql.AppendLine("     , CONVERT(VARCHAR,CONVERT(DECIMAL, B.J_ID)) AS REF2           ");
                            strSql.AppendLine("     , E.JUNPYOID AS REF3          ");
                            strSql.AppendLine("     , '' AS RK                    ");
                            strSql.AppendLine("     , F2.EMP_ID AS CUSER          ");
                            strSql.AppendLine("     , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE          ");
                            strSql.AppendLine("  FROM ACBUNF A                    ");
                            strSql.AppendLine("  LEFT OUTER JOIN INLIST B         ");
                            strSql.AppendLine("    ON B.JUNPYOID = @JUNPYOID              ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACMSTF C         ");
                            strSql.AppendLine("    ON A.CACOD = C.ACCOD           ");
                            strSql.AppendLine("    OR A.DACOD = C.ACCOD           ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D  ");
                            strSql.AppendLine("    ON B.J_ID1 = D.DEALER_CD       ");
                            strSql.AppendLine("  LEFT OUTER JOIN MESURING E       ");
                            strSql.AppendLine("    ON B.J_RID = E.JUNPYOID        ");
                            strSql.AppendLine("  LEFT OUTER JOIN ZUSRLST F1       ");
                            strSql.AppendLine("    ON F1.USRCD = @USRCD                ");
                            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS F2  ");
                            strSql.AppendLine("    ON F2.EMP_ID = F1.INSANO       ");
                            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEPT_CD G    ");
                            strSql.AppendLine("    ON F2.DEPT_CD = G.DEPT_CD      ");
                            strSql.AppendLine(" WHERE A.BUNCD = '21'              ");

                            cmd.Parameters.Clear();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.Parameters.AddWithValue("@AAUTO", "A01");
                            cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                            cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                            cmd.ExecuteNonQuery();

                            #endregion[부가세 포함 시 수행되는 로직]
                        }
                    }
                    #endregion

                    string sRmk = string.Format("[마감승인]JUNPYOID : {0}, 거래처명 : {1}, 차량 : {2}, 인수량 : {3} ", sJunpyoId, row["DEALER_NM"], row["J_BNUM"], row["ACCEPTWEIGHT"]);

                    //strSql.Clear();
                    //strSql.AppendLine(" CALL DP_IST_LOG2 (NOW()," + FmMainToolBar2.UserID + ", 'U', " + i.ToString() + " , 'INLIST', '" + sRmk + "'); ");
                    ////strSql.AppendLine(" CALL DP_IST_LOG (NOW(), " + FmMainToolBar2.UserID + ", 'U', 'INLIST', '" + sRmk + "'); ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.ExecuteNonQuery();
                }

                //for (int i = 0; i < dt.Rows.Count; i++)
                //{
                //    string sJunpyoId = dt.Rows[i]["JUNPYOID"].ToString();
                //    string sChk = dt.Rows[i]["CHK"].ToString();
                //    if (sChk.Equals("Y"))
                //    {
                //        Cursor = Cursors.WaitCursor;

                //        strSql.Clear();
                //        strSql.AppendLine(" ");
                //        strSql.AppendLine(" UPDATE INLIST ");
                //        strSql.AppendLine("    SET APRV_YN = 'Y' ");
                //        strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

                //        cmd.CommandType = CommandType.Text;
                //        cmd.CommandText = strSql.ToString();
                //        cmd.ExecuteNonQuery();

                //        Cursor = Cursors.Default;
                //    }
                //}

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                BtnRetr_Click(null, null);
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        ///     직송 관련 전표생성
        /// </summary>
        /// <param name="selectedRows">GridView RowHandle(index)</param>
        /// <param name="dt">GridControl</param>
        private void IssueingSlipOfDirectSend(int[] selectedRows, DataTable dt)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selectedRows.Length; i++)
                {
                    DataRow row = dt.NewRow();

                    if (GridBuyer.FocusedView == GridViewBuyerDtSend)
                        row = GridViewBuyerDtSend.GetDataRow(selectedRows[i]);
                    
                    string sMesure_JunpyoID = row["JUNPYOID_MESURING"].ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT JUNPYOID, KERATYPE ");
                    strSql.AppendLine("   FROM INLIST ");
                    strSql.AppendLine("  WHERE J_RID = " + sMesure_JunpyoID + " ");

                    DataTable dtInlist = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    string sIn_JunpyoId = string.Empty;
                    string sOut_JunpyoId = string.Empty;
                    foreach (DataRow drRow in dtInlist.Rows)
                    {
                        string sKeraType = drRow["KERATYPE"]?.ToString();
                        if (sKeraType.Equals("매입"))
                            sIn_JunpyoId = drRow["JUNPYOID"]?.ToString();
                        else
                            sOut_JunpyoId = drRow["JUNPYOID"]?.ToString();

                        //INLIST 승인 Y 세팅
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = NULL ");
                        strSql.AppendLine("      , APRV_DATE = NULL ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", drRow["JUNPYOID"]);
                        cmd.ExecuteNonQuery();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = 'Y' ");
                        strSql.AppendLine("      , APRV_DATE = CONVERT(VARCHAR(10),GETDATE(),23) ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", drRow["JUNPYOID"]);
                        cmd.ExecuteNonQuery();
                        
                    }

                    strSql.Clear();
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO ACTRAN ");
                    //strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO, ACCOD ");
                    //strSql.AppendLine("           , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT ");
                    //strSql.AppendLine("           , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER ");
                    //strSql.AppendLine("           , REF1, REF2, REF3, RK, CUSER, CDATE ) ");
                    //strSql.AppendLine(" SELECT TDATE, ATGUB, SEQNO, (@ROWNUM := @ROWNUM + 1) AS LINNO , ACCOD ");
                    //strSql.AppendLine("      , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT ");
                    //strSql.AppendLine("      , ACAMT, ADPCD, ADPNM, AAUTO, APVYN ");
                    //strSql.AppendLine("      , ADATE, AUSER, REF1, REF2, REF3 ");
                    //strSql.AppendLine("      , RK, CUSER, CDATE  ");
                    //strSql.AppendLine("   FROM ( ");
                    //strSql.AppendLine("          SELECT DATE_FORMAT(B.J_DATE, '%Y%m%d') AS TDATE  ");
                    //strSql.AppendLine("          	  , '3' AS ATGUB   ");
                    //strSql.AppendLine("          	  , (   ");
                    //strSql.AppendLine("                    SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE LPAD(CAST((CAST(MAX(X1.SEQNO) AS INT) + 1) AS CHAR), 4, '0') END AS SEQNO  ");
                    //strSql.AppendLine("                      FROM ACTRAN X1  ");
                    //strSql.AppendLine("                     WHERE TDATE = DATE_FORMAT(B.J_DATE, '%Y%m%d')  ");
                    //strSql.AppendLine("                       AND ATGUB = '3' -- HARDCODING, '대체'  ");
                    //strSql.AppendLine("                  ) AS SEQNO  ");
                    //strSql.AppendLine("               , A.SEQNO AS LINNO  ");
                    //strSql.AppendLine("               , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD  ");
                    //strSql.AppendLine("               , C.ACNAM  ");
                    //strSql.AppendLine("               , B.J_ID1 AS CVCOD  ");
                    //strSql.AppendLine("               , D.DEALER_NM AS CVNAM  ");
                    //strSql.AppendLine("               , CONCAT(D.DEALER_NM, ', ', E.GUBUN1 , ', ', E.J_BNUM, ', ', B.DANGA, ', ', B.DANJUNG, ', ', '[직송]') AS ATEXT  ");
                    //strSql.AppendLine("               , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP  ");
                    //strSql.AppendLine("                      WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO = 2 THEN B.BUGASE END AS ADAMT   ");
                    //strSql.AppendLine("               , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP + B.BUGASE END AS ACAMT  ");
                    //strSql.AppendLine("               , F2.DEPT_CD AS ADPCD  ");
                    //strSql.AppendLine("               , G.DEPT_NM AS ADPNM  ");
                    //strSql.AppendLine("               , @AAUTO AS AAUTO  ");
                    //strSql.AppendLine("               , 'Y' AS APVYN  ");
                    //strSql.AppendLine("               , NOW() AS ADATE  ");
                    //strSql.AppendLine("               , F2.EMP_ID AS AUSER  ");
                    //strSql.AppendLine("               , B.JUNPYOID AS REF1   ");
                    //strSql.AppendLine("               , B.J_ID AS REF2  ");
                    //strSql.AppendLine("               , E.JUNPYOID AS REF3   ");
                    //strSql.AppendLine("               , '' AS RK  ");
                    //strSql.AppendLine("               , F2.EMP_ID AS CUSER  ");
                    //strSql.AppendLine("               , NOW() AS CDATE  ");
                    //strSql.AppendLine("           FROM ACBUNF2 A  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN INLIST B  ");
                    //strSql.AppendLine("             ON B.JUNPYOID = @OUT_JUNPYOID ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF C  ");
                    //strSql.AppendLine("             ON A.CACOD = C.ACCOD  ");
                    //strSql.AppendLine("             OR A.DACOD = C.ACCOD  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ACC_DEALER_CD D  ");
                    //strSql.AppendLine("             ON B.J_ID1 = D.DEALER_CD  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN MESURING E  ");
                    //strSql.AppendLine("             ON B.J_RID = E.JUNPYOID  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ZUSRLST F1  ");
                    //strSql.AppendLine("             ON F1.USRCD = @USRCD ");
                    //strSql.AppendLine("           LEFT OUTER JOIN HR_EMP_BASIS F2  ");
                    //strSql.AppendLine("             ON F2.EMP_ID = F1.INSANO  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ACC_DEPT_CD G  ");
                    //strSql.AppendLine("             ON F2.DEPT_CD = G.DEPT_CD  ");
                    //strSql.AppendLine("          WHERE A.BUNCD = '11'  ");
                    //strSql.AppendLine("          UNION ALL  ");
                    //strSql.AppendLine("         SELECT DATE_FORMAT(B.J_DATE, '%Y%m%d') AS TDATE  ");
                    //strSql.AppendLine("   	          , '3' AS ATGUB   ");
                    //strSql.AppendLine("   	          , (   ");
                    //strSql.AppendLine("                     SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE LPAD(CAST((CAST(MAX(X1.SEQNO) AS INT) + 1) AS CHAR), 4, '0') END AS SEQNO  ");
                    //strSql.AppendLine("                       FROM ACTRAN X1  ");
                    //strSql.AppendLine("                      WHERE TDATE = DATE_FORMAT(B.J_DATE, '%Y%m%d')  ");
                    //strSql.AppendLine("                        AND ATGUB = '3' -- HARDCODING, '대체'  ");
                    //strSql.AppendLine("                   ) AS SEQNO  ");
                    //strSql.AppendLine("              , A.SEQNO AS LINNO  ");
                    //strSql.AppendLine("              , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD  ");
                    //strSql.AppendLine("              , C.ACNAM  ");
                    //strSql.AppendLine("              , B.J_ID1 AS CVCOD  ");
                    //strSql.AppendLine("              , D.DEALER_NM AS CVNAM  ");
                    //strSql.AppendLine("              , CONCAT(D.DEALER_NM, ',', E.GUBUN1 , ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG, ', [직송]') AS ATEXT  ");
                    //strSql.AppendLine("              , CASE WHEN (A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP + B.BUGASE END AS ADAMT -- 주의   ");
                    //strSql.AppendLine("              , CASE WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP   ");
                    //strSql.AppendLine("                     WHEN (A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO = 1 THEN B.BUGASE END AS ACAMT -- 주의  ");
                    //strSql.AppendLine("              , F2.DEPT_CD AS ADPCD  ");
                    //strSql.AppendLine("              , G.DEPT_NM AS ADPNM  ");
                    //strSql.AppendLine("              , @AAUTO AS AAUTO  ");
                    //strSql.AppendLine("              , 'Y' AS APVYN  ");
                    //strSql.AppendLine("              , NOW() AS ADATE  ");
                    //strSql.AppendLine("              , F2.EMP_ID AS AUSER  ");
                    //strSql.AppendLine("              , B.JUNPYOID AS REF1  ");
                    //strSql.AppendLine("              , B.J_ID AS REF2  ");
                    //strSql.AppendLine("              , E.JUNPYOID AS REF3  ");
                    //strSql.AppendLine("              , '' AS RK  ");
                    //strSql.AppendLine("              , F2.EMP_ID AS CUSER  ");
                    //strSql.AppendLine("              , NOW() AS CDATE  ");
                    //strSql.AppendLine("           FROM ACBUNF2 A  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN INLIST B  ");
                    //strSql.AppendLine("             ON B.JUNPYOID = @IN_JUNPYOID ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ACMSTF C  ");
                    //strSql.AppendLine("             ON A.CACOD = C.ACCOD  ");
                    //strSql.AppendLine("             OR A.DACOD = C.ACCOD  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ACC_DEALER_CD D  ");
                    //strSql.AppendLine("             ON B.J_ID1 = D.DEALER_CD  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN MESURING E  ");
                    //strSql.AppendLine("             ON B.J_RID = E.JUNPYOID  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ZUSRLST F1  ");
                    //strSql.AppendLine("             ON F1.USRCD = @USRCD ");
                    //strSql.AppendLine("           LEFT OUTER JOIN HR_EMP_BASIS F2  ");
                    //strSql.AppendLine("             ON F2.EMP_ID = F1.INSANO  ");
                    //strSql.AppendLine("           LEFT OUTER JOIN ACC_DEPT_CD G  ");
                    //strSql.AppendLine("             ON F2.DEPT_CD = G.DEPT_CD  ");
                    //strSql.AppendLine("          WHERE A.BUNCD = '21'  ");
                    //strSql.AppendLine("       ) Y, (SELECT @ROWNUM := - 1) AS R ");
                    #endregion                                                                                                  ");
                                                                                                                             
                    strSql.AppendLine("INSERT INTO ACTRAN                                                                    ");
                    strSql.AppendLine("          (TDATE, ATGUB, SEQNO, LINNO, ACCOD                                          ");
                    strSql.AppendLine("          , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT                                         ");
                    strSql.AppendLine("          , ACAMT, ADPCD, ADPNM, AAUTO, APVYN, ADATE, AUSER                           ");
                    strSql.AppendLine("          , REF1, REF2, REF3, RK, CUSER, CDATE)                                       ");
                    strSql.AppendLine("SELECT TDATE, ATGUB, SEQNO, row_number() over(order by BUNCD, LINNO) AS LINNO, ACCOD  ");
                    strSql.AppendLine("     , ACNAM, CVCOD, CVNAM, ATEXT, ADAMT                                              ");
                    strSql.AppendLine("     , ACAMT, ADPCD, ADPNM, AAUTO, APVYN                                              ");
                    strSql.AppendLine("     , ADATE, AUSER, REF1, REF2, REF3                                                 ");
                    strSql.AppendLine("     , RK, CUSER, CDATE                                                               ");
                    strSql.AppendLine("  FROM(                                                                               ");
                    strSql.AppendLine("         SELECT CONVERT(VARCHAR(8), CONVERT(DATE,B.J_DATE), 112) AS TDATE                           ");
                    strSql.AppendLine("               , '3' AS ATGUB                                                         ");
                    strSql.AppendLine("               , (                                                                    ");
                    strSql.AppendLine("                   SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE REPLICATE(0, 4 - LEN(MAX(X1.SEQNO))) + (MAX(X1.SEQNO)+1) END AS SEQNO");
                    strSql.AppendLine("                     FROM ACTRAN X1                                 ");
                    strSql.AppendLine("                    WHERE TDATE = CONVERT(VARCHAR(8), CONVERT(DATE,B.J_DATE), 112)");
                    strSql.AppendLine("                      AND ATGUB = '3'-- HARDCODING, '대체'          ");
                    strSql.AppendLine("                 ) AS SEQNO                                         ");
                    strSql.AppendLine("              , A.SEQNO AS LINNO                                    ");
                    strSql.AppendLine("              , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD");
                    strSql.AppendLine("              , C.ACNAM              ");
                    strSql.AppendLine("              , B.J_ID1 AS CVCOD     ");
                    strSql.AppendLine("              , D.DEALER_NM AS CVNAM ");
                    strSql.AppendLine("              , CONCAT(D.DEALER_NM, ', ', E.GUBUN1, ', ', E.J_BNUM, ', ', B.DANGA, ', ', B.DANJUNG, ', ', '[직송]') AS ATEXT");
                    strSql.AppendLine("             , CASE WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP                                   ");
                    strSql.AppendLine("                     WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO = 2 THEN B.BUGASE END AS ADAMT                       ");
                    strSql.AppendLine("             , CASE WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 2 THEN B.KONGKEP + B.BUGASE END AS ACAMT           ");
                    strSql.AppendLine("              , F2.DEPT_CD AS ADPCD        ");
                    strSql.AppendLine("              , G.DEPT_NM AS ADPNM         ");
                    strSql.AppendLine("              , @AAUTO AS AAUTO             ");
                    strSql.AppendLine("              , 'Y' AS APVYN               ");
                    strSql.AppendLine("              , CONVERT(VARCHAR(20),GETDATE(),20) AS ADATE         ");
                    strSql.AppendLine("              , F2.EMP_ID AS AUSER         ");
                    strSql.AppendLine("              , B.JUNPYOID AS REF1         ");
                    strSql.AppendLine("              , CONVERT(VARCHAR,CONVERT(DECIMAL, B.J_ID)) AS REF2             ");
                    strSql.AppendLine("              , E.JUNPYOID AS REF3         ");
                    strSql.AppendLine("              , '' AS RK                   ");
                    strSql.AppendLine("              , F2.EMP_ID AS CUSER         ");
                    strSql.AppendLine("              , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE         ");
                    strSql.AppendLine("              , A.BUNCD                    ");
                    strSql.AppendLine("          FROM ACBUNF2 A                   ");
                    strSql.AppendLine("          LEFT OUTER JOIN INLIST B         ");
                    strSql.AppendLine("            ON B.JUNPYOID = @OUT_JUNPYOID              ");
                    strSql.AppendLine("          LEFT OUTER JOIN ACMSTF C         ");
                    strSql.AppendLine("            ON A.CACOD = C.ACCOD           ");
                    strSql.AppendLine("            OR A.DACOD = C.ACCOD           ");
                    strSql.AppendLine("          LEFT OUTER JOIN ACC_DEALER_CD D  ");
                    strSql.AppendLine("            ON B.J_ID1 = D.DEALER_CD       ");
                    strSql.AppendLine("          LEFT OUTER JOIN MESURING E       ");
                    strSql.AppendLine("            ON B.J_RID = E.JUNPYOID        ");
                    strSql.AppendLine("          LEFT OUTER JOIN ZUSRLST F1       ");
                    strSql.AppendLine("            ON F1.USRCD = @USRCD               ");
                    strSql.AppendLine("          LEFT OUTER JOIN HR_EMP_BASIS F2  ");
                    strSql.AppendLine("            ON F2.EMP_ID = F1.INSANO       ");
                    strSql.AppendLine("          LEFT OUTER JOIN ACC_DEPT_CD G    ");
                    strSql.AppendLine("            ON F2.DEPT_CD = G.DEPT_CD      ");
                    strSql.AppendLine("         WHERE A.BUNCD = '11'              ");
                    strSql.AppendLine("         UNION ALL                         ");
                    strSql.AppendLine("        SELECT CONVERT(VARCHAR(8),CONVERT(DATE,B.J_DATE),112) AS TDATE");
                    strSql.AppendLine("                , '3' AS ATGUB  ");
                    strSql.AppendLine("                , (             ");
                    strSql.AppendLine("                    SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '5000' ELSE REPLICATE(0, 4 - LEN(MAX(X1.SEQNO))) + (MAX(X1.SEQNO)+1) END AS SEQNO");
                    strSql.AppendLine("                      FROM ACTRAN X1                                 ");
                    strSql.AppendLine("                     WHERE TDATE = CONVERT(VARCHAR(8), CONVERT(DATE,B.J_DATE), 112)");
                    strSql.AppendLine("                       AND ATGUB = '3'-- HARDCODING, '대체'          ");
                    strSql.AppendLine("                  ) AS SEQNO                                         ");
                    strSql.AppendLine("             , A.SEQNO AS LINNO                                      ");
                    strSql.AppendLine("             , CASE WHEN A.CACOD IS NULL OR A.CACOD = '' THEN A.DACOD ELSE A.CACOD END AS ACCOD");
                    strSql.AppendLine("             , C.ACNAM              ");
                    strSql.AppendLine("             , B.J_ID1 AS CVCOD     ");
                    strSql.AppendLine("             , D.DEALER_NM AS CVNAM ");
                    strSql.AppendLine("             , CONCAT(D.DEALER_NM, ',', E.GUBUN1, ',', E.J_BNUM, ',', B.DANGA, ',', B.DANJUNG, ', [직송]') AS ATEXT      ");
                    strSql.AppendLine("             , CASE WHEN(A.CACOD IS NULL OR A.CACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP + B.BUGASE END AS ADAMT-- 주의");
                    strSql.AppendLine("             , CASE WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO <> 1 THEN B.IKONGKEP                               ");
                    strSql.AppendLine("                    WHEN(A.DACOD IS NULL OR A.DACOD = '') AND A.SEQNO = 1 THEN B.BUGASE END AS ACAMT --주의              ");
                    strSql.AppendLine("             , F2.DEPT_CD AS ADPCD         ");
                    strSql.AppendLine("             , G.DEPT_NM AS ADPNM          ");
                    strSql.AppendLine("             , @AAUTO AS AAUTO              ");
                    strSql.AppendLine("             , 'Y' AS APVYN                ");
                    strSql.AppendLine("             , CONVERT(VARCHAR(20),GETDATE(),20) AS ADATE          ");
                    strSql.AppendLine("             , F2.EMP_ID AS AUSER          ");
                    strSql.AppendLine("             , B.JUNPYOID AS REF1          ");
                    strSql.AppendLine("             , CONVERT(VARCHAR,CONVERT(DECIMAL, B.J_ID)) AS REF2             ");
                    strSql.AppendLine("             , E.JUNPYOID AS REF3          ");
                    strSql.AppendLine("             , '' AS RK                    ");
                    strSql.AppendLine("             , F2.EMP_ID AS CUSER          ");
                    strSql.AppendLine("             , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE          ");
                    strSql.AppendLine("             , A.BUNCD                     ");
                    strSql.AppendLine("          FROM ACBUNF2 A                   ");
                    strSql.AppendLine("          LEFT OUTER JOIN INLIST B         ");
                    strSql.AppendLine("            ON B.JUNPYOID = @IN_JUNPYOID              ");
                    strSql.AppendLine("          LEFT OUTER JOIN ACMSTF C         ");
                    strSql.AppendLine("            ON A.CACOD = C.ACCOD           ");
                    strSql.AppendLine("            OR A.DACOD = C.ACCOD           ");
                    strSql.AppendLine("          LEFT OUTER JOIN ACC_DEALER_CD D  ");
                    strSql.AppendLine("            ON B.J_ID1 = D.DEALER_CD       ");
                    strSql.AppendLine("          LEFT OUTER JOIN MESURING E       ");
                    strSql.AppendLine("            ON B.J_RID = E.JUNPYOID        ");
                    strSql.AppendLine("          LEFT OUTER JOIN ZUSRLST F1       ");
                    strSql.AppendLine("            ON F1.USRCD = @USRCD               ");
                    strSql.AppendLine("          LEFT OUTER JOIN HR_EMP_BASIS F2  ");
                    strSql.AppendLine("            ON F2.EMP_ID = F1.INSANO       ");
                    strSql.AppendLine("          LEFT OUTER JOIN ACC_DEPT_CD G    ");
                    strSql.AppendLine("            ON F2.DEPT_CD = G.DEPT_CD      ");
                    strSql.AppendLine("         WHERE A.BUNCD = '21'              ");
                    strSql.AppendLine("      ) Y                                  ");

                    cmd.Parameters.Clear();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.AddWithValue("@AAUTO", "A02");
                    cmd.Parameters.AddWithValue("@IN_JUNPYOID", Convert.ToDouble(sIn_JunpyoId));
                    cmd.Parameters.AddWithValue("@OUT_JUNPYOID", Convert.ToDouble(sOut_JunpyoId));
                    cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID); 
                    cmd.ExecuteNonQuery();

                    string sRmk = string.Format("[마감승인][직송]매입거래처명 : {0}, 매출거래처명 : {1}, 차량 : {2}, 매출중량 : {3}, JUNPYOID : {4}(매입), {5}(매출)", row["SALE_DEALER_NM"], row["PURC_DEALER_NM"], row["J_BNUM"], row["OWEIGHT"], sIn_JunpyoId, sOut_JunpyoId);

                    //strSql.Clear();
                    //strSql.AppendLine(" CALL DP_IST_LOG2 (NOW(), " + FmMainToolBar2.UserID + ", 'U', " + i.ToString() + " , 'INLIST', '" + sRmk + "'); ");
                    ////strSql.AppendLine(" CALL DP_IST_LOG (NOW(), " + FmMainToolBar2.UserID + ", 'U', 'INLIST', '" + sRmk + "'); ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                BtnRetr_Click(null, null);
                Cursor = Cursors.Default;
            }
        }

        private void BtnAprvCancle_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (CheckApproval())
                return;

            DataTable dt = (DataTable)GridBuyer.DataSource;

            int[] selectedRows = GridViewBuyerPurc.GetSelectedRows();

            if (GridBuyer.FocusedView == GridViewBuyerPurc)
                selectedRows = GridViewBuyerPurc.GetSelectedRows();
            else if (GridBuyer.FocusedView == GridViewBuyerSales)
                selectedRows = GridViewBuyerSales.GetSelectedRows();
            else if (GridBuyer.FocusedView == GridViewBuyerDtSend)
                selectedRows = GridViewBuyerDtSend.GetSelectedRows();

            if (selectedRows.Length == 0)
            {
                XtraMessageBox.Show("승인하려는 데이터에 체크하세요.");
                return;
            }
            
            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("조회내역이 존재하지 않습니다.");
                DateEditFrom.Focus();
                DateEditFrom.SelectAll();
                return;
            }

            if (XtraMessageBox.Show("선택된 데이터와 관련된 전표내역까지 삭제됩니다.\r\n진행하시겠습니까?", "취소여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            if (GridBuyer.FocusedView == GridViewBuyerPurc)
            {
                CancleIssuedSlipOfPurcAndSale(GridViewBuyerPurc.GetSelectedRows(), dt);
            }
            else if (GridBuyer.FocusedView == GridViewBuyerSales)
            {
                CancleIssuedSlipOfPurcAndSale(GridViewBuyerSales.GetSelectedRows(), dt);
            }
            else if (GridBuyer.FocusedView == GridViewBuyerDtSend)
            {
                CancelIuusedSlipOfDirectSend(GridViewBuyerDtSend.GetSelectedRows(), dt);
            }
        }

        /// <summary>
        ///  매입출 관련 마감승인취소
        /// </summary>
        /// <param name="selectedRows">GridView RowHandle(Index)</param>
        /// <param name="dt">GridControl.DataSource</param>
        private void CancleIssuedSlipOfPurcAndSale(int[] selectedRows, DataTable dt)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selectedRows.Length; i++)
                {
                    DataRow row = dt.NewRow();

                    if (GridBuyer.FocusedView == GridViewBuyerPurc)
                        row = GridViewBuyerPurc.GetDataRow(selectedRows[i]);
                    else if (GridBuyer.FocusedView == GridViewBuyerSales)
                        row = GridViewBuyerSales.GetDataRow(selectedRows[i]);

                    string sJunpyoId = row["JUNPYOID"].ToString();
                    double dJunpyoId = Convert.ToDouble(sJunpyoId);
                    Cursor = Cursors.WaitCursor;

                    //APRV_YN DB 상 찾지못한 에러로 쿼리 두번날려 수정하도록 함(2020-06-17)
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE INLIST ");
                    strSql.AppendLine("    SET APRV_YN = NULL ");
                    strSql.AppendLine("      , APRV_DATE = NULL ");
                    strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", dJunpyoId);
                    cmd.ExecuteNonQuery();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE INLIST ");
                    strSql.AppendLine("    SET APRV_YN = 'N' ");
                    strSql.AppendLine("      , APRV_DATE = NULL ");
                    strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", dJunpyoId);
                    cmd.ExecuteNonQuery();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT CONCAT('TABLE : ACTRAN / TDATE : ', TDATE, ', ATGUB : ', ATGUB, ', SEQNO : ', SEQNO) AS REF_RMK ");
                    strSql.AppendLine("   FROM ACTRAN ");
                    strSql.AppendLine("  WHERE REF1 = '" + sJunpyoId + "'" );
                    strSql.AppendLine("    AND AAUTO = 'A01' ");

                    string sREF_RMK = string.Empty;
                    DataTable dtLog = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    if(dtLog.Rows.Count > 0)
                    {
                        sREF_RMK = dtLog.Rows[0]["REF_RMK"]?.ToString();
                    }

                    strSql.Clear();
                    strSql.AppendLine(" DELETE                ");
                    strSql.AppendLine("   FROM ACTRAN         ");
                    strSql.AppendLine("  WHERE REF1 = @JUNPYOID");
                    strSql.AppendLine("    AND AAUTO = 'A01' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", dJunpyoId.ToString());
                    cmd.ExecuteNonQuery();
                    
                    /* 
                     * #0001
                     */
                    string sRmk = "[마감승인취소]";
                    string sSTD_COLS = string.Format("{0}/{1}/{2}/차량:{3}/{4}/인수량:{5}", row["J_DATE"]?.ToString().Substring(0, 10), row["KERATYPE"], row["DEALER_NM"], row["J_BNUM"], row["GUBUN1"], row["ACCEPTWEIGHT"]);

                    strSql.Clear();
                    #region mariaDB
                    //strSql.AppendLine("   INSERT INTO ZSYS_LOG ");
                    //strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK, STD_COLS, REF_RMK )   ");
                    //strSql.AppendLine(" 	      VALUES ");
                    //strSql.AppendLine(" 	           ( NOW() ");
                    //strSql.AppendLine(" 	           , '" + FmMainToolBar2.UserID + "' ");
                    //strSql.AppendLine(" 	           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    //strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    //strSql.AppendLine("                     WHERE X1.OCCUR_DATE = NOW() ");
                    //strSql.AppendLine("                       AND X1.USRCD = '" + FmMainToolBar2.UserID + "' ) #LOG_SEQ(구분자) ");
                    //strSql.AppendLine(" 	           , '" + this.Name + "' ");
                    //strSql.AppendLine(" 	           , 'D' ");
                    //strSql.AppendLine(" 	           , '" + ClsFunc.GetLocalIP() + "' ");
                    //strSql.AppendLine(" 	           , '" + sRmk + "' ");
                    //strSql.AppendLine(" 	           , '" + sSTD_COLS + "' ");
                    //strSql.AppendLine(" 	           , '" + sREF_RMK + "' ); ");
                    #endregion

                    strSql.AppendLine("   INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK, STD_COLS, REF_RMK )   ");
                    strSql.AppendLine(" 	      VALUES ");
                    strSql.AppendLine(" 	           ( CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine(" 	           , '" + FmMainToolBar2.UserID + "' ");
                    strSql.AppendLine(" 	           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    strSql.AppendLine("                     WHERE X1.OCCUR_DATE = CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine("                       AND X1.USRCD = '" + FmMainToolBar2.UserID + "' ) --LOG_SEQ(구분자) ");
                    strSql.AppendLine(" 	           , '" + this.Name + "' ");
                    strSql.AppendLine(" 	           , 'D' ");
                    strSql.AppendLine(" 	           , '" + ClsFunc.GetLocalIP() + "' ");
                    strSql.AppendLine(" 	           , '" + sRmk + "' ");
                    strSql.AppendLine(" 	           , '" + sSTD_COLS + "' ");
                    strSql.AppendLine(" 	           , '" + sREF_RMK + "' ); ");

                    cmd.Parameters.Clear();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                    //cmd.Parameters.AddWithValue("@PGM_ID", "INLIST");
                    //cmd.Parameters.AddWithValue("@EDIT_KIND", "D");
                    //cmd.Parameters.AddWithValue("@IP", ClsFunc.GetLocalIP());
                    //cmd.Parameters.AddWithValue("@EDIT_RMK", sRmk);
                    cmd.Parameters.Clear();
                    cmd.ExecuteNonQuery();

                    //strSql.Clear();
                    //strSql.AppendLine(" CALL DP_IST_LOG2 (NOW(), " + FmMainToolBar2.UserID + ", 'D', " + i.ToString() + " , 'INLIST', '" + sRmk + "'); ");
                    ////strSql.AppendLine(" CALL DP_IST_LOG (NOW(), " + FmMainToolBar2.UserID + ", 'D', 'INLIST', '" + sRmk + "'); ");
                    ////strSql.AppendLine(" CALL DP_IST_LOG (CAST('" + DateTime.Now + "' AS DATETIME), " + FmMainToolBar2.UserID + ", 'D', 'INLIST', '" + sRmk + "'); ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("승인취소가 완료되었습니다.");
                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        ///  직송 관련 마감승인취소
        /// </summary>
        /// <param name="selectedRows">GridView RowHandle(Index)</param>
        /// <param name="dt">GridControl.DataSource</param>
        private void CancelIuusedSlipOfDirectSend(int[] selectedRows, DataTable dt)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selectedRows.Length; i++)
                {
                    Cursor = Cursors.WaitCursor;

                    DataRow row = dt.NewRow();

                    if (GridBuyer.FocusedView == GridViewBuyerDtSend)
                        row = GridViewBuyerDtSend.GetDataRow(selectedRows[i]);

                    string sMesure_JunpyoID = row["JUNPYOID_MESURING"].ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT JUNPYOID, KERATYPE ");
                    strSql.AppendLine("   FROM INLIST ");
                    strSql.AppendLine("  WHERE J_RID = " + sMesure_JunpyoID + " ");

                    DataTable dtInlist = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    string sIn_JunpyoId = string.Empty;
                    string sOut_JunpyoId = string.Empty;
                    foreach (DataRow drRow in dtInlist.Rows)
                    {
                        string sKeraType = drRow["KERATYPE"]?.ToString();
                        if (sKeraType.Equals("매입"))
                            sIn_JunpyoId = drRow["JUNPYOID"]?.ToString();
                        else
                            sOut_JunpyoId = drRow["JUNPYOID"]?.ToString();

                        //INLIST 승인 Y 세팅
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = NULL ");
                        strSql.AppendLine("      , APRV_DATE = NULL ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", drRow["JUNPYOID"]);
                        cmd.ExecuteNonQuery();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = 'N' ");
                        strSql.AppendLine("      , APRV_DATE = CONVERT(VARCHAR(10),GETDATE(),23) ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", drRow["JUNPYOID"]);
                        cmd.ExecuteNonQuery();
                        
                    }

                    //string sJunpyoId = row["JUNPYOID"].ToString();
                    //double dJunpyoId = Convert.ToDouble(sJunpyoId);
                    //Cursor = Cursors.WaitCursor;

                    ////APRV_YN DB 상 찾지못한 에러로 쿼리 두번날려 수정하도록 함(2020-06-17)
                    //strSql.Clear();
                    //strSql.AppendLine(" ");
                    //strSql.AppendLine(" UPDATE INLIST ");
                    //strSql.AppendLine("    SET APRV_YN = NULL ");
                    //strSql.AppendLine("      , APRV_DATE = NULL ");
                    //strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //cmd.Parameters.AddWithValue("@JUNPYOID", dJunpyoId);
                    //cmd.ExecuteNonQuery();

                    //strSql.Clear();
                    //strSql.AppendLine(" ");
                    //strSql.AppendLine(" UPDATE INLIST ");
                    //strSql.AppendLine("    SET APRV_YN = 'N' ");
                    //strSql.AppendLine("      , APRV_DATE = NULL ");
                    //strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //cmd.Parameters.AddWithValue("@JUNPYOID", dJunpyoId);
                    //cmd.ExecuteNonQuery();

                    /*
                     * #0001
                     */
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT CONCAT('TABLE : ACTRAN / TDATE : ', TDATE, ', ATGUB : ', ATGUB, ', SEQNO : ', SEQNO) AS REF_RMK ");
                    strSql.AppendLine("   FROM ACTRAN ");
                    strSql.AppendLine("  WHERE REF3 = '" + sMesure_JunpyoID + "'");
                    strSql.AppendLine("    AND AAUTO = 'A02' ");

                    string sREF_RMK = string.Empty;
                    DataTable dtLog = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    if (dtLog.Rows.Count > 0)
                    {
                        sREF_RMK = dtLog.Rows[0]["REF_RMK"]?.ToString();
                    }

                    strSql.Clear();
                    strSql.AppendLine(" DELETE                ");
                    strSql.AppendLine("   FROM ACTRAN         ");
                    strSql.AppendLine("  WHERE REF3 = @JUNPYOID"); //ACTRAN, AAUTO -> A02 경우 직송으로 REF3(MesuringJunpyo)로 찾아 삭제
                    strSql.AppendLine("    AND AAUTO = 'A02' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@JUNPYOID", sMesure_JunpyoID);
                    cmd.ExecuteNonQuery();

                    string sRmk = "[마감승인취소]";
                    string sSTD_COLS = string.Format("{0}/{1}/{2}/차량:{3}/{4}/매출중량:{5}", row["J_DATE"]?.ToString().Substring(0, 10), row["KERATYPE"]?.ToString(), row["SALE_DEALER_NM"], row["J_BNUM"], row["GUBUN1"], row["OWEIGHT"]);
                    
                    strSql.Clear();
                    #region mariaDB
                    //strSql.AppendLine("   INSERT INTO ZSYS_LOG ");
                    //strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK, STD_COLS, REF_RMK )   ");
                    //strSql.AppendLine(" 	      VALUES ");
                    //strSql.AppendLine(" 	           ( NOW() ");
                    //strSql.AppendLine(" 	           , '" + FmMainToolBar2.UserID + "' ");
                    //strSql.AppendLine(" 	           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    //strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    //strSql.AppendLine("                     WHERE X1.OCCUR_DATE = NOW() ");
                    //strSql.AppendLine("                       AND X1.USRCD = '" + FmMainToolBar2.UserID + "' ) #LOG_SEQ(구분자) ");
                    //strSql.AppendLine(" 	           , '" + this.Name + "' ");
                    //strSql.AppendLine(" 	           , 'D' ");
                    //strSql.AppendLine(" 	           , '" + ClsFunc.GetLocalIP() + "' ");
                    //strSql.AppendLine(" 	           , '" + sRmk + "' ");
                    //strSql.AppendLine(" 	           , '" + sSTD_COLS + "' ");
                    //strSql.AppendLine(" 	           , '" + sREF_RMK + "' ); ");
                    #endregion

                    strSql.AppendLine("   INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK, STD_COLS, REF_RMK )   ");
                    strSql.AppendLine(" 	      VALUES ");
                    strSql.AppendLine(" 	           ( CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine(" 	           , '" + FmMainToolBar2.UserID + "' ");
                    strSql.AppendLine(" 	           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    strSql.AppendLine("                     WHERE X1.OCCUR_DATE = CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine("                       AND X1.USRCD = '" + FmMainToolBar2.UserID + "' ) --LOG_SEQ(구분자) ");
                    strSql.AppendLine(" 	           , '" + this.Name + "' ");
                    strSql.AppendLine(" 	           , 'D' ");
                    strSql.AppendLine(" 	           , '" + ClsFunc.GetLocalIP() + "' ");
                    strSql.AppendLine(" 	           , '" + sRmk + "' ");
                    strSql.AppendLine(" 	           , '" + sSTD_COLS + "' ");
                    strSql.AppendLine(" 	           , '" + sREF_RMK + "' ); ");

                    cmd.Parameters.Clear();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                    //cmd.Parameters.AddWithValue("@PGM_ID", "INLIST");
                    //cmd.Parameters.AddWithValue("@EDIT_KIND", "D");
                    //cmd.Parameters.AddWithValue("@IP", ClsFunc.GetLocalIP());
                    //cmd.Parameters.AddWithValue("@EDIT_RMK", sRmk);
                    cmd.Parameters.Clear();
                    cmd.ExecuteNonQuery();

                    Cursor = Cursors.Default;
                }

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("취소가 완료되었습니다.");
                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);

                Cursor = Cursors.Default;
            }
            
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            ComnEtcFunc.ExportExcelFile(this.Text + "_", GridBuyer);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void CboKeraType_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
            DataTable dt = GetLookUpData();
            ComGrid.SetLookUpEdit(LkupDealerNm, dt, "CD", "CD", "Y");
            LkupDealerNm.Properties.PopulateColumns();
            LkupDealerNm.Properties.Columns[1].Visible = false;
        }

        private void RdgbAprvGb_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetButtonVisible();
            BtnRetr_Click(null, null);
        }

        private void TxtDealerNm_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }
        
        private void AC11001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F3)
            {
                BtnAprvClose_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnAprvCancle_Click(null, null);
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

        #region[Query]

        private void PurchacePerBuyer(string sYmdFrom, string sYmdTo, string sDealerNm, string sAprvGb)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            #region mariaDB
            //strSql.AppendLine(" SELECT A.JUNPYOID ");
            //strSql.AppendLine("      , 'N' AS CHK ");
            //strSql.AppendLine("      , A.KERATYPE ");
            //strSql.AppendLine("      , A.J_DATE ");
            //strSql.AppendLine("      , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("      , A.APRV_DATE ");
            //strSql.AppendLine("	     , D.DEALER_NM ");
            //strSql.AppendLine("	     , E.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("	     , B.J_BNUM ");
            //strSql.AppendLine("	     , G.WEIGHT AS LANDEDWEIGHT ");
            //strSql.AppendLine("	     , A.HALIN AS LOSS ");
            //strSql.AppendLine("	     , A.IWEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine("	     , IFNULL(A.IWEIGHT, 0) - IFNULL(G.WEIGHT, 0) AS DIFFWEIGHT ");
            //strSql.AppendLine("	     , F.GUBUN1 ");
            //strSql.AppendLine("	     , A.MIDANGA AS STDDUNITPRICE ");
            //strSql.AppendLine("	     , A.DANGA AS PAYEDUNITPRICE ");
            ////strSql.AppendLine("	     , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("	     , (((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT) - IFNULL(A.MIDANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("	     , ((IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("	     , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("	     , A.IKONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("	     , (IFNULL(A.IWEIGHT, 0) * IFNULL(A.DANGA, 0)) + IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("	     , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("	     , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("	     , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("      , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("      , IFNULL(A.SEAKPOHAM, 'N') AS SEAKPOHAM "); //포함 : N, 미포함 : Y
            //strSql.AppendLine("   FROM INLIST A ");
            //strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("	    ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("     ON E.EMP_ID = A.CHRG_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("     ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("     ON G.IPCHULGO_MAIPID = A.J_ID ");
            #endregion

            strSql.AppendLine("SELECT A.JUNPYOID                                                  ");
            strSql.AppendLine("     , 'N' AS CHK                                                  ");
            strSql.AppendLine("     , A.KERATYPE                                                  ");
            strSql.AppendLine("     , A.J_DATE                                                    ");
            strSql.AppendLine("     , G.J_DATE AS MESURE_DT                                       ");
            strSql.AppendLine("     , A.APRV_DATE                                                 ");
            strSql.AppendLine("     , D.DEALER_NM                                                 ");
            strSql.AppendLine("     , E.EMP_NM AS INSPECTOR                                       ");
            strSql.AppendLine("     , B.J_BNUM                                                    ");
            strSql.AppendLine("     , G.WEIGHT AS LANDEDWEIGHT                                    ");
            strSql.AppendLine("     , A.HALIN AS LOSS                                             ");
            strSql.AppendLine("     , A.IWEIGHT AS ACCEPTWEIGHT                                   ");
            strSql.AppendLine("     , ISNULL(A.IWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT    ");
            strSql.AppendLine("     , F.GUBUN1                                                    ");
            strSql.AppendLine("     , A.MIDANGA AS STDDUNITPRICE                                  ");
            strSql.AppendLine("     , A.DANGA AS PAYEDUNITPRICE                                   ");
            //strSql.AppendLine("     , ISNULL(A.MIDANGA, 0) - ISNULL(A.DANGA, 0) AS DIFFUNITPRICE  ");
            strSql.AppendLine("     , (((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0)) / A.IWEIGHT) -ISNULL(A.MIDANGA, 0) AS DIFFUNITPRICE");
            strSql.AppendLine("    , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0)) / A.IWEIGHT AS ARRVUNITPRICE");
            strSql.AppendLine("    , ISNULL(A.CKONGKEP, 0) AS CARRYCOST  ");
            strSql.AppendLine("    , A.IKONGKEP AS TOTALPRICE            ");
            strSql.AppendLine("     , (ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE");
            strSql.AppendLine("     , G.J_STATE AS LOSSCAUSE       ");
            strSql.AppendLine("     , G.JUNPYOID AS IMAGE          ");
            strSql.AppendLine("     , G.GUMSUBIGO AS INSPECTNOTE   ");
            strSql.AppendLine("     , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG), 1, 1) AS CARRY_UNIT_PRICE ");
            strSql.AppendLine("       , ISNULL(A.SEAKPOHAM, 'N') AS SEAKPOHAM                              ");
            strSql.AppendLine("  FROM INLIST A                      ");
            strSql.AppendLine("  LEFT OUTER JOIN IPCHULGO B         ");
            strSql.AppendLine("    ON B.J_ID = A.J_ID               ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE C            ");
            strSql.AppendLine("    ON C.J_SERIAL = A.J_SERIAL       ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D    ");
            strSql.AppendLine("    ON D.DEALER_CD = A.J_ID1         ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS E     ");
            strSql.AppendLine("    ON E.EMP_ID = A.CHRG_ID          ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE F            ");
            strSql.AppendLine("    ON F.GUBUN1 = C.GUBUN1           ");
            strSql.AppendLine("  LEFT OUTER JOIN MESURING G         ");
            strSql.AppendLine("    ON G.IPCHULGO_MAIPID = A.J_ID    ");
            strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.KERATYPE = '매입' ");
            strSql.AppendLine("    AND A.DANGA != 0");
            if (!string.IsNullOrEmpty(sDealerNm)) strSql.AppendLine("    AND D.DEALER_NM = '" + sDealerNm + "' ");
            if (!string.IsNullOrEmpty(sAprvGb))
            {
                strSql.AppendLine("    AND ( ('" + sAprvGb + "' != 'Y' AND (A.APRV_YN IS NULL OR A.APRV_YN = 'N'))  ");
                strSql.AppendLine("          OR  ");
                strSql.AppendLine("          ('" + sAprvGb + "' = 'Y' AND A.APRV_YN = 'Y') ");
                strSql.AppendLine("        ) ");
            }
            strSql.AppendLine("    AND B.KERAGUBUN <> '직송' ");
            strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerPurc;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;

        }

        private void SalesPerBuyer(string sYmdFrom, string sYmdTo, string sDealerNm, string sAprvGb)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            #region mariaDB
            //strSql.AppendLine(" SELECT A.JUNPYOID ");
            //strSql.AppendLine("      , 'N' AS CHK");
            //strSql.AppendLine("      , A.KERATYPE ");
            //strSql.AppendLine("      , A.J_DATE ");
            //strSql.AppendLine("      , G.J_DATE AS MESURE_DT ");
            //strSql.AppendLine("      , A.APRV_DATE ");
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
            ////strSql.AppendLine("      , IFNULL(A.MIDANGA, 0) - IFNULL(A.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , (((IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT) - IFNULL(A.MIDANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine("      , ((IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE ");
            //strSql.AppendLine("      , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine("      , A.KONGKEP AS TOTALPRICE ");
            //strSql.AppendLine("      , (IFNULL(A.OWEIGHT, 0) * IFNULL(A.DANGA, 0)) - IFNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            //strSql.AppendLine("      , G.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("      , G.JUNPYOID AS IMAGE ");
            //strSql.AppendLine("      , G.GUMSUBIGO AS INSPECTNOTE ");
            //strSql.AppendLine("      , TRUNCATE((IFNULL(A.CKONGKEP, 0) / A.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
            //strSql.AppendLine("      , IFNULL(A.SEAKPOHAM, 'N') AS SEAKPOHAM "); //포함 : N, 미포함 : Y
            //strSql.AppendLine("   FROM INLIST A ");
            //strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            //strSql.AppendLine("     ON B.J_ID = A.J_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            //strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            //strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS E ");
            //strSql.AppendLine("     ON E.EMP_ID = A.CHRG_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE F ");
            //strSql.AppendLine("     ON F.GUBUN1 = C.GUBUN1 ");
            //strSql.AppendLine("   LEFT OUTER JOIN MESURING G ");
            //strSql.AppendLine("     ON G.IPCHULGO_MACHULID = A.J_ID ");
            #endregion

            strSql.AppendLine("SET ANSI_WARNINGS OFF                                            ");
            strSql.AppendLine("SET ARITHIGNORE ON                                               ");
            strSql.AppendLine("SET ARITHABORT OFF                                               ");
            strSql.AppendLine("SELECT A.JUNPYOID                                                ");
            strSql.AppendLine("     , 'N' AS CHK                                                ");
            strSql.AppendLine("     , A.KERATYPE                                                ");
            strSql.AppendLine("     , A.J_DATE                                                  ");
            strSql.AppendLine("     , G.J_DATE AS MESURE_DT                                     ");
            strSql.AppendLine("     , A.APRV_DATE                                               ");
            strSql.AppendLine("     , D.DEALER_NM                                               ");
            strSql.AppendLine("     , E.EMP_NM AS INSPECTOR                                     ");
            strSql.AppendLine("     , B.J_BNUM                                                  ");
            strSql.AppendLine("     , G.WEIGHT AS LANDEDWEIGHT                                  ");
            strSql.AppendLine("     , A.HALIN AS LOSS                                           ");
            strSql.AppendLine("     , A.OWEIGHT AS ACCEPTWEIGHT                                 ");
            strSql.AppendLine("     , ISNULL(A.OWEIGHT, 0) - ISNULL(G.WEIGHT, 0) AS DIFFWEIGHT  ");
            strSql.AppendLine("     , F.GUBUN1                                                  ");
            strSql.AppendLine("     , A.MIDANGA AS STDDUNITPRICE                                ");
            strSql.AppendLine("     , A.DANGA AS SALEUNITPRICE                                  ");
            //strSql.AppendLine("     , ISNULL(A.MIDANGA, 0) - ISNULL(A.DANGA, 0) AS DIFFUNITPRICE");
            strSql.AppendLine("     , (((ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0)) / A.OWEIGHT) -ISNULL(A.MIDANGA, 0) AS DIFFUNITPRICE");
            strSql.AppendLine("     , ((ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRVUNITPRICE");
            strSql.AppendLine("     , ISNULL(A.CKONGKEP, 0) AS CARRYCOST");
            strSql.AppendLine("     , A.KONGKEP AS TOTALPRICE           ");
            strSql.AppendLine("     , (ISNULL(A.OWEIGHT, 0) * ISNULL(A.DANGA, 0)) - ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE");
            strSql.AppendLine("     , G.J_STATE AS LOSSCAUSE     ");
            strSql.AppendLine("     , G.JUNPYOID AS IMAGE        ");
            strSql.AppendLine("     , G.GUMSUBIGO AS INSPECTNOTE ");
            strSql.AppendLine("     , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG), 1, 1) AS CARRY_UNIT_PRICE");
            strSql.AppendLine("     , ISNULL(A.SEAKPOHAM, 'N') AS SEAKPOHAM");
            strSql.AppendLine("  FROM INLIST A                             ");
            strSql.AppendLine("  LEFT OUTER JOIN IPCHULGO B                ");
            strSql.AppendLine("    ON B.J_ID = A.J_ID                      ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE C                   ");
            strSql.AppendLine("    ON C.J_SERIAL = A.J_SERIAL              ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD D           ");
            strSql.AppendLine("    ON D.DEALER_CD = A.J_ID1                ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS E            ");
            strSql.AppendLine("    ON E.EMP_ID = A.CHRG_ID                 ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE F                   ");
            strSql.AppendLine("    ON F.GUBUN1 = C.GUBUN1                  ");
            strSql.AppendLine("  LEFT OUTER JOIN MESURING G                ");
            strSql.AppendLine("    ON G.IPCHULGO_MACHULID = A.J_ID         ");
            strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            strSql.AppendLine("    AND A.DANGA != 0");
            if (!string.IsNullOrEmpty(sDealerNm)) strSql.AppendLine("    AND D.DEALER_NM = '" + sDealerNm + "' ");
            if (!string.IsNullOrEmpty(sAprvGb))
            {
                strSql.AppendLine("    AND ( ('" + sAprvGb + "' != 'Y' AND (A.APRV_YN IS NULL OR A.APRV_YN = 'N'))  ");
                strSql.AppendLine("          OR  ");
                strSql.AppendLine("          ('" + sAprvGb + "' = 'Y' AND A.APRV_YN = 'Y') ");
                strSql.AppendLine("        ) ");
            }
            strSql.AppendLine("    AND B.KERAGUBUN <> '직송' ");
            strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerSales;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void DirectSendPerBuyer(string sYmdFrom, string sYmdTo, string sDealerNm, string sAprvGb)
        {
            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            #region mariaDB
            //strSql.AppendLine(" SELECT B.J_DATE ");
            //strSql.AppendLine("      , A.KERATYPE ");
            //strSql.AppendLine("      , A.JUNPYOID AS JUNPYOID_MESURING ");
            //strSql.AppendLine("      , D.EMP_NM AS NAME ");
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
            //strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS XX ");
            //strSql.AppendLine("    ON C.CHRG_ID = XX.EMP_ID ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD F ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD1 = F.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD G ");
            //strSql.AppendLine("    ON A.ETC_DEALER_CD2 = G.DEALER_CD ");
            //strSql.AppendLine("  LEFT OUTER JOIN JAJAE H ");
            //strSql.AppendLine("    ON A.J_SERIAL = H.J_SERIAL  ");
            //strSql.AppendLine(" WHERE A.KERATYPE = '직송' ");
            #endregion

            strSql.AppendLine("SELECT B.J_DATE                        ");
            strSql.AppendLine("     , A.KERATYPE                      ");
            strSql.AppendLine("     , A.JUNPYOID AS JUNPYOID_MESURING ");
            strSql.AppendLine("     , D.EMP_NM AS NAME                ");
            strSql.AppendLine("     , H.GUBUN1                        ");
            strSql.AppendLine("     , A.J_BNUM                        ");
            strSql.AppendLine("     , D.DEALER_NM AS SALE_DEALER_NM   ");
            strSql.AppendLine("     , B.DANJUNG AS OWEIGHT            ");
            strSql.AppendLine("     , B.DANGA AS SALEUNITPRICE        ");
            strSql.AppendLine("     , B.KONGKEP AS SALEPRICE          ");
            strSql.AppendLine("     , E.DEALER_NM AS PURC_DEALER_NM   ");
            strSql.AppendLine("     , C.DANGA AS PURCHUNITPRICE       ");
            strSql.AppendLine("     , C.IKONGKEP AS PURC_AMT          ");
            strSql.AppendLine("     , ROUND(((ISNULL(B.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1, 1) AS ARRIVEUNITPRICE");
            strSql.AppendLine("       , ISNULL(B.CKONGKEP, 0) AS CARRYCOST                                                    ");
            strSql.AppendLine("       , ROUND((ISNULL(B.CKONGKEP, 0) / B.DANJUNG), 1, 1) AS CARRY_UNIT_PRICE                  ");
            strSql.AppendLine("        , (ISNULL(C.DANJUNG, 0) * ISNULL(C.DANGA, 0)) +ISNULL(B.CKONGKEP, 0) AS ARRVTOTALPRICE ");
            strSql.AppendLine("       , ISNULL(B.DANGA, 0) -ROUND(((ISNULL(B.DANJUNG, 0) * ISNULL(C.DANGA, 0)) + ISNULL(B.CKONGKEP, 0)) / B.DANJUNG, 1, 1) AS DIFFUNITPRICE");
            strSql.AppendLine("        , C.HALIN                    ");
            strSql.AppendLine("     , I.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("     , A.ETC_COST1                   ");
            strSql.AppendLine("     , A.ETC_REMARK1                 ");
            strSql.AppendLine("     , J.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("     , A.ETC_COST2                   ");
            strSql.AppendLine("     , A.ETC_REMARK2                 ");
            strSql.AppendLine("  FROM MESURING A                    ");
            strSql.AppendLine(" LEFT OUTER JOIN INLIST B            ");
            strSql.AppendLine("   ON A.IPCHULGO_MACHULID = B.J_ID   ");
            strSql.AppendLine(" LEFT OUTER JOIN INLIST C            ");
            strSql.AppendLine("   ON A.IPCHULGO_MAIPID = C.J_ID     ");
            strSql.AppendLine(" LEFT OUTER JOIN(SELECT X2.EMP_NM, X1.DEALER_CD, X1.DEALER_NM, X1.CHRG_ID");
            strSql.AppendLine("                     FROM ACC_DEALER_CD X1                               ");
            strSql.AppendLine("                     LEFT OUTER JOIN HR_EMP_BASIS X2                     ");
            strSql.AppendLine("                       ON X1.CHRG_ID = X2.EMP_ID) D                      ");
            strSql.AppendLine("  ON A.J_ASSIGNID = D.DEALER_CD--매출처                                  ");
            strSql.AppendLine("LEFT OUTER JOIN ACC_DEALER_CD E --매입처                                 ");
            strSql.AppendLine("   ON C.J_ID1 = E.DEALER_CD                                              ");
            strSql.AppendLine(" LEFT OUTER JOIN HR_EMP_BASIS XX                                         ");
            strSql.AppendLine("   ON C.CHRG_ID = XX.EMP_ID                                              ");
            strSql.AppendLine(" LEFT OUTER JOIN JAJAE H                                                 ");
            strSql.AppendLine("   ON A.J_SERIAL = H.J_SERIAL                                            ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD I          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD1 = CONVERT(VARCHAR,I.DEALER_CD)");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD J          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD2 = CONVERT(VARCHAR,J.DEALER_CD)");
            strSql.AppendLine("WHERE A.KERATYPE = '직송'");
            strSql.AppendLine("   AND B.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("   AND C.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("   AND B.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            if (!string.IsNullOrEmpty(sAprvGb))
            {
                strSql.AppendLine("    AND ( ('" + sAprvGb + "' != 'Y' AND (B.APRV_YN IS NULL OR B.APRV_YN = 'N'))  ");
                strSql.AppendLine("          OR  ");
                strSql.AppendLine("          ('" + sAprvGb + "' = 'Y' AND B.APRV_YN = 'Y') ");
                strSql.AppendLine("        ) ");
            }
            strSql.AppendLine("  ORDER BY B.J_DATE, REPLACE(E.DEALER_NM, '(주)', ''), A.J_SERIAL , B.KONGKEP DESC");

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
            //strSql.AppendLine(" 	 , F.JUNPYOID AS JUNPYOID_MESURING ");
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
            ////strSql.AppendLine("    AND G.KERATYPE = '매입' ");
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
            //if (!string.IsNullOrEmpty(sAprvGb))
            //{
            //    strSql.AppendLine("    AND ( ('" + sAprvGb + "' != 'Y' AND (A.APRV_YN IS NULL OR A.APRV_YN = 'N'))  ");
            //    strSql.AppendLine("          OR  ");
            //    strSql.AppendLine("          ('" + sAprvGb + "' = 'Y' AND A.APRV_YN = 'Y') ");
            //    strSql.AppendLine("        ) ");
            //}
            //strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerDtSend;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;

            //strSql.AppendLine(" SELECT A.JUNPYOID ");		object.ToString 반환됨	" SELECT A.J_DATE \r\n \t , C.GUBUN1 \r\n \t , I.EMP_NM \r\n \t , B.J_BNUM \r\n \t , D.DEALER_NM \r\n      , A.OWEIGHT\r\n\t     , A.DANGA AS SALEUNITPRICE \r\n \t , A.KONGKEP AS SALEPRICE \r\n \t , A.J_BOOKING \r\n      , G.DANGA AS PURCHUNITPRICE \r\n\t     , G.IKONGKEP \r\n \t , IFNULL(A.DANGA, 0) - IFNULL(G.DANGA, 0) AS DIFFUNITPRICE \r\n \t , F.JUNPYOID AS JUNPYOID_MESURING \r\n \t , ((IFNULL(A.OWEIGHT, 0) * IFNULL(G.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRIVEUNITPRICE \r\n \t , IFNULL(A.CKONGKEP, 0) AS CARRYCOST \r\n \t , (IFNULL(A.OWEIGHT, 0) * IFNULL(G.DANGA, 0)) + IFNULL(A.KONGKEP, 0) ARRVTOTALPRICE \r\n \t , G.HALIN \r\n      , TRUNCATE((IFNULL(A.CKONGKEP, 0) / G.DANJUNG), 1) AS CARRY_UNIT_PRICE \r\n      , J.DEALER_NM AS ETC_DEALER_NM1 \r\n      , F.ETC_COST1 \r\n      , F.ETC_REMARK1 \r\n      , K.DEALER_NM AS ETC_DEALER_NM2 \r\n      , F.ETC_COST2 \r\n      , F.ETC_REMARK2 \r\n   FROM INLIST A \r\n   LEFT OUTER JOIN IPCHULGO B \r\n     ON B.J_ID = A.J_ID \r\n   LEFT OUTER JOIN JAJAE C \r\n     ON C.J_SERIAL = A.J_SERIAL \r\n   LEFT OUTER JOIN ACC_DEALER_CD D \r\n     ON D.DEALER_CD = A.J_ID1 \r\n   LEFT OUTER JOIN JAJAE E \r\n     ON E.GUBUN1 = C.GUBUN1 \r\n   LEFT OUTER JOIN MESURING F \r\n     ON F.IPCHULGO_MACHULID = A.J_ID \r\n   LEFT OUTER JOIN INLIST G \r\n     ON G.JUNPYOID = A.JUNPYOID \r\n   LEFT OUTER JOIN ACC_DEALER_CD H \r\n    AND G.KERATYPE = '매입' \r\n    \tON H.DEALER_CD = G.J_ID1 \r\n   LEFT OUTER JOIN HR_EMP_BASIS I \r\n    \tON I.EMP_ID = H.CHRG_ID \r\n   LEFT OUTER JOIN ACC_DEALER_CD J\r\n     ON J.DEALER_CD = F.ETC_DEALER_CD1 \r\n   LEFT OUTER JOIN ACC_DEALER_CD K\r\n     ON K.DEALER_CD = F.ETC_DEALER_CD2 \r\n  WHERE A.J_DATE >= '2020-06-01' \r\n    AND A.J_DATE <= '2020-06-04' \r\n    AND A.KERATYPE = '매출' \r\n    AND F.KERATYPE = '직송' \r\n    AND F.J_CHECK = '1' \r\n    AND (('N' = '' AND 1 = 1)  \r\n         OR \r\n         ('N' = 'Y' AND A.APRV_YN = 'Y')\r\n         OR\r\n         ('N' = 'N' AND A.APRV_YN = 'N') )\r\n  ORDER BY A.J_DATE, D.DEALER_NM \r\n"	string

            //strSql.AppendLine("      , 'N' AS CHK");
            //strSql.AppendLine("      , A.J_DATE ");
            //strSql.AppendLine("      , A.APRV_DATE ");
            //strSql.AppendLine(" 	 , C.GUBUN1 ");
            //strSql.AppendLine(" 	 , I.EMP_NM ");
            //strSql.AppendLine(" 	 , B.J_BNUM ");
            //strSql.AppendLine(" 	 , D.DEALER_NM ");
            //strSql.AppendLine("      , A.OWEIGHT");
            //strSql.AppendLine("	     , A.DANGA AS SALEUNITPRICE ");
            //strSql.AppendLine(" 	 , A.KONGKEP AS SALEPRICE ");
            //strSql.AppendLine(" 	 , A.J_BOOKING ");
            //strSql.AppendLine("      , G.DANGA AS PURCHUNITPRICE ");
            //strSql.AppendLine("	     , G.IKONGKEP ");
            //strSql.AppendLine(" 	 , IFNULL(A.DANGA, 0) - IFNULL(G.DANGA, 0) AS DIFFUNITPRICE ");
            //strSql.AppendLine(" 	 , F.JUNPYOID AS JUNPYOID_MESURING");
            //strSql.AppendLine(" 	 , ((IFNULL(A.OWEIGHT, 0) * IFNULL(G.DANGA, 0)) + IFNULL(A.CKONGKEP, 0)) / A.OWEIGHT AS ARRIVEUNITPRICE ");
            //strSql.AppendLine(" 	 , IFNULL(A.CKONGKEP, 0) AS CARRYCOST ");
            //strSql.AppendLine(" 	 , (IFNULL(A.OWEIGHT, 0) * IFNULL(G.DANGA, 0)) + IFNULL(A.KONGKEP, 0) ARRVTOTALPRICE ");
            //strSql.AppendLine(" 	 , G.HALIN ");
            //strSql.AppendLine("      , TRUNCATE((IFNULL(A.CKONGKEP, 0) / G.DANJUNG), 1) AS CARRY_UNIT_PRICE ");
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
            //strSql.AppendLine("     ON G.J_LOTNO = A.J_LOTNO ");
            //strSql.AppendLine("    AND G.KERATYPE = '매입' ");
            //strSql.AppendLine("        LEFT OUTER JOIN ACC_DEALER_CD H ");
            //strSql.AppendLine("    	ON H.DEALER_CD = G.J_ID1 ");
            //strSql.AppendLine("    	   LEFT OUTER JOIN HR_EMP_BASIS I ");
            //strSql.AppendLine("    	ON I.EMP_ID = H.CHRG_ID ");
            //strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            //strSql.AppendLine("    AND F.KERATYPE = '직송' ");
            //if(!string.IsNullOrEmpty(sDealerNm)) strSql.AppendLine("    AND D.DEALER_NM = '" + sDealerNm + "' ");
            //if (!string.IsNullOrEmpty(sAprvGb))
            //{
            //    strSql.AppendLine("    AND ( ('" + sAprvGb + "' != 'Y' AND (A.APRV_YN IS NULL OR A.APRV_YN = 'N'))  ");
            //    strSql.AppendLine("          OR  ");
            //    strSql.AppendLine("          ('" + sAprvGb + "' = 'Y' AND A.APRV_YN = 'Y') ");
            //    strSql.AppendLine("        ) ");
            //}
            //strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //GridBuyer.MainView = GridViewBuyerDtSend;
            //GridBuyer.DataSource = dt;

            //Cursor = Cursors.Default;
        }

        #endregion[Query]

        #region[기타 함수]

        private void SetButtonVisible()
        {
            string sAprvYn = RdgbAprvGb.EditValue?.ToString();
            if (string.IsNullOrEmpty(sAprvYn))
            {
                LayoutAprvClose.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutAprvCancle.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else if (sAprvYn.Equals("Y"))
            {
                LayoutAprvClose.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutAprvCancle.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else if (sAprvYn.Equals("N"))
            {
                LayoutAprvClose.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutAprvCancle.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }

        private bool CheckApproval()
        {
            string sAprvYn = RdgbAprvGb.EditValue?.ToString();
            if (sAprvYn.Equals("Y"))
            {
                if (LayoutAprvCancle.Visibility == DevExpress.XtraLayout.Utils.LayoutVisibility.Always)
                    return false;
                else
                    return true;
            }
            else if (sAprvYn.Equals("N"))
            {
                if (LayoutAprvClose.Visibility == DevExpress.XtraLayout.Utils.LayoutVisibility.Always)
                    return false;
                else
                    return true;
            }
            else
            {
                return true;
            }
        }

        private bool GetApprovalInfo()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.APV_Y ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("  WHERE A.PGMID = '" + this.Name + "' ");
            strSql.AppendLine("    AND A.USRCD = '" + FmMainToolBar2.UserID + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count > 0)
            {
                string sAprvYn = dt.Rows[0]["APV_Y"].ToString();
                if (sAprvYn.Equals("Y"))
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
        }

        #endregion[기타 함수]

        private void GridViewBuyerPurc_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewBuyerPurc_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void RepoPurcChk_EditValueChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            string sYn = chk.EditValue?.ToString();
            if (sYn.Equals("Y"))
            {
                GridView view = (GridView)GridBuyer.MainView;
                string sJunpyoId = view.GetFocusedRowCellValue("JUNPYOID")?.ToString();
                Cursor = Cursors.WaitCursor;
                if (CheckTaxBillIssueYn(sJunpyoId))
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("해당 건은 세금계산서가 존재하여 취소할 수 없습니다.");
                    chk.EditValue = "N";
                    return;
                }
            }
        }

        private bool CheckTaxBillIssueYn(string sJunpyoId)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
            strSql.AppendLine("   FROM ACC_TAX_MGT ");
            strSql.AppendLine("  WHERE BILL_KEY = '" + sJunpyoId + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count > 0)   //카운트가 0 이거나 데이터가 없을 때 true 리턴
            {
                double dCnt = dt.Rows[0]["CNT"] == null ? 0 : Convert.ToDouble(dt.Rows[0]["CNT"]);
                if (dCnt > 0)
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
        }

        private void LkupDealerNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewBuyerPurc_RowClick(object sender, RowClickEventArgs e)
        {
            if (GridViewBuyerPurc.IsRowSelected(e.RowHandle))
                GridViewBuyerPurc.UnselectRow(e.RowHandle);
            else
                GridViewBuyerPurc.SelectRow(e.RowHandle);
        }

        private void GridViewBuyerSales_RowClick(object sender, RowClickEventArgs e)
        {
            if (GridViewBuyerSales.IsRowSelected(e.RowHandle))
                GridViewBuyerSales.UnselectRow(e.RowHandle);
            else
                GridViewBuyerSales.SelectRow(e.RowHandle);
        }

        #region [정렬기능(2020-06-02 정은영)]
        private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        {
            //GridView view = (GridView)sender;
            //GridHitInfo hitInfo = view.CalcHitInfo(e.Location);


            //if (hitInfo.InColumn)
            //{
            //    string sColName = hitInfo.Column.Name;

            //    if (!sColName.Equals("DX$CheckboxSelectorColumn"))
            //    {
            //        if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.None)
            //        {
            //            hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
            //            view.FocusedRowHandle = 0;
            //        }
            //        else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Descending)
            //        {
            //            hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            //            view.FocusedRowHandle = 0;
            //        }
            //        else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Ascending)
            //        {
            //            hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.None;
            //            view.FocusedRowHandle = 0;
            //        }
            //    }

            //    //if ((ModifierKeys & Keys.Control) == Keys.Control) return;
            //    //if ((ModifierKeys & Keys.Shift) != Keys.Shift) view.ClearSorting();
            //}
        }
        #endregion

        private void AC11001F01_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }
    }
}