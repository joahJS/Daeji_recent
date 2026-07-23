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
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraBars;
using System.Data.SqlClient;
using DevExpress.DataAccess.Native.Json;
using Newtonsoft.Json;
using DevExpress.Data;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * --------------------HISTORY------------------
 * 1. 2021-02-07 (현업요청)
 *    거래처 초성검색 추가
 *    
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2021-03-31
 * 수정자   : 고혜성
 * Reference Key : #0001
 * 수정내용 : (현업요청)
 *            세금계산서 삭제기능 추가(INLIST 자료)
 *            
 * 수정일자 : 2023-01-05
 * 수정자   : 정은영
 * Reference Key: #0002
 * 수정내용 : (현업요청)
 *             1. 매출 데이터 인센티브 추가
 *             
 * 수정일자 : 2023-01-18
 * 수정자   : 정은영
 * Reference Key: #0003
 * 수정내용 : 1. 인센티브 계산서 발급 오류 수정
 */
namespace AccAdm
{
    public partial class AC18001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC18001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC18001F01_Load(object sender, EventArgs e)
        {
            //DateTime dtFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));
            //DateTime dLastDayOfLastMonth = dtFirstDayOfThisMonth.AddDays(-1);
            //DateTime dFirstDayOfLastMonth = dtFirstDayOfThisMonth.AddMonths(-1);

            //DateEditFrom.EditValue = dFirstDayOfLastMonth;
            //DateEditTo.EditValue = dLastDayOfLastMonth;
            LayoutDropBtn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AC18001F01");
            arrGrdView = new GridView[] { GridViewL1, GridViewL2, GridViewR1, GridViewR2, GridViewP1, GridViewP2 };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout1.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl2.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout2.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl3.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout3.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl6.RestoreLayoutFromXml(sFile);
            }
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            int iFindSbj = CboFindSbj.SelectedIndex;
            string sFindWord = TxtFindWord.EditValue?.ToString().Trim();

            if (string.IsNullOrEmpty(sFindWord))
                sFindWord = "";

            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("검색일자를 설정하세요.");
                DateEditFrom.SelectAll();
                DateEditTo.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("검색일자를 설정하세요.");
                DateEditTo.SelectAll();
                DateEditTo.Focus();
                return;
            }

            if(iFindSbj < 0)
            {
                XtraMessageBox.Show("찾을항목을 바르게 선택하세요.");
                CboFindSbj.Focus();
                return;
            }

            Cursor = Cursors.WaitCursor;

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", sYmdFrom);
            dicParams.Add("DATE_T", sYmdTo);
            dicParams.Add("FIND_IDX", iFindSbj.ToString());
            dicParams.Add("FIND_WORD", sFindWord);
            
            if(TabControl.SelectedTabPage == TabPageTax)
            {
                GridRetrL1.DataSource = null;
                GridRetrL1.DataSource = GetIssueInfo(dicParams);
                if(GridViewL1.RowCount > 0)
                {
                    GridViewL1.Focus();
                }
                else if(GridViewL1.RowCount == 0)
                {
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }
            }
            else if(TabControl.SelectedTabPage == TabPageIssueSale)
            {
                GridRetrR1.DataSource = null;
                GridRetrR1.DataSource = GetIssueDealerInfo(dicParams);
                if(GridViewR1.RowCount > 0)
                {
                    GridViewR1.Focus();
                }
                else if(GridViewR1.RowCount == 0)
                {
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }
            }
            else if(TabControl.SelectedTabPage == TabPageIssuePurc)
            {
                GridRetrP1.DataSource = null;
                GridRetrP1.DataSource = GetIssueDealerInfo2(dicParams);
            }

            Cursor = Cursors.Default;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            AC18001F02 frm = new AC18001F02();
            if(frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            try
            {
                if (TabControl.SelectedTabPage == TabPageTax)
                {
                    DeleteTaxInfo();
                }
                else if (TabControl.SelectedTabPage == TabPageIssuePurc)
                {
                    DataRow drFocusedRow = GridViewP1.GetFocusedDataRow();
                    if(drFocusedRow == null)
                    {
                        XtraMessageBox.Show("올바른 데이터를 선택하세요.");
                        return;
                    }
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Add("DATE_F", drFocusedRow["DATE_F"].ToString());
                    dicParams.Add("DATE_T", drFocusedRow["DATE_T"].ToString());
                    dicParams.Add("DEALER_CD", drFocusedRow["J_ID1"].ToString());
                    dicParams.Add("GROUP_JUNPYOID", drFocusedRow["TEMP"].ToString());

                    string sJunpyoId = GridViewP2.GetFocusedRowCellValue(GridColR2JunpyoId)?.ToString();
                    DeleteIssueInfo(sJunpyoId, dicParams);
                }
                else if (TabControl.SelectedTabPage == TabPageIssueSale)
                {
                    Cursor = Cursors.WaitCursor;
                    DataRow drFocusedRow = GridViewR1.GetFocusedDataRow();
                    if (drFocusedRow == null)
                    {
                        XtraMessageBox.Show("올바른 데이터를 선택하세요.");
                        return;
                    }
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Add("DATE_F", drFocusedRow["DATE_F"].ToString());
                    dicParams.Add("DATE_T", drFocusedRow["DATE_T"].ToString());
                    dicParams.Add("DEALER_CD", drFocusedRow["J_ID1"].ToString());
                    dicParams.Add("GROUP_JUNPYOID", drFocusedRow["TEMP"].ToString());

                    string sJunpyoId = GridViewR2.GetFocusedRowCellValue(GridColP2JunpyoId)?.ToString();
                    DeleteIssueInfo(sJunpyoId, dicParams);
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            
        }

        /*
         * #0001
         */
        private void DeleteIssueInfo(string sJunpyoID, Dictionary<string, string> dicParams)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                StringBuilder strSql = new StringBuilder();

                //해당 건이 직송 건인지 체크
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.J_LOTNO ");
                strSql.AppendLine("      , A.GUBUN1 ");
                strSql.AppendLine("      , B.DEALER_NM ");
                strSql.AppendLine("      , C.KERATYPE AS M_KERA ");
                strSql.AppendLine("      , C.J_DATE AS M_J_DATE ");
                strSql.AppendLine("      , C.J_BNUM ");
                strSql.AppendLine("   FROM INLIST A ");
                strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B ");
                strSql.AppendLine("     ON A.J_ID1 = B.DEALER_CD ");
                strSql.AppendLine("   LEFT JOIN MESURING C ");
                strSql.AppendLine("     ON A.J_ID = CASE WHEN A.KERATYPE = '매입' THEN IPCHULGO_MAIPID ELSE IPCHULGO_MACHULID END ");
                strSql.AppendLine("  WHERE A.JUNPYOID =  " + sJunpyoID + "");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                bool bDirectSendYn = false;
                if(dt.Rows.Count > 0)
                {
                    string sJ_LOTNO = dt.Rows[0]["J_LOTNO"]?.ToString() ?? string.Empty; //4:직송 아닐경우 입/출고
                    if (sJ_LOTNO.Equals("4"))
                    {
                        bDirectSendYn = true;
                    }
                }
                else
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("해당 건은 현재 INLIST에 존재하지 않습니다.\r\nJunpyoID : " + sJunpyoID);
                    return;
                }

                string sMSG = string.Empty;
                string sMesuringDt = dt.Rows[0]["M_J_DATE"]?.ToString() ?? string.Empty;
                if(sMesuringDt.Length >= 10)
                {
                    sMesuringDt = sMesuringDt.Substring(0, 10);
                }
                string sKeraType = dt.Rows[0]["M_KERA"]?.ToString() ?? string.Empty;
                string sJ_BNum = dt.Rows[0]["J_BNUM"]?.ToString() ?? string.Empty;
                string sGrade = dt.Rows[0]["GUBUN1"]?.ToString() ?? string.Empty;
                string sDealerNm = dt.Rows[0]["DEALER_NM"]?.ToString() ?? string.Empty;

                sMSG = string.Format("계근일자 : {0}" +
                    "\r\n거래구분 : {1}" +
                    "\r\n거래처명 : {2}" +
                    "\r\n등급 : {3}" +
                    "\r\n차번 : {4}"
                    , sMesuringDt
                    , sKeraType
                    , sDealerNm
                    , sGrade
                    , sJ_BNum);
                if (!bDirectSendYn)
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("입출고 건은 삭제할 수 없습니다.\r\n" + sMSG);
                    return;
                }
                
                //해당 건이 전표발행여부 체크
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CONVERT(DATE,A.TDATE) AS DT ");
                strSql.AppendLine("      , A.TDATE ");
                strSql.AppendLine("      , A.ATGUB ");
                strSql.AppendLine("      , B.COM_NM AS ATGUB_NM ");
                strSql.AppendLine("      , A.SEQNO ");
                strSql.AppendLine("   FROM ACTRAN A  ");
                strSql.AppendLine("   LEFT JOIN COM_BASE_CD B  ");
                strSql.AppendLine("     ON A.ATGUB = B.COM_CD ");
                strSql.AppendLine("    AND B.CD_GB = 'AC02001_01' ");
                strSql.AppendLine("  WHERE A.REF1 =  '" + sJunpyoID + "'");

                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt2.Rows.Count > 0)
                {
                    string sTDATE = dt2.Rows[0]["DT"]?.ToString() ?? string.Empty;
                    string sATGUB = dt2.Rows[0]["ATGUB_NM"]?.ToString() ?? string.Empty;
                    string sSEQNO = dt2.Rows[0]["SEQNO"]?.ToString() ?? string.Empty;

                    Cursor = Cursors.Default;
                    string sMsg = string.Format("해당 건은 전표발행 중입니다." +
                        "\r\n전표일자 : {0}" +
                        "\r\n전표구분 : {1}" +
                        "\r\n전표번호 : {2}"
                        , sTDATE
                        , sATGUB
                        , sSEQNO);
                    XtraMessageBox.Show(sMsg);
                    return;
                }

                if (XtraMessageBox.Show("해당 건을 삭제하시겠습니까?\r\n" + sMSG, "직송 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    return;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM INLIST ");
                strSql.AppendLine("       WHERE JUNPYOID = " + sJunpyoID + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제가 완료되었습니다.");
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void DeleteTaxInfo()
        {
            if (GridViewL1.RowCount == 0)
                return;

            string sTaxNo = GridViewL1.GetFocusedRowCellValue(GridColL1TaxNo)?.ToString();
            string sMSG = string.Format("계산서번호 : {0}" +
                                    "\r\n거래처명 : {1}" +
                                    "\r\n합계금액 : {2}" +
                                    "\r\n해당 계산서를 삭제하시겠습니까?"
                                    , sTaxNo
                                    , GridViewL1.GetFocusedRowCellValue(GridColL1DealerNm)?.ToString()
                                    , GridViewL1.GetFocusedRowCellValue(GridColL1TotAmt1)?.ToString());

            if (XtraMessageBox.Show(sMSG, "계산서 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            if (string.IsNullOrEmpty(sTaxNo))
            {
                XtraMessageBox.Show("계산서번호가 유효하지 않습니다.\r\n관리자에게 문의하세요.");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                //INLIST TAXNO 초기화
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" UPDATE INLIST ");
                strSql.AppendLine("    SET TAXNO = NULL ");
                strSql.AppendLine("  WHERE TAXNO = @TAXNO ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@TAXNO", sTaxNo);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();


                //ACTRAN TAXNO 초기화
                strSql.Clear();
                strSql.AppendLine(" UPDATE ACTRAN         ");
                strSql.AppendLine("    SET TAXNO = ''     ");
                strSql.AppendLine("  WHERE TAXNO = @TAXNO ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@TAXNO", sTaxNo);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                //TAXF 삭제
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM TAXF ");
                strSql.AppendLine("       WHERE TAXNO = @TAXNO ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@TAXNO", sTaxNo);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제가 완료되었습니다.");

                BtnRetr.PerformClick();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnIssue_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (TabControl.SelectedTabPage == TabPageIssueSale)
            {
                IssueSale();
            }
            else if(TabControl.SelectedTabPage == TabPageIssuePurc)
            {
                IssuePurc();
            }
            else
            {
                XtraMessageBox.Show("매출입 세금계산서 발행 탭에서 가능한 버튼입니다.");
                return;
            }
        }

        private void IssueSale()
        {
            //TEMP 컬럼에 값(INLIST 테이블의 PK값인 JUNPYOID) 이 전체 그리드에 하나도 존재하지 않는 경우
            //동시에 있는 값들은 Dictionary 포함
            Dictionary<string, string> dicDelaerInfo = new Dictionary<string, string>();
            Dictionary<string, string> dicDelaerInfoInsen = new Dictionary<string, string>();//#0003

            int iCnt = 0;
            for (int i = 0; i < GridViewR1.RowCount; i++)
            {
                string sCnt = GridViewR1.GetRowCellValue(i, GridColR1Cnt)?.ToString();
                string sDealerCd = GridViewR1.GetRowCellValue(i, GridColR1CvCod)?.ToString();
                string[] sArr = GridViewR1.GetRowCellValue(i, GridColR1Temp)?.ToString().Replace(" ", "").Split(',');

                if (sArr.Length == 1)
                {
                    if (string.IsNullOrEmpty(sArr[0]))
                        continue;
                }

                if (sArr.Length > 0)
                {
                    dicDelaerInfo.Add(sDealerCd, GridViewR1.GetRowCellValue(i, GridColR1Temp)?.ToString().Replace(" ", ""));
                    //iCnt++;
                }

                //#0003
                string sJson = GridViewR1.GetRowCellValue(i, GridColR1Temp2)?.ToString();

                if (!string.IsNullOrEmpty(sJson))
                {
                    dicDelaerInfoInsen.Add(sDealerCd, sJson);
                }

                //#0003 선택된 값 확인하는걸로 변경
                if(int.TryParse(sCnt, out int dbl))
                {
                    iCnt += dbl;
                }
            }

            if (iCnt == 0)
            {
                XtraMessageBox.Show("리스트의 어떠한 마감자료도 선택되지 않았습니다.\r\n발행하려는 마감자료를 선택해주세요.");
                return;
            }

            #region 2023-01-18 미사용 삭제
            //Int32 iChk = 0;
            //for (int i = 0; i < GridViewR1.RowCount; i++)
            //{
            //    string sChk = GridViewR1.GetRowCellValue(i, GridColR1Chk)?.ToString();
            //    if (string.IsNullOrEmpty(sChk))
            //        return;

            //    if (sChk.Equals("Y"))
            //        iChk++;
            //}
            #endregion

            string sMSG = string.Format("선택된 건들에 대하여 일괄저장을 진행하시겠습니까?");
            if (XtraMessageBox.Show(sMSG, "계산서 발행여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                foreach (KeyValuePair<string, string> dic in dicDelaerInfo)
                {
                    string sDealerCd = string.Empty;
                    string sPDAT1 = string.Empty;
                    string sPDAT2 = string.Empty;
                    string sCVCOD = string.Empty;
                    string sRK = string.Empty;
                    string sMMDD1 = string.Empty;
                    string sITNM1 = string.Empty;
                    string sTWGT1 = string.Empty;
                    string sCOST1 = string.Empty;
                    string sTAMT1 = string.Empty;
                    string sTVAT1 = string.Empty;

                    sDealerCd = dic.Key;
                    string[] sArrJunpyoIDs = RemoveAt(dic.Value.Replace(" ", "").Split(','), "인센티브");

                    string sArrJunpyoID = string.Empty;
                    if (sArrJunpyoIDs != null && sArrJunpyoIDs.Length != 0)
                    {
                        sArrJunpyoID = sArrJunpyoIDs.Aggregate((text, next) => text + "," + next);
                    }

                    string sJson = string.Empty;
                    foreach(KeyValuePair<string, string> dic2 in dicDelaerInfoInsen)
                    {
                        string sDealerCd2 = dic2.Key;

                        if (sDealerCd.Equals(sDealerCd2))
                        {
                            sJson = dic2.Value;
                        }
                    }

                    //TAXF INSERT 위한 내역 조회
                    dicParams.Clear();
                    dicParams.Add("DEALER_CD", sDealerCd);

                    strSql.Clear();
                    strSql.AppendLine(" ");

                    #region mariaDB
                    //strSql.AppendLine(" SELECT DATE_FORMAT(X1.PDAT1, '%Y-%m-%d') AS PDAT1 ");
                    //strSql.AppendLine("      , DATE_FORMAT(X1.PDAT2, '%Y-%m-%d') AS PDAT2 ");
                    //strSql.AppendLine("      , X1.CVCOD ");
                    //strSql.AppendLine("      , CONCAT('거래기간 : ', X1.PDAT1, ' ~ ', X1.PDAT2) AS RK ");
                    //strSql.AppendLine("      , DATE_FORMAT(NOW(), '%m%d') AS MMDD1 ");
                    //strSql.AppendLine("      , CONCAT(X2.GUBUN1, ' 외 ', X1.CNT, ' 건') AS ITNM1 ");
                    //strSql.AppendLine("      , X1.WGT AS TWGT1 ");
                    //strSql.AppendLine("      , X1.DANGA AS COST1 ");
                    //strSql.AppendLine("      , X1.AMT AS TAMT1 ");
                    //strSql.AppendLine("      , X1.VAT AS TVAT1 ");
                    //strSql.AppendLine("   FROM ( ");
                    //strSql.AppendLine("          SELECT MAX(A.J_SERIAL) AS J_SERIAL ");
                    //strSql.AppendLine("               , MAX(A.J_ID1) AS CVCOD ");
                    //strSql.AppendLine("               , COUNT(*) - 1 AS CNT ");
                    //strSql.AppendLine("               , SUM(A.DANJUNG) AS WGT ");
                    //strSql.AppendLine("               , TRUNCATE(AVG(A.DANGA), 1) AS DANGA ");
                    //strSql.AppendLine("               , SUM(A.KONGKEP) AS AMT  ");
                    //strSql.AppendLine("               , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END) AS VAT ");
                    //strSql.AppendLine("               , MIN(A.J_DATE) AS PDAT1 ");
                    //strSql.AppendLine("               , MAX(A.J_DATE) AS PDAT2 ");
                    //strSql.AppendLine("            FROM INLIST A ");
                    //strSql.AppendLine("            LEFT JOIN MESURING B  ");
                    //strSql.AppendLine("              ON A.J_ID = B.IPCHULGO_MACHULID ");
                    //strSql.AppendLine("           WHERE A.J_ID1 = " + sDealerCd + "");
                    //strSql.AppendLine("             AND A.JUNPYOID IN (" + dic.Value + ") ");
                    //strSql.AppendLine("        ) X1  ");
                    //strSql.AppendLine("   LEFT JOIN JAJAE X2 ");
                    //strSql.AppendLine("     ON X1.J_SERIAL = X2.J_SERIAL ");
                    #endregion

                    #region 2023-01-18 이전 인센티브 저장 전 로직
                    //strSql.AppendLine("SELECT X1.PDAT1 AS PDAT1 ");
                    //strSql.AppendLine("     , X1.PDAT2 AS PDAT2 ");
                    //strSql.AppendLine("     , X1.CVCOD ");
                    //strSql.AppendLine("     , CONCAT('거래기간 : ', X1.PDAT1, ' ~ ', X1.PDAT2) AS RK ");
                    //strSql.AppendLine("     , SUBSTRING(CONVERT(VARCHAR(8), GETDATE(), 112), 5, 4) AS MMDD1 ");
                    //strSql.AppendLine("     , CONCAT(X2.GUBUN1, ' 외 ') AS ITNM1 ");
                    //strSql.AppendLine("     , X1.CNT AS ITNM2 ");
                    //strSql.AppendLine("     , X1.WGT AS TWGT1 ");
                    //strSql.AppendLine("     , X1.DANGA AS COST1 ");
                    //strSql.AppendLine("     , X1.AMT AS TAMT1 ");
                    //strSql.AppendLine("     , X1.VAT AS TVAT1 ");
                    //strSql.AppendLine("  FROM(");
                    //strSql.AppendLine("         SELECT MAX(A.J_SERIAL) AS J_SERIAL ");
                    //strSql.AppendLine("              , MAX(A.J_ID1) AS CVCOD ");
                    //strSql.AppendLine("              , COUNT(*) - 1 AS CNT ");
                    //strSql.AppendLine("              , SUM(A.DANJUNG) AS WGT ");
                    //strSql.AppendLine("              , ROUND(AVG(A.DANGA),0,1) AS DANGA ");
                    //strSql.AppendLine("              , SUM(A.KONGKEP) AS AMT  ");
                    //strSql.AppendLine("              , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END) AS VAT");
                    //strSql.AppendLine("              , MIN(A.J_DATE) AS PDAT1 ");
                    //strSql.AppendLine("              , MAX(A.J_DATE) AS PDAT2 ");
                    //strSql.AppendLine("           FROM INLIST A ");
                    //strSql.AppendLine("           LEFT JOIN MESURING B  ");
                    //strSql.AppendLine("             ON A.J_ID = B.IPCHULGO_MACHULID ");
                    //strSql.AppendLine("          WHERE A.J_ID1 = " + sDealerCd + "");
                    //strSql.AppendLine("            AND A.JUNPYOID IN(" + sArrJunpyoID + ") ");
                    //strSql.AppendLine("       ) X1  ");
                    //strSql.AppendLine("  LEFT JOIN JAJAE X2 ");
                    //strSql.AppendLine("    ON X1.J_SERIAL = X2.J_SERIAL ");

                    //DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    //sPDAT1 = dtResult.Rows[0]["PDAT1"]?.ToString();
                    //sPDAT2 = dtResult.Rows[0]["PDAT2"]?.ToString();
                    //sCVCOD = dtResult.Rows[0]["CVCOD"]?.ToString();
                    //sRK = dtResult.Rows[0]["RK"]?.ToString();
                    //sMMDD1 = dtResult.Rows[0]["MMDD1"]?.ToString();
                    //sITNM1 = dtResult.Rows[0]["ITNM1"]?.ToString();
                    //sITNM2 = dtResult.Rows[0]["ITNM2"]?.ToString();
                    //sTWGT1 = dtResult.Rows[0]["TWGT1"]?.ToString();
                    //sCOST1 = dtResult.Rows[0]["COST1"]?.ToString();
                    //sTAMT1 = dtResult.Rows[0]["TAMT1"]?.ToString();
                    //sTVAT1 = dtResult.Rows[0]["TVAT1"]?.ToString();
                    #endregion

                    //#0003
                    strSql.Clear();
                    strSql.AppendLine(" WITH TEMP1 AS (                                           ");
                    if (!string.IsNullOrEmpty(sArrJunpyoID))
                    {
                        strSql.AppendLine("     SELECT X1.PDAT1 AS PDAT1                          ");
                        strSql.AppendLine("          , X1.PDAT2 AS PDAT2                          ");
                        strSql.AppendLine("          , X1.CVCOD                                   ");
                        strSql.AppendLine("          , X2.GUBUN1 AS ITNM1                         ");
                        strSql.AppendLine("          , X1.CNT AS ITNM2                            ");
                        strSql.AppendLine("          , X1.WGT AS TWGT1                            ");
                        strSql.AppendLine("          , X1.DANGA AS COST1                          ");
                        strSql.AppendLine("          , X1.AMT AS TAMT1                            ");
                        strSql.AppendLine("          , X1.VAT AS TVAT1                            ");
                        strSql.AppendLine("       FROM(                                           ");
                        strSql.AppendLine("              SELECT MAX(A.J_SERIAL) AS J_SERIAL       ");
                        strSql.AppendLine("                   , MAX(A.J_ID1) AS CVCOD             ");
                        strSql.AppendLine("                   , COUNT(*) -1 AS CNT                ");
                        strSql.AppendLine("                  , SUM(A.DANJUNG) AS WGT              ");
                        strSql.AppendLine("                  , ROUND(AVG(A.DANGA), 0, 1) AS DANGA ");
                        strSql.AppendLine("                    , SUM(A.KONGKEP) AS AMT            ");
                        strSql.AppendLine("                    , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END) AS VAT");
                        strSql.AppendLine("                    , MIN(A.J_DATE) AS PDAT1       ");
                        strSql.AppendLine("                    , MAX(A.J_DATE) AS PDAT2       ");
                        strSql.AppendLine("                FROM INLIST A                      ");
                        strSql.AppendLine("                LEFT JOIN MESURING B               ");
                        strSql.AppendLine("                  ON A.J_ID = B.IPCHULGO_MACHULID  ");
                        strSql.AppendLine("               WHERE A.J_ID1 = " + sDealerCd + "          ");
                        strSql.AppendLine("                 AND A.JUNPYOID IN(" + sArrJunpyoID + ")               ");
                        strSql.AppendLine("            ) X1                                   ");
                        strSql.AppendLine("       LEFT JOIN JAJAE X2                          ");
                        strSql.AppendLine("         ON X1.J_SERIAL = X2.J_SERIAL              ");
                    }
                    
                    if(!string.IsNullOrEmpty(sArrJunpyoID) && !string.IsNullOrEmpty(sJson))
                    {
                        strSql.AppendLine("      UNION ALL                                    ");
                    }

                    if (!string.IsNullOrEmpty(sJson))
                    {
                        strSql.AppendLine("     SELECT X1.PDAT1 AS PDAT1                                ");
                        strSql.AppendLine("          , X1.PDAT2 AS PDAT2                                ");
                        strSql.AppendLine("          , X1.CVCOD                                         ");
                        strSql.AppendLine("          , X1.RK AS ITNM1                                   ");
                        strSql.AppendLine("          , X1.CNT AS ITNM2                                  ");
                        strSql.AppendLine("          , X1.WGT AS TWGT1                                  ");
                        strSql.AppendLine("          , X1.DANGA AS COST1                                ");
                        strSql.AppendLine("          , X1.AMT AS TAMT1                                  ");
                        strSql.AppendLine("          , X1.VAT AS TVAT1                                  ");
                        strSql.AppendLine("       FROM(SELECT MIN(CONVERT(DATE, A1.TDATE)) AS PDAT1     ");
                        strSql.AppendLine("                   , MAX(CONVERT(DATE, A1.TDATE)) AS PDAT2   ");
                        strSql.AppendLine("                   , MAX(A1.CVCOD) AS CVCOD                  ");
                        strSql.AppendLine("                   , '인센티브' AS RK                        ");
                        strSql.AppendLine("                   , COUNT(*) - 1 AS CNT                     ");
                        strSql.AppendLine("                   , 0 AS WGT                                ");
                        strSql.AppendLine("                   , 0 AS DANGA                              ");
                        strSql.AppendLine("                   , SUM(A1.AMT) AS AMT                      ");
                        strSql.AppendLine("                   , SUM(A1.VAT) AS VAT                      ");
                        strSql.AppendLine("                FROM(SELECT Z1.TDATE                         ");
                        strSql.AppendLine("                           , Z1.ATGUB                        ");
                        strSql.AppendLine("                           , Z1.SEQNO                        ");
                        strSql.AppendLine("                           , Z1.CVCOD                        ");
                        strSql.AppendLine("                           , SUM(Z1.AMT) AS AMT              ");
                        strSql.AppendLine("                           , SUM(Z1.VAT) AS VAT              ");
                        strSql.AppendLine("                        FROM(SELECT A1.TDATE                 ");
                        strSql.AppendLine("                                    , A1.ATGUB               ");
                        strSql.AppendLine("                                    , A1.SEQNO               ");
                        strSql.AppendLine("                                    , A2.CVCOD               ");
                        strSql.AppendLine("                                    , CASE WHEN A2.ACCOD = '0404' THEN A2.ADAMT ELSE 0 END AS AMT");
                        strSql.AppendLine("                                    , CASE WHEN A2.ACCOD = '0255' THEN A2.ADAMT ELSE 0 END AS VAT");
                        strSql.AppendLine("                                 FROM(SELECT TDATE                                               ");
                        strSql.AppendLine("                                             , ATGUB                                             ");
                        strSql.AppendLine("                                             , SEQNO                                             ");
                        strSql.AppendLine("                                          FROM OPENJSON('"+sJson+"')                 ");
                        strSql.AppendLine("                                          WITH(TDATE VARCHAR(10) '$.TDATE'           ");
                        strSql.AppendLine("                                               , ATGUB VARCHAR(5) '$.ATGUB'          ");
                        strSql.AppendLine("                                               , SEQNO NUMERIC(10, 0) '$.SEQNO'))A1  ");
                        strSql.AppendLine("                                 LEFT JOIN ACTRAN A2                                 ");
                        strSql.AppendLine("                                   ON A1.TDATE = A2.TDATE                            ");
                        strSql.AppendLine("                                  AND A1.ATGUB = A2.ATGUB                            ");
                        strSql.AppendLine("                                  AND A1.SEQNO = A2.SEQNO                            ");
                        strSql.AppendLine("                                WHERE A2.ACCOD IN('0404', '0255'))Z1                 ");
                        strSql.AppendLine("                       GROUP BY Z1.TDATE, Z1.ATGUB, Z1.SEQNO, Z1.CVCOD) A1) X1       ");
                    }
                    
                    strSql.AppendLine(" )                                                                 ");
                    strSql.AppendLine("                                                                   ");
                    strSql.AppendLine(" SELECT MIN(Z1.PDAT1) AS PDAT1                                     ");
                    strSql.AppendLine("      , MAX(Z1.PDAT2) AS PDAT2                                     ");
                    strSql.AppendLine("      , MAX(Z1.CVCOD) AS CVCOD                                     ");
                    strSql.AppendLine("      , CONCAT('거래기간 : ', MIN(Z1.PDAT1), ' ~ ', MAX(Z1.PDAT2)) AS RK ");
                    strSql.AppendLine("      , SUBSTRING(CONVERT(VARCHAR(8), GETDATE(), 112), 5, 4) AS MMDD1    ");
                    strSql.AppendLine("      , STUFF((SELECT ',' + ITNM1                                        ");
                    strSql.AppendLine("                 FROM TEMP1                                              ");
                    strSql.AppendLine("                  FOR XML PATH('')),1,1,'') AS ITNM1                     ");
                    strSql.AppendLine("      , SUM(Z1.ITNM2) AS ITNM2                                             ");
                    strSql.AppendLine("      , SUM(Z1.TWGT1) AS TWGT1                                             ");
                    strSql.AppendLine("      , SUM(Z1.COST1) AS COST1                                           ");
                    strSql.AppendLine("      , SUM(Z1.TAMT1) AS TAMT1                                             ");
                    strSql.AppendLine("      , SUM(Z1.TVAT1) AS TVAT1                                             ");
                    strSql.AppendLine("   FROM TEMP1 Z1                                                         ");

                    DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    sPDAT1 = dtResult.Rows[0]["PDAT1"]?.ToString().Substring(0,10);
                    sPDAT2 = dtResult.Rows[0]["PDAT2"]?.ToString().Substring(0,10);
                    sCVCOD = dtResult.Rows[0]["CVCOD"]?.ToString();
                    sRK = dtResult.Rows[0]["RK"]?.ToString();
                    sMMDD1 = dtResult.Rows[0]["MMDD1"]?.ToString();
                    sITNM1 = dtResult.Rows[0]["ITNM1"]?.ToString() + " 외 " + dtResult.Rows[0]["ITNM2"]?.ToString() + " 건";
                    sTWGT1 = dtResult.Rows[0]["TWGT1"]?.ToString();
                    sCOST1 = dtResult.Rows[0]["COST1"]?.ToString();
                    sTAMT1 = dtResult.Rows[0]["TAMT1"]?.ToString();
                    sTVAT1 = dtResult.Rows[0]["TVAT1"]?.ToString();

                    //TAXNO 채번
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" SELECT CONCAT('TX', DATE_FORMAT(@TDATE, '%Y%m%d') , LPAD(IFNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1, 5, '0')) AS MAX_VAL ");
                    //strSql.AppendLine("   FROM TAXF A  ");
                    //strSql.AppendLine("  WHERE A.TDATE = @TDATE ");
                    #endregion

                    strSql.AppendLine("SELECT CONCAT('TX', CONVERT(VARCHAR(8), CONVERT(DATE, @TDATE), 112)                    ");
                    strSql.AppendLine("               , CAST(REPLICATE(0, 5 - LEN(ISNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1)) AS VARCHAR) ");
                    strSql.AppendLine("               + CAST(ISNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1 AS VARCHAR)) AS MAX_VAL            ");
                    strSql.AppendLine("  FROM TAXF A                                                                                ");
                    strSql.AppendLine(" WHERE A.TDATE = @TDATE                                                                ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@TDATE", sPDAT2);
                    string sTAXNO = cmd.ExecuteScalar()?.ToString();

                    //TAXF INSERT
                    dicParams.Clear();
                    dicParams.Add("TAXNO", sTAXNO);
                    dicParams.Add("TDATE", sPDAT2);
                    dicParams.Add("PDAT1", sPDAT1);
                    dicParams.Add("PDAT2", sPDAT2);
                    dicParams.Add("CVCOD", sCVCOD);
                    dicParams.Add("TGUBN", "1"); //HARDCODING -> COM_BASE_CD / CD_GB : AC18001_01 참조
                    dicParams.Add("PAYGB", "5"); //HARDCODING -> COM_BASE_CD / CD_GB : AC18001_02 참조
                    dicParams.Add("TAXGU", "S"); //HARDCODING -> COM_BASE_CD / CD_GB : AC18001_03 참조
                    dicParams.Add("RK", sRK);
                    dicParams.Add("MMDD1", sMMDD1);
                    dicParams.Add("ITNM1", sITNM1);
                    dicParams.Add("TWGT1", sTWGT1);
                    dicParams.Add("COST1", sCOST1);
                    dicParams.Add("TAMT1", sTAMT1);
                    dicParams.Add("TVAT1", sTVAT1);
                    dicParams.Add("CUSER", FmMainToolBar2.UserID);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO TAXF ");
                    strSql.AppendLine("           ( TAXNO ");
                    strSql.AppendLine("           , TDATE ");
                    strSql.AppendLine("           , PDAT1 ");
                    strSql.AppendLine("           , PDAT2 ");
                    strSql.AppendLine("           , CVCOD ");
                    strSql.AppendLine("           , TGUBN ");
                    strSql.AppendLine("           , PAYGB ");
                    strSql.AppendLine("           , TAXGU ");
                    strSql.AppendLine("           , AUTGB ");
                    strSql.AppendLine("           , RK ");
                    strSql.AppendLine("           , MMDD1 ");
                    strSql.AppendLine("           , ITNM1 ");
                    strSql.AppendLine("           , TWGT1 ");
                    strSql.AppendLine("           , COST1 ");
                    strSql.AppendLine("           , TAMT1 ");
                    strSql.AppendLine("           , TVAT1 ");
                    strSql.AppendLine("           , CDATE ");
                    strSql.AppendLine("           , CUSER )");
                    strSql.AppendLine("     VALUES( @TAXNO ");
                    strSql.AppendLine("           , @TDATE ");
                    strSql.AppendLine("           , @PDAT1 ");
                    strSql.AppendLine("           , @PDAT2 ");
                    strSql.AppendLine("           , @CVCOD ");
                    strSql.AppendLine("           , @TGUBN ");
                    strSql.AppendLine("           , @PAYGB ");
                    strSql.AppendLine("           , @TAXGU ");
                    strSql.AppendLine("           , '1' "); //HARDCODING -> COM_BASE_CD(CD_GB : AC18001_04) 
                    strSql.AppendLine("           , @RK ");
                    strSql.AppendLine("           , @MMDD1 ");
                    strSql.AppendLine("           , @ITNM1 ");
                    strSql.AppendLine("           , @TWGT1 ");
                    strSql.AppendLine("           , @COST1 ");
                    strSql.AppendLine("           , @TAMT1 ");
                    strSql.AppendLine("           , @TVAT1 ");
                    strSql.AppendLine("           , CONVERT(VARCHAR(20), GETDATE(), 20) ");
                    strSql.AppendLine("           , @CUSER )");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    foreach (KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    //INLIST 마감자료 TAXNO UPDATE 반영
                    if(sArrJunpyoIDs != null && sArrJunpyoIDs.Length > 0)
                    {
                        foreach (string str in sArrJunpyoIDs)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" UPDATE INLIST ");
                            strSql.AppendLine("    SET TAXNO = @TAXNO ");
                            strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@TAXNO", sTAXNO);
                            cmd.Parameters.AddWithValue("@JUNPYOID", str);
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                    }

                    //#0003 전표 TAXNO UPDATE 반영
                    if (!string.IsNullOrEmpty(sJson))
                    {
                        DataTable jsonTable = JsonConvert.DeserializeObject<DataTable>(sJson);

                        foreach(DataRow row in jsonTable.Rows)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" UPDATE ACTRAN     ");
                            strSql.AppendLine("    SET TAXNO = '"+ sTAXNO + "' ");
                            strSql.AppendLine("  WHERE TDATE = '" + row["TDATE"] + "' ");
                            strSql.AppendLine("    AND ATGUB = '"+ row["ATGUB"] + "' ");
                            strSql.AppendLine("    AND SEQNO =  "+ row["SEQNO"] + "  ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                BtnRetr.PerformClick();
                //TabControl.SelectedTabPage = TabPageTax;
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
        /// list 특정값 제거
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="filter"></param>
        /// <returns></returns>
        private string[] RemoveAt(string[] arr, string filter)
        {
            return arr.Where((value, i) => !value.Contains(filter)).ToArray();
        }

        private void IssuePurc()
        {
            //TEMP 컬럼에 값(INLIST 테이블의 PK값인 JUNPYOID) 이 전체 그리드에 하나도 존재하지 않는 경우
            //동시에 있는 값들은 Dictionary 포함
            Dictionary<string, string> dicDelaerInfo = new Dictionary<string, string>();
            int iCnt = 0;
            for (int i = 0; i < GridViewP1.RowCount; i++)
            {
                string[] sArr = GridViewP1.GetRowCellValue(i, GridColP1Temp)?.ToString().Replace(" ", "").Split(',');

                if (sArr.Length == 1)
                {
                    if (string.IsNullOrEmpty(sArr[0]))
                        continue;
                }

                if (sArr.Length > 0)
                {
                    dicDelaerInfo.Add(GridViewP1.GetRowCellValue(i, GridColP1CvCod)?.ToString(), GridViewP1.GetRowCellValue(i, GridColP1Temp)?.ToString().Replace(" ", ""));
                    iCnt++;
                }
            }

            if (iCnt == 0)
            {
                XtraMessageBox.Show("리스트의 어떠한 마감자료도 선택되지 않았습니다.\r\n발행하려는 마감자료를 선택해주세요.");
                return;
            }

            Int32 iChk = 0;
            for (int i = 0; i < GridViewP1.RowCount; i++)
            {
                string sChk = GridViewP1.GetRowCellValue(i, GridColP1Chk)?.ToString();
                if (string.IsNullOrEmpty(sChk))
                    return;

                if (sChk.Equals("Y"))
                    iChk++;
            }

            string sMSG = string.Format("선택된 건들에 대하여 일괄저장을 진행하시겠습니까?");
            if (XtraMessageBox.Show(sMSG, "계산서 발행여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                string sDealerCd = string.Empty;
                string sPDAT1 = string.Empty;
                string sPDAT2 = string.Empty;
                string sCVCOD = string.Empty;
                string sRK = string.Empty;
                string sMMDD1 = string.Empty;
                string sITNM1 = string.Empty;
                string sTWGT1 = string.Empty;
                string sCOST1 = string.Empty;
                string sTAMT1 = string.Empty;
                string sTVAT1 = string.Empty;

                foreach (KeyValuePair<string, string> dic in dicDelaerInfo)
                {
                    sDealerCd = string.Empty;
                    sPDAT1 = string.Empty;
                    sPDAT2 = string.Empty;
                    sCVCOD = string.Empty;
                    sRK = string.Empty;
                    sMMDD1 = string.Empty;
                    sITNM1 = string.Empty;
                    sTWGT1 = string.Empty;
                    sCOST1 = string.Empty;
                    sTAMT1 = string.Empty;
                    sTVAT1 = string.Empty;

                    sDealerCd = dic.Key;
                    string[] sArrJunpyoIDs = dic.Value.Replace(" ", "").Split(',');

                    ////TAXNO 채번
                    //strSql.Clear();
                    //strSql.AppendLine(" ");
                    //strSql.AppendLine(" SELECT CONCAT('TX', DATE_FORMAT(NOW(), '%Y%m%d') , LPAD(IFNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1, 5, '0')) AS MAX_VAL ");
                    //strSql.AppendLine("   FROM TAXF A  ");
                    //strSql.AppendLine("  WHERE A.TDATE = @TDATE ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //cmd.Parameters.AddWithValue("@TDATE", DateTime.Today.ToString("yyyy-MM-dd"));
                    //string sTAXNO = cmd.ExecuteScalar()?.ToString();

                    //TAXF INSERT 위한 내역 조회
                    dicParams.Clear();
                    dicParams.Add("DEALER_CD", sDealerCd);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" SELECT DATE_FORMAT(X1.PDAT1, '%Y-%m-%d') AS PDAT1 ");
                    //strSql.AppendLine("      , DATE_FORMAT(X1.PDAT2, '%Y-%m-%d') AS PDAT2 ");
                    //strSql.AppendLine("      , X1.CVCOD ");
                    //strSql.AppendLine("      , CONCAT('거래기간 : ', X1.PDAT1, ' ~ ', X1.PDAT2) AS RK ");
                    //strSql.AppendLine("      , DATE_FORMAT(NOW(), '%m%d') AS MMDD1 ");
                    //strSql.AppendLine("      , CONCAT(X2.GUBUN1, ' 외 ', X1.CNT, ' 건') AS ITNM1 ");
                    //strSql.AppendLine("      , X1.WGT AS TWGT1 ");
                    //strSql.AppendLine("      , X1.DANGA AS COST1 ");
                    //strSql.AppendLine("      , X1.AMT AS TAMT1 ");
                    //strSql.AppendLine("      , X1.VAT AS TVAT1 ");
                    //strSql.AppendLine("   FROM ( ");
                    //strSql.AppendLine("          SELECT MAX(A.J_SERIAL) AS J_SERIAL ");
                    //strSql.AppendLine("               , MAX(A.J_ID1) AS CVCOD ");
                    //strSql.AppendLine("               , COUNT(*) - 1 AS CNT ");
                    //strSql.AppendLine("               , SUM(A.DANJUNG) AS WGT ");
                    //strSql.AppendLine("               , TRUNCATE(AVG(A.DANGA), 1) AS DANGA ");
                    //strSql.AppendLine("               , SUM(A.IKONGKEP) AS AMT  ");
                    //strSql.AppendLine("               , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END) AS VAT ");
                    //strSql.AppendLine("               , MIN(A.J_DATE) AS PDAT1 ");
                    //strSql.AppendLine("               , MAX(A.J_DATE) AS PDAT2 ");
                    //strSql.AppendLine("            FROM INLIST A ");
                    //strSql.AppendLine("            LEFT JOIN MESURING B  ");
                    //strSql.AppendLine("              ON A.J_ID = B.IPCHULGO_MAIPID ");
                    //strSql.AppendLine("           WHERE A.J_ID1 = " + sDealerCd + "");
                    //strSql.AppendLine("             AND A.JUNPYOID IN (" + dic.Value + ") ");
                    //strSql.AppendLine("        ) X1  ");
                    //strSql.AppendLine("   LEFT JOIN JAJAE X2 ");
                    //strSql.AppendLine("     ON X1.J_SERIAL = X2.J_SERIAL ");
                    #endregion

                    strSql.AppendLine("SELECT X1.PDAT1 AS PDAT1 ");
                    strSql.AppendLine("     , X1.PDAT2 AS PDAT2 ");
                    strSql.AppendLine("     , X1.CVCOD ");
                    strSql.AppendLine("     , CONCAT('거래기간 : ', X1.PDAT1, ' ~ ', X1.PDAT2) AS RK ");
                    strSql.AppendLine("     , SUBSTRING(CONVERT(VARCHAR(8), GETDATE(), 112),5, 4) AS MMDD1 ");
                    strSql.AppendLine("     , CONCAT(X2.GUBUN1, ' 외 ', X1.CNT, ' 건') AS ITNM1 ");
                    strSql.AppendLine("     , X1.WGT AS TWGT1 ");
                    strSql.AppendLine("     , X1.DANGA AS COST1 ");
                    strSql.AppendLine("     , X1.AMT AS TAMT1 ");
                    strSql.AppendLine("     , X1.VAT AS TVAT1 ");
                    strSql.AppendLine("  FROM(");
                    strSql.AppendLine("         SELECT MAX(A.J_SERIAL) AS J_SERIAL ");
                    strSql.AppendLine("              , MAX(A.J_ID1) AS CVCOD ");
                    strSql.AppendLine("              , COUNT(*) - 1 AS CNT ");
                    strSql.AppendLine("              , SUM(A.DANJUNG) AS WGT ");
                    strSql.AppendLine("              , ROUND(AVG(A.DANGA),0,1) AS DANGA ");
                    strSql.AppendLine("              , SUM(A.IKONGKEP) AS AMT  ");
                    strSql.AppendLine("              , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END) AS VAT ");
                    strSql.AppendLine("              , MIN(A.J_DATE) AS PDAT1 ");
                    strSql.AppendLine("              , MAX(A.J_DATE) AS PDAT2 ");
                    strSql.AppendLine("           FROM INLIST A ");
                    strSql.AppendLine("           LEFT JOIN MESURING B  ");
                    strSql.AppendLine("             ON A.J_ID = B.IPCHULGO_MAIPID ");
                    strSql.AppendLine("          WHERE A.J_ID1 = " + sDealerCd + "");
                    strSql.AppendLine("            AND A.JUNPYOID IN(" + dic.Value + ") ");
                    strSql.AppendLine("       ) X1  ");
                    strSql.AppendLine("  LEFT JOIN JAJAE X2 ");
                    strSql.AppendLine("    ON X1.J_SERIAL = X2.J_SERIAL ");

                    DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    sPDAT1 = dtResult.Rows[0]["PDAT1"]?.ToString();
                    sPDAT2 = dtResult.Rows[0]["PDAT2"]?.ToString();
                    sCVCOD = dtResult.Rows[0]["CVCOD"]?.ToString();
                    sRK = dtResult.Rows[0]["RK"]?.ToString();
                    sMMDD1 = dtResult.Rows[0]["MMDD1"]?.ToString();
                    sITNM1 = dtResult.Rows[0]["ITNM1"]?.ToString();
                    sTWGT1 = dtResult.Rows[0]["TWGT1"]?.ToString();
                    sCOST1 = dtResult.Rows[0]["COST1"]?.ToString();
                    sTAMT1 = dtResult.Rows[0]["TAMT1"]?.ToString();
                    sTVAT1 = dtResult.Rows[0]["TVAT1"]?.ToString();

                    //TAXNO 채번
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" SELECT CONCAT('TX', DATE_FORMAT(@TDATE, '%Y%m%d') , LPAD(IFNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1, 5, '0')) AS MAX_VAL ");
                    //strSql.AppendLine("   FROM TAXF A  ");
                    //strSql.AppendLine("  WHERE A.TDATE = @TDATE ");
                    #endregion

                    strSql.AppendLine("SELECT CONCAT('TX', CONVERT(VARCHAR(8), CONVERT(DATE, @TDATE), 112)                    ");
                    strSql.AppendLine("               , CAST(REPLICATE(0, 5 - LEN(ISNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1)) AS VARCHAR) ");
                    strSql.AppendLine("               + CAST(ISNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1 AS VARCHAR)) AS MAX_VAL            ");
                    strSql.AppendLine("  FROM TAXF A                                                                                ");
                    strSql.AppendLine(" WHERE A.TDATE = @TDATE                                                                ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@TDATE", sPDAT2);
                    string sTAXNO = cmd.ExecuteScalar()?.ToString();
                    
                    //TAXF INSERT
                    dicParams.Clear();
                    dicParams.Add("TAXNO", sTAXNO);
                    dicParams.Add("TDATE", sPDAT2);
                    dicParams.Add("PDAT1", sPDAT1);
                    dicParams.Add("PDAT2", sPDAT2);
                    dicParams.Add("CVCOD", sCVCOD);
                    dicParams.Add("TGUBN", "2"); //HARDCODING -> COM_BASE_CD / CD_GB : AC18001_01 참조
                    dicParams.Add("PAYGB", "4"); //HARDCODING -> COM_BASE_CD / CD_GB : AC18001_02 참조
                    dicParams.Add("TAXGU", "P"); //HARDCODING -> COM_BASE_CD / CD_GB : AC18001_03 참조
                    dicParams.Add("RK", sRK);
                    dicParams.Add("MMDD1", sMMDD1);
                    dicParams.Add("ITNM1", sITNM1);
                    dicParams.Add("TWGT1", sTWGT1);
                    dicParams.Add("COST1", sCOST1);
                    dicParams.Add("TAMT1", sTAMT1);
                    dicParams.Add("TVAT1", sTVAT1);
                    dicParams.Add("CUSER", FmMainToolBar2.UserID);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO TAXF ");
                    strSql.AppendLine("           ( TAXNO ");
                    strSql.AppendLine("           , TDATE ");
                    strSql.AppendLine("           , PDAT1 ");
                    strSql.AppendLine("           , PDAT2 ");
                    strSql.AppendLine("           , CVCOD ");
                    strSql.AppendLine("           , TGUBN ");
                    strSql.AppendLine("           , PAYGB ");
                    strSql.AppendLine("           , TAXGU ");
                    strSql.AppendLine("           , AUTGB ");
                    strSql.AppendLine("           , RK ");
                    strSql.AppendLine("           , MMDD1 ");
                    strSql.AppendLine("           , ITNM1 ");
                    strSql.AppendLine("           , TWGT1 ");
                    strSql.AppendLine("           , COST1 ");
                    strSql.AppendLine("           , TAMT1 ");
                    strSql.AppendLine("           , TVAT1 ");
                    strSql.AppendLine("           , CDATE ");
                    strSql.AppendLine("           , CUSER )");
                    strSql.AppendLine("     VALUES( @TAXNO ");
                    strSql.AppendLine("           , @TDATE ");
                    strSql.AppendLine("           , @PDAT1 ");
                    strSql.AppendLine("           , @PDAT2 ");
                    strSql.AppendLine("           , @CVCOD ");
                    strSql.AppendLine("           , @TGUBN ");
                    strSql.AppendLine("           , @PAYGB ");
                    strSql.AppendLine("           , @TAXGU ");
                    strSql.AppendLine("           , '1' "); //HARDCODING -> COM_BASE_CD(CD_GB : AC18001_04) 
                    strSql.AppendLine("           , @RK ");
                    strSql.AppendLine("           , @MMDD1 ");
                    strSql.AppendLine("           , @ITNM1 ");
                    strSql.AppendLine("           , @TWGT1 ");
                    strSql.AppendLine("           , @COST1 ");
                    strSql.AppendLine("           , @TAMT1 ");
                    strSql.AppendLine("           , @TVAT1 ");
                    strSql.AppendLine("           , CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine("           , @CUSER )");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    foreach (KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    //INLIST 마감자료 TAXNO UPDATE 반영
                    foreach (string str in sArrJunpyoIDs)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET TAXNO = @TAXNO ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@TAXNO", sTAXNO);
                        cmd.Parameters.AddWithValue("@JUNPYOID", str);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                }

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                BtnRetr.PerformClick();
                //TabControl.SelectedTabPage = TabPageTax;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (TabControl.SelectedTabPage == TabPageTax)
            {
                ComnEtcFunc.ExportExcelFile(string.Format("{0}_", this.Text), GridRetrL1);
            }
            else if(TabControl.SelectedTabPage == TabPageIssueSale)
            {

            }
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("PRINT", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 출력 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (GridViewL1.RowCount == 0)
            {
                return;
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("TAXNO", GridViewL1.GetFocusedRowCellValue(GridColL1TaxNo)?.ToString());

            string sTaxGu = GridViewL1.GetFocusedRowCellValue(GridColL1TaxGu1)?.ToString();

            ReportViewer fm = new ReportViewer(GetReportData(dicParams, sTaxGu), "RptTaxBill");
            fm.ShowDialog();
        }

        private DataTable GetReportData(Dictionary<string, string> dicParams, string sTaxGu)
        {
            StringBuilder strSql = new StringBuilder();

            if (sTaxGu.Equals("P"))
            {
                #region mariaDB
                //strSql.AppendLine(" SELECT CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-',SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS PV_IDT_NO ");
                //strSql.AppendLine("      , B.DEALER_NM AS PV_CVNAM ");
                //strSql.AppendLine("      , B.REP_NM AS PV_REP_NM ");
                //strSql.AppendLine("      , CASE WHEN B.ADDR IS NULL OR REPLACE(B.ADDR, ' ', '') = '' THEN B.DTL_ADDR ELSE B.ADDR END AS PV_ADDR ");
                //strSql.AppendLine("      , B.BIZ_NM AS PV_BIZ_NM ");
                //strSql.AppendLine("      , B.TYPE_NM AS PV_TYPE_NM ");
                //strSql.AppendLine("      , C.SANO AS BY_IDT_NO ");
                //strSql.AppendLine("      , C.COMNM AS BY_CVNAM ");
                //strSql.AppendLine("      , C.OWNAM AS BY_REP_NM ");
                //strSql.AppendLine("      , C.ADDR1 AS BY_ADDR ");
                //strSql.AppendLine("      , C.UPTAE AS BY_BIZ_NM ");
                //strSql.AppendLine("      , C.JONGK AS BY_TYPE_NM ");
                //strSql.AppendLine("      , A.TDATE ");
                //strSql.AppendLine("      , A.RK ");
                //strSql.AppendLine("      , A.PAYGB ");
                //strSql.AppendLine("      , FORMAT(IFNULL(A.TAMT1, 0) + IFNULL(A.TAMT2, 0) + IFNULL(A.TAMT3, 0) + IFNULL(A.TAMT4, 0), 0) AS TOT_AMT ");
                //strSql.AppendLine("      , FORMAT(IFNULL(A.TVAT1, 0) + IFNULL(A.TVAT2, 0) + IFNULL(A.TVAT3, 0) + IFNULL(A.TVAT4, 0), 0) AS TOT_VAT ");
                //strSql.AppendLine("      , FORMAT(IFNULL(A.TAMT1, 0) + IFNULL(A.TAMT2, 0) + IFNULL(A.TAMT3, 0) + IFNULL(A.TAMT4, 0) +  ");
                //strSql.AppendLine("        IFNULL(A.TVAT1, 0) + IFNULL(A.TVAT2, 0) + IFNULL(A.TVAT3, 0) + IFNULL(A.TVAT4, 0), 0) AS SUM_AMT ");
                //strSql.AppendLine("      , LEFT(A.MMDD1, 2) AS MM1, RIGHT(A.MMDD1, 2) AS DD1 ");
                //strSql.AppendLine("      , A.ITNM1, A.SPEC1, A.TWGT1, A.COST1, FORMAT(A.TAMT1, 0) AS TAMT1, FORMAT(A.TVAT1, 0) AS TVAT1");
                //strSql.AppendLine("      , LEFT(A.MMDD2, 2) AS MM2, RIGHT(A.MMDD2, 2) AS DD2 ");
                //strSql.AppendLine("      , A.ITNM2, A.SPEC2, A.TWGT2, A.COST2, FORMAT(A.TAMT2, 0) AS TAMT2, FORMAT(A.TVAT2, 0) AS TVAT2 ");
                //strSql.AppendLine("      , LEFT(A.MMDD3, 2) AS MM3, RIGHT(A.MMDD3, 2) AS DD3 ");
                //strSql.AppendLine("      , A.ITNM3, A.SPEC3, A.TWGT3, A.COST3, FORMAT(A.TAMT3, 0) AS TAMT3, FORMAT(A.TVAT3, 0) AS TVAT3 ");
                //strSql.AppendLine("      , LEFT(A.MMDD4, 2) AS MM4, RIGHT(A.MMDD4, 2) AS DD4 ");
                //strSql.AppendLine("      , A.ITNM4, A.SPEC4, A.TWGT4, A.COST4, FORMAT(A.TAMT4, 0) AS TAMT4, FORMAT(A.TVAT4, 0) AS TVAT4 ");
                //strSql.AppendLine("   FROM TAXF A  ");
                //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B  ");
                //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
                //strSql.AppendLine("   LEFT JOIN ( SELECT COMNM, SANO, OWNAM, ADDR1, UPTAE, JONGK ");
                //strSql.AppendLine("                 FROM COMPANYINFO X1 ) C  ");
                //strSql.AppendLine("     ON 1 = 1 ");
                //strSql.AppendLine("  WHERE A.TAXNO = @TAXNO ");
                #endregion

                strSql.AppendLine("SELECT CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-',SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS PV_IDT_NO ");
                strSql.AppendLine("     , B.DEALER_NM AS PV_CVNAM ");
                strSql.AppendLine("     , B.REP_NM AS PV_REP_NM ");
                strSql.AppendLine("     , CASE WHEN B.ADDR IS NULL OR REPLACE(B.ADDR, ' ', '') = '' THEN B.DTL_ADDR ELSE B.ADDR END AS PV_ADDR ");
                strSql.AppendLine("     , B.BIZ_NM AS PV_BIZ_NM ");
                strSql.AppendLine("     , B.TYPE_NM AS PV_TYPE_NM ");
                strSql.AppendLine("     , C.SANO AS BY_IDT_NO ");
                strSql.AppendLine("     , C.COMNM AS BY_CVNAM ");
                strSql.AppendLine("     , C.OWNAM AS BY_REP_NM ");
                strSql.AppendLine("     , C.ADDR1 AS BY_ADDR ");
                strSql.AppendLine("     , C.UPTAE AS BY_BIZ_NM ");
                strSql.AppendLine("     , C.JONGK AS BY_TYPE_NM ");
                strSql.AppendLine("     , A.TDATE ");
                strSql.AppendLine("     , A.RK ");
                strSql.AppendLine("     , A.PAYGB ");
                strSql.AppendLine("     , FORMAT(ISNULL(A.TAMT1, 0) + ISNULL(A.TAMT2, 0) + ISNULL(A.TAMT3, 0) + ISNULL(A.TAMT4, 0), '#,0') AS TOT_AMT ");
                strSql.AppendLine("     , FORMAT(ISNULL(A.TVAT1, 0) + ISNULL(A.TVAT2, 0) + ISNULL(A.TVAT3, 0) + ISNULL(A.TVAT4, 0), '#,0') AS TOT_VAT ");
                strSql.AppendLine("     , FORMAT(ISNULL(A.TAMT1, 0) + ISNULL(A.TAMT2, 0) + ISNULL(A.TAMT3, 0) + ISNULL(A.TAMT4, 0) + ");
                strSql.AppendLine("       ISNULL(A.TVAT1, 0) + ISNULL(A.TVAT2, 0) + ISNULL(A.TVAT3, 0) + ISNULL(A.TVAT4, 0), '#,0') AS SUM_AMT ");
                strSql.AppendLine("     , LEFT(A.MMDD1, 2) AS MM1, RIGHT(A.MMDD1, 2) AS DD1 ");
                strSql.AppendLine("     , A.ITNM1, A.SPEC1, FORMAT(A.TWGT1,'#,0') AS TWGT1, FORMAT(A.COST1,'#,0') AS COST1, FORMAT(A.TAMT1, '#,0') AS TAMT1, FORMAT(A.TVAT1, '#,0') AS TVAT1");
                strSql.AppendLine("     , LEFT(A.MMDD2, 2) AS MM2, RIGHT(A.MMDD2, 2) AS DD2 ");
                strSql.AppendLine("     , A.ITNM2, A.SPEC2, FORMAT(A.TWGT2,'#,0') AS TWGT2, FORMAT(A.COST2,'#,0') AS COST2, FORMAT(A.TAMT2, '#,0') AS TAMT2, FORMAT(A.TVAT2, '#,0') AS TVAT2 ");
                strSql.AppendLine("     , LEFT(A.MMDD3, 2) AS MM3, RIGHT(A.MMDD3, 2) AS DD3 ");
                strSql.AppendLine("     , A.ITNM3, A.SPEC3, FORMAT(A.TWGT3,'#,0') AS TWGT3, FORMAT(A.COST3,'#,0') AS COST3, FORMAT(A.TAMT3, '#,0') AS TAMT3, FORMAT(A.TVAT3, '#,0') AS TVAT3 ");
                strSql.AppendLine("     , LEFT(A.MMDD4, 2) AS MM4, RIGHT(A.MMDD4, 2) AS DD4 ");
                strSql.AppendLine("     , A.ITNM4, A.SPEC4, FORMAT(A.TWGT4,'#,0') AS TWGT4, FORMAT(A.COST4,'#,0') AS COST4, FORMAT(A.TAMT4, '#,0') AS TAMT4, FORMAT(A.TVAT4, '#,0') AS TVAT4 ");
                strSql.AppendLine("  FROM TAXF A  ");
                strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD B  ");
                strSql.AppendLine("    ON A.CVCOD = B.DEALER_CD ");
                strSql.AppendLine("  LEFT JOIN(SELECT COMNM, SANO, OWNAM, ADDR1, UPTAE, JONGK ");
                strSql.AppendLine("                FROM COMPANYINFO X1 ) C  ");
                strSql.AppendLine("    ON 1 = 1 ");
                strSql.AppendLine(" WHERE A.TAXNO = @TAXNO ");
            }
            else
            {
                #region mariaDB
                //strSql.AppendLine(" SELECT CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-',SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS BY_IDT_NO ");
                //strSql.AppendLine("      , B.DEALER_NM AS BY_CVNAM ");
                //strSql.AppendLine("      , B.REP_NM AS BY_REP_NM ");
                //strSql.AppendLine("      , CASE WHEN B.ADDR IS NULL OR REPLACE(B.ADDR, ' ', '') = '' THEN B.DTL_ADDR ELSE B.ADDR END AS BY_ADDR ");
                //strSql.AppendLine("      , B.BIZ_NM AS BY_BIZ_NM ");
                //strSql.AppendLine("      , B.TYPE_NM AS BY_TYPE_NM ");
                //strSql.AppendLine("      , C.SANO AS PV_IDT_NO ");
                //strSql.AppendLine("      , C.COMNM AS PV_CVNAM ");
                //strSql.AppendLine("      , C.OWNAM AS PV_REP_NM ");
                //strSql.AppendLine("      , C.ADDR1 AS PV_ADDR ");
                //strSql.AppendLine("      , C.UPTAE AS PV_BIZ_NM ");
                //strSql.AppendLine("      , C.JONGK AS PV_TYPE_NM ");
                //strSql.AppendLine("      , A.TDATE ");
                //strSql.AppendLine("      , A.RK ");
                //strSql.AppendLine("      , A.PAYGB ");
                //strSql.AppendLine("      , FORMAT(IFNULL(A.TAMT1, 0) + IFNULL(A.TAMT2, 0) + IFNULL(A.TAMT3, 0) + IFNULL(A.TAMT4, 0), 0) AS TOT_AMT ");
                //strSql.AppendLine("      , FORMAT(IFNULL(A.TVAT1, 0) + IFNULL(A.TVAT2, 0) + IFNULL(A.TVAT3, 0) + IFNULL(A.TVAT4, 0), 0) AS TOT_VAT ");
                //strSql.AppendLine("      , FORMAT(IFNULL(A.TAMT1, 0) + IFNULL(A.TAMT2, 0) + IFNULL(A.TAMT3, 0) + IFNULL(A.TAMT4, 0) +  ");
                //strSql.AppendLine("        IFNULL(A.TVAT1, 0) + IFNULL(A.TVAT2, 0) + IFNULL(A.TVAT3, 0) + IFNULL(A.TVAT4, 0), 0) AS SUM_AMT ");
                //strSql.AppendLine("      , LEFT(A.MMDD1, 2) AS MM1, RIGHT(A.MMDD1, 2) AS DD1 ");
                //strSql.AppendLine("      , A.ITNM1, A.SPEC1, A.TWGT1, A.COST1, FORMAT(A.TAMT1, 0) AS TAMT1, FORMAT(A.TVAT1, 0) AS TVAT1");
                //strSql.AppendLine("      , LEFT(A.MMDD2, 2) AS MM2, RIGHT(A.MMDD2, 2) AS DD2 ");
                //strSql.AppendLine("      , A.ITNM2, A.SPEC2, A.TWGT2, A.COST2, FORMAT(A.TAMT2, 0) AS TAMT2, FORMAT(A.TVAT2, 0) AS TVAT2 ");
                //strSql.AppendLine("      , LEFT(A.MMDD3, 2) AS MM3, RIGHT(A.MMDD3, 2) AS DD3 ");
                //strSql.AppendLine("      , A.ITNM3, A.SPEC3, A.TWGT3, A.COST3, FORMAT(A.TAMT3, 0) AS TAMT3, FORMAT(A.TVAT3, 0) AS TVAT3 ");
                //strSql.AppendLine("      , LEFT(A.MMDD4, 2) AS MM4, RIGHT(A.MMDD4, 2) AS DD4 ");
                //strSql.AppendLine("      , A.ITNM4, A.SPEC4, A.TWGT4, A.COST4, FORMAT(A.TAMT4, 0) AS TAMT4, FORMAT(A.TVAT4, 0) AS TVAT4 ");
                //strSql.AppendLine("   FROM TAXF A  ");
                //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B  ");
                //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
                //strSql.AppendLine("   LEFT JOIN ( SELECT COMNM, SANO, OWNAM, ADDR1, UPTAE, JONGK ");
                //strSql.AppendLine("                 FROM COMPANYINFO X1 ) C  ");
                //strSql.AppendLine("     ON 1 = 1 ");
                //strSql.AppendLine("  WHERE A.TAXNO = @TAXNO ");
                #endregion

                strSql.AppendLine("SELECT CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-',SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS BY_IDT_NO ");
                strSql.AppendLine("     , B.DEALER_NM AS BY_CVNAM ");
                strSql.AppendLine("     , B.REP_NM AS BY_REP_NM ");
                strSql.AppendLine("     , CASE WHEN B.ADDR IS NULL OR REPLACE(B.ADDR, ' ', '') = '' THEN B.DTL_ADDR ELSE B.ADDR END AS BY_ADDR ");
                strSql.AppendLine("     , B.BIZ_NM AS BY_BIZ_NM ");
                strSql.AppendLine("     , B.TYPE_NM AS BY_TYPE_NM ");
                strSql.AppendLine("     , C.SANO AS PV_IDT_NO ");
                strSql.AppendLine("     , C.COMNM AS PV_CVNAM ");
                strSql.AppendLine("     , C.OWNAM AS PV_REP_NM ");
                strSql.AppendLine("     , C.ADDR1 AS PV_ADDR ");
                strSql.AppendLine("     , C.UPTAE AS PV_BIZ_NM ");
                strSql.AppendLine("     , C.JONGK AS PV_TYPE_NM ");
                strSql.AppendLine("     , A.TDATE ");
                strSql.AppendLine("     , A.RK ");
                strSql.AppendLine("     , A.PAYGB ");
                strSql.AppendLine("     , FORMAT(ISNULL(A.TAMT1, 0) + ISNULL(A.TAMT2, 0) + ISNULL(A.TAMT3, 0) + ISNULL(A.TAMT4, 0), '#,0') AS TOT_AMT ");
                strSql.AppendLine("     , FORMAT(ISNULL(A.TVAT1, 0) + ISNULL(A.TVAT2, 0) + ISNULL(A.TVAT3, 0) + ISNULL(A.TVAT4, 0), '#,0') AS TOT_VAT ");
                strSql.AppendLine("     , FORMAT(ISNULL(A.TAMT1, 0) + ISNULL(A.TAMT2, 0) + ISNULL(A.TAMT3, 0) + ISNULL(A.TAMT4, 0) + ");
                strSql.AppendLine("       ISNULL(A.TVAT1, 0) + ISNULL(A.TVAT2, 0) + ISNULL(A.TVAT3, 0) + ISNULL(A.TVAT4, 0), '#,0') AS SUM_AMT ");
                strSql.AppendLine("     , LEFT(A.MMDD1, 2) AS MM1, RIGHT(A.MMDD1, 2) AS DD1 ");
                strSql.AppendLine("     , A.ITNM1, A.SPEC1, FORMAT(A.TWGT1,'#,0') AS TWGT1, FORMAT(A.COST1,'#,0') AS COST1, FORMAT(A.TAMT1, '#,0') AS TAMT1, FORMAT(A.TVAT1, '#,0') AS TVAT1");
                strSql.AppendLine("     , LEFT(A.MMDD2, 2) AS MM2, RIGHT(A.MMDD2, 2) AS DD2 ");
                strSql.AppendLine("     , A.ITNM2, A.SPEC2, FORMAT(A.TWGT2,'#,0') AS TWGT2, FORMAT(A.COST2,'#,0') AS COST2, FORMAT(A.TAMT2, '#,0') AS TAMT2, FORMAT(A.TVAT2, '#,0') AS TVAT2 ");
                strSql.AppendLine("     , LEFT(A.MMDD3, 2) AS MM3, RIGHT(A.MMDD3, 2) AS DD3 ");
                strSql.AppendLine("     , A.ITNM3, A.SPEC3, FORMAT(A.TWGT3,'#,0') AS TWGT3, FORMAT(A.COST3,'#,0') AS COST3, FORMAT(A.TAMT3, '#,0') AS TAMT3, FORMAT(A.TVAT3, '#,0') AS TVAT3 ");
                strSql.AppendLine("     , LEFT(A.MMDD4, 2) AS MM4, RIGHT(A.MMDD4, 2) AS DD4 ");
                strSql.AppendLine("     , A.ITNM4, A.SPEC4, FORMAT(A.TWGT4,'#,0') AS TWGT4, FORMAT(A.COST4,'#,0') AS COST4, FORMAT(A.TAMT4, '#,0') AS TAMT4, FORMAT(A.TVAT4, '#,0') AS TVAT4 ");
                strSql.AppendLine("  FROM TAXF A  ");
                strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD B  ");
                strSql.AppendLine("    ON A.CVCOD = B.DEALER_CD ");
                strSql.AppendLine("  LEFT JOIN(SELECT COMNM, SANO, OWNAM, ADDR1, UPTAE, JONGK ");
                strSql.AppendLine("                FROM COMPANYINFO X1 ) C  ");
                strSql.AppendLine("    ON 1 = 1 ");
                strSql.AppendLine(" WHERE A.TAXNO = @TAXNO ");

            }

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC18001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
            else if (e.KeyCode == Keys.F9)
                BtnPrint.PerformClick();
        }

        /// <summary>
        ///     세금계산서 현황 조회
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetIssueInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" SELECT A.TAXNO ");
            //strSql.AppendLine("      , A.TDATE ");
            //strSql.AppendLine("      , A.TAXGU ");
            //strSql.AppendLine("      , C.COM_NM AS TAXGU2 ");
            //strSql.AppendLine("      , A.AUTGB ");
            //strSql.AppendLine("      , H.COM_NM AS AUTGB2 ");
            //strSql.AppendLine("      , A.PDAT1 ");
            //strSql.AppendLine("      , A.PDAT2 ");
            //strSql.AppendLine("      , A.CVCOD ");
            //strSql.AppendLine("      , B.DEALER_NM ");
            //strSql.AppendLine("      , B.IDT_NO ");
            //strSql.AppendLine("      , B.REP_NM ");
            //strSql.AppendLine("      , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE ");
            //strSql.AppendLine("      , A.TGUBN ");
            //strSql.AppendLine("      , E.COM_NM AS TGUBN2 ");
            //strSql.AppendLine("      , A.PAYGB ");
            //strSql.AppendLine("      , F.COM_NM AS PAYGB2 ");
            //strSql.AppendLine("      , A.RK ");
            //strSql.AppendLine("      , A.MMDD1 ");
            //strSql.AppendLine("      , A.ITNM1 ");
            //strSql.AppendLine("      , A.SPEC1 ");
            //strSql.AppendLine("      , A.TWGT1 ");
            //strSql.AppendLine("      , A.COST1 ");
            //strSql.AppendLine("      , A.TAMT1 ");
            //strSql.AppendLine("      , A.TVAT1 ");
            //strSql.AppendLine("      , IFNULL(A.TAMT1, 0) + IFNULL(A.TVAT1, 0) AS TOT_AMT1 ");
            //strSql.AppendLine("      , A.TRMK1 ");
            //strSql.AppendLine("      , G1.USRNM AS CUSER ");
            //strSql.AppendLine("      , A.CDATE ");
            //strSql.AppendLine("      , G2.USRNM AS MUSER ");
            //strSql.AppendLine("      , A.MDATE ");
            //strSql.AppendLine("   FROM TAXF A ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B  ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD C #매입출구분 ");
            //strSql.AppendLine("     ON A.TAXGU = C.COM_CD ");
            //strSql.AppendLine("    AND C.CD_GB = 'AC18001_03' ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD E #청구/영수구분 ");
            //strSql.AppendLine("     ON A.TGUBN = E.COM_CD ");
            //strSql.AppendLine("    AND E.CD_GB = 'AC18001_01' ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD F #지급구분 ");
            //strSql.AppendLine("     ON A.PAYGB = F.COM_CD ");
            //strSql.AppendLine("    AND F.CD_GB = 'AC18001_02' ");
            //strSql.AppendLine("   LEFT JOIN ZUSRLST G1 ");
            //strSql.AppendLine("     ON A.CUSER = G1.USRCD ");
            //strSql.AppendLine("   LEFT JOIN ZUSRLST G2 ");
            //strSql.AppendLine("     ON A.MUSER = G1.MUSER ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD H ");
            //strSql.AppendLine("     ON A.AUTGB = H.COM_CD ");
            //strSql.AppendLine("    AND H.CD_GB = 'AC18001_04' ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T ");
            //strSql.AppendLine("    AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '2' AND A.TAXNO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '3' AND G1.USRNM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '4' AND G2.USRNM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            #endregion

            strSql.AppendLine("SELECT* ");
            strSql.AppendLine("  FROM( ");
            strSql.AppendLine("           SELECT A.TAXNO ");
            strSql.AppendLine("              , A.TDATE ");
            strSql.AppendLine("              , A.TAXGU ");
            strSql.AppendLine("              , C.COM_NM AS TAXGU2 ");
            strSql.AppendLine("              , A.AUTGB ");
            strSql.AppendLine("              , H.COM_NM AS AUTGB2 ");
            strSql.AppendLine("              , A.PDAT1 ");
            strSql.AppendLine("              , A.PDAT2 ");
            strSql.AppendLine("              , A.CVCOD ");
            strSql.AppendLine("              , B.DEALER_NM ");
            strSql.AppendLine("              , B.INITIAL_NM");
            strSql.AppendLine("              , B.IDT_NO ");
            strSql.AppendLine("              , B.REP_NM ");
            strSql.AppendLine("              , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE ");
            strSql.AppendLine("              , A.TGUBN ");
            strSql.AppendLine("              , E.COM_NM AS TGUBN2 ");
            strSql.AppendLine("              , A.PAYGB ");
            strSql.AppendLine("              , F.COM_NM AS PAYGB2 ");
            strSql.AppendLine("              , A.RK ");
            strSql.AppendLine("              , A.MMDD1 ");
            strSql.AppendLine("              , A.ITNM1 ");
            strSql.AppendLine("              , A.SPEC1 ");
            strSql.AppendLine("              , A.TWGT1 ");
            strSql.AppendLine("              , A.COST1 ");
            strSql.AppendLine("              , A.TAMT1 ");
            strSql.AppendLine("              , A.TVAT1 ");
            strSql.AppendLine("              , ISNULL(A.TAMT1, 0) + ISNULL(A.TVAT1, 0) AS TOT_AMT1 ");
            strSql.AppendLine("              , A.TRMK1 ");
            strSql.AppendLine("              , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER");
            strSql.AppendLine("              , A.CDATE ");
            strSql.AppendLine("              , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER");
            strSql.AppendLine("              , A.MDATE ");
            strSql.AppendLine("           FROM TAXF A ");
            strSql.AppendLine("           LEFT JOIN ACC_DEALER_CD B  ");
            strSql.AppendLine("             ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendLine("           LEFT JOIN COM_BASE_CD C--매입출구분 ");
            strSql.AppendLine("             ON A.TAXGU = C.COM_CD ");
            strSql.AppendLine("            AND C.CD_GB = 'AC18001_03' ");
            strSql.AppendLine("           LEFT JOIN COM_BASE_CD E--청구 / 영수구분 ");
            strSql.AppendLine("             ON A.TGUBN = E.COM_CD ");
            strSql.AppendLine("            AND E.CD_GB = 'AC18001_01' ");
            strSql.AppendLine("           LEFT JOIN COM_BASE_CD F--지급구분 ");
            strSql.AppendLine("             ON A.PAYGB = F.COM_CD ");
            strSql.AppendLine("            AND F.CD_GB = 'AC18001_02' ");
            strSql.AppendLine("           LEFT JOIN COM_BASE_CD H ");
            strSql.AppendLine("             ON A.AUTGB = H.COM_CD ");
            strSql.AppendLine("            AND H.CD_GB = 'AC18001_04' ");
            strSql.AppendLine("          WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T) A");
            strSql.AppendLine("   WHERE 1 = 1");
            strSql.AppendLine("     AND(('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("            OR ");
            strSql.AppendLine("            ('" + dicParams["FIND_IDX"] + "' = '0' AND(A.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR A.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("            OR ");
            strSql.AppendLine("            ('" + dicParams["FIND_IDX"] + "' = '1' AND A.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("            OR ");
            strSql.AppendLine("            ('" + dicParams["FIND_IDX"] + "' = '2' AND A.TAXNO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("            OR ");
            strSql.AppendLine("            ('" + dicParams["FIND_IDX"] + "' = '3' AND A.CUSER LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("            OR ");
            strSql.AppendLine("            ('" + dicParams["FIND_IDX"] + "' = '4' AND A.MUSER LIKE '%" + dicParams["FIND_WORD"] + "%')) ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        /// <summary>
        ///     세금계산서에 대한 마감자료 조회
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetCloseInfo(Dictionary<string, string> dicParams, string sKeraType)
        {
            StringBuilder strSql = new StringBuilder();

            if (sKeraType.Equals("P"))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT X1.* ");
                strSql.AppendLine("      , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("          SELECT A.JUNPYOID  ");
                strSql.AppendLine("               , A.TAXNO  ");
                strSql.AppendLine("               , A.J_DATE  ");
                strSql.AppendLine("               , A.KERATYPE  ");
                strSql.AppendLine("               , A.J_SERIAL  ");
                strSql.AppendLine("               , B.GUBUN1  ");
                strSql.AppendLine("               , A.DANJUNG  ");
                strSql.AppendLine("               , A.DANGA  ");
                strSql.AppendLine("               , CASE WHEN A.KERATYPE = '매출' THEN A.KONGKEP ELSE A.IKONGKEP END AS AMT  ");
                strSql.AppendLine("               , CASE WHEN A.KERATYPE = '매출' THEN A.KONGKEP * 0.1 ELSE A.IKONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("               , C.J_BNUM ");
                strSql.AppendLine("            FROM INLIST A   ");
                strSql.AppendLine("            LEFT JOIN JAJAE B   ");
                strSql.AppendLine("              ON A.J_SERIAL = B.J_SERIAL  ");
                strSql.AppendLine("            LEFT JOIN MESURING C ");
                strSql.AppendLine("              ON A.J_ID = C.IPCHULGO_MAIPID  ");
                strSql.AppendLine("           WHERE A.TAXNO = @TAXNO ");
                strSql.AppendLine("        ) X1 ");
                strSql.AppendLine("ORDER BY X1.J_DATE, X1.JUNPYOID");
            }
            else if (sKeraType.Equals("S"))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT X1.* ");
                strSql.AppendLine("      , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("          SELECT A.JUNPYOID  ");
                strSql.AppendLine("               , A.TAXNO  ");
                strSql.AppendLine("               , A.J_DATE  ");
                strSql.AppendLine("               , A.KERATYPE  ");
                strSql.AppendLine("               , A.J_SERIAL  ");
                strSql.AppendLine("               , B.GUBUN1  ");
                strSql.AppendLine("               , A.DANJUNG  ");
                strSql.AppendLine("               , A.DANGA  ");
                strSql.AppendLine("               , CASE WHEN A.KERATYPE = '매출' THEN A.KONGKEP ELSE A.IKONGKEP END AS AMT  ");
                strSql.AppendLine("               , CASE WHEN A.KERATYPE = '매출' THEN A.KONGKEP * 0.1 ELSE A.IKONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("               , C.J_BNUM ");
                strSql.AppendLine("            FROM INLIST A   ");
                strSql.AppendLine("            LEFT JOIN JAJAE B   ");
                strSql.AppendLine("              ON A.J_SERIAL = B.J_SERIAL  ");
                strSql.AppendLine("            LEFT JOIN MESURING C ");
                strSql.AppendLine("              ON A.J_ID = C.IPCHULGO_MACHULID  ");
                strSql.AppendLine("           WHERE A.TAXNO = @TAXNO ");
                strSql.AppendLine("           UNION ALL                                                       ");
                strSql.AppendLine("          SELECT NULL AS JUNPYOID                                            ");
                strSql.AppendLine("               , NULL AS TAXNO                                               ");
                strSql.AppendLine("               , CONVERT(VARCHAR(10), CONVERT(DATE, Z1.TDATE), 23) AS J_DATE ");
                strSql.AppendLine("               , NULL AS KERATYPE                                            ");
                strSql.AppendLine("               , NULL AS J_SERIAL                                            ");
                strSql.AppendLine("               , '인센티브' AS GUBUN1                                        ");
                strSql.AppendLine("               , 0 AS DANJUNG                                                ");
                strSql.AppendLine("               , 0 AS DANGA                                                  ");
                strSql.AppendLine("               , SUM(Z1.AMT) AS AMT                                          ");
                strSql.AppendLine("               , SUM(Z1.VAT) AS VAT                                          ");
                strSql.AppendLine("               , NULL AS J_BNUM                                              ");
                strSql.AppendLine("            FROM(SELECT A.TDATE                                              ");
                strSql.AppendLine("                        , A.ATGUB                                            ");
                strSql.AppendLine("                        , A.SEQNO                                            ");
                strSql.AppendLine("                        , CASE WHEN A.ACCOD = '0404' THEN A.ADAMT ELSE 0 END AS AMT");
                strSql.AppendLine("                        , CASE WHEN A.ACCOD = '0255' THEN A.ADAMT ELSE 0 END AS VAT");
                strSql.AppendLine("                     FROM ACTRAN A                                                 ");
                strSql.AppendLine("                     LEFT OUTER JOIN ACMSTF B                                      ");
                strSql.AppendLine("                       ON A.ACCOD = B.ACCOD                                        ");
                strSql.AppendLine("                     LEFT OUTER JOIN ACC_DEALER_CD C                               ");
                strSql.AppendLine("                       ON A.CVCOD = C.DEALER_CD                                    ");
                strSql.AppendLine("                    WHERE A.TAXNO = @TAXNO                                         ");
                strSql.AppendLine("                      AND A.ACCOD IN('0404', '0255'))Z1                            ");
                strSql.AppendLine("           GROUP BY Z1.TDATE, Z1.ATGUB, Z1.SEQNO                                   ");
                strSql.AppendLine("        ) X1 ");
                strSql.AppendLine("ORDER BY X1.J_DATE, X1.JUNPYOID");
            }
            
            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        /// <summary>
        ///     매출 -> 계산서발행 중 거래처 정보 (상단그리드)
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetIssueDealerInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");

            #region mariaDB
            //strSql.AppendLine(" SELECT X1.* ");
            //strSql.AppendLine("      , X1.AMT + X1.VAT AS TOT_AMT #합계금액 ");
            //strSql.AppendLine("      , 'N' AS CHK #전체선택 ");
            //strSql.AppendLine("      , @DATE_F AS DATE_F ");
            //strSql.AppendLine("      , @DATE_T AS DATE_T ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_ID1 #거래처코드 ");
            //strSql.AppendLine("               , A.KeraType #거래구분 ");
            //strSql.AppendLine("               , A.J_Date #작성일자 ");
            //strSql.AppendLine("               , A.JUNPYOID #전표ID ");
            //strSql.AppendLine("               , B.DEALER_NM #거래처명 ");
            //strSql.AppendLine("               , B.IDT_NO #사업자번호 ");
            //strSql.AppendLine("               , B.REP_NM #대표자명 ");
            //strSql.AppendLine("               , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE #연락처 ");
            //strSql.AppendLine("               , CAST(0 AS DOUBLE) AS TOT_WGT #총중량 ");
            //strSql.AppendLine("               , CAST(0 AS DOUBLE) AS AMT #공급가액 ");
            //strSql.AppendLine("               , CAST(0 AS DOUBLE) AS VAT #SEAK_POHAM : Y일 경우 부가세 미포함, N은 포함 ");
            ////strSql.AppendLine("               , SUM(IFNULL(A.DANJUNG, 0)) AS TOT_WGT #총중량 ");
            ////strSql.AppendLine("               , SUM(IFNULL(A.KONGKEP, 0)) AS AMT #공급가액 ");
            ////strSql.AppendLine("               , SUM(CASE C.SEAK_POHAM WHEN 'Y' THEN 0 ELSE IFNULL(A.KONGKEP, 0) END) * 0.1 AS VAT #SEAK_POHAM : Y일 경우 부가세 미포함, N은 포함 ");
            //strSql.AppendLine("               , '' AS TEMP ");
            //strSql.AppendLine("               , 0 AS CNT ");
            ////strSql.AppendLine("               , GROUP_CONCAT(A.JUNPYOID) AS TEMP ");
            ////strSql.AppendLine("               , COUNT(*) AS CNT ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B ");
            //strSql.AppendLine("              ON A.J_ID1 = B.DEALER_CD ");
            //strSql.AppendLine("            LEFT JOIN MESURING C ");
            //strSql.AppendLine("              ON A.J_ID = C.IPCHULGO_MACHULID ");
            //strSql.AppendLine("            LEFT JOIN TAXF D ");
            //strSql.AppendLine("              ON A.TAXNO = D.TAXNO ");
            //strSql.AppendLine("           WHERE A.J_DATE BETWEEN  @DATE_F AND @DATE_T   ");
            //strSql.AppendLine("             AND D.TAXNO IS NULL ");
            //strSql.AppendLine("             AND C.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("             AND A.KERATYPE = '매출' ");
            //strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND A.TAXNO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '3' AND 1 = 1 ) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '4' AND 1 = 1 ) ) ");
            //strSql.AppendLine("           GROUP BY A.J_ID1 ");
            //strSql.AppendLine("        ) X1        ");
            #endregion

            strSql.AppendLine("SELECT X1.J_ID1                                                                                                                 ");
            strSql.AppendLine("     , X1.KeraType--거래구분                                                                                                    ");
            strSql.AppendLine("     , X1.DEALER_NM--거래처명                                                                                                   ");
            strSql.AppendLine("     , X1.IDT_NO--사업자번호                                                                                                    ");
            strSql.AppendLine("     , X1.REP_NM--대표자명                                                                                                      ");
            strSql.AppendLine("     , X1.PHONE--연락처                                                                                                         ");
            strSql.AppendLine("     , 0 AS TOT_WGT--총중량                                                                                                     ");
            strSql.AppendLine("     , '0' AS AMT--공급가액                                                                                                       ");
            strSql.AppendLine("     , '0' AS VAT--SEAK_POHAM: Y일 경우 부가세 미포함, N은 포함                                                                   ");
            strSql.AppendLine("     , 0 AS CNT                                                                                                                 ");
            strSql.AppendLine("     , '0' AS TOT_AMT --합계금액                                                                                                  ");
            strSql.AppendLine("     , 'N' AS CHK --전체선택                                                                                                    ");
            strSql.AppendLine("     , '" + dicParams["DATE_F"] + "' AS DATE_F ");
            strSql.AppendLine("     , '" + dicParams["DATE_T"] + "' AS DATE_T ");
            //strSql.AppendLine("     , STRING_AGG(X1.JUNPYOID, ',') AS TEMP                                                                                     ");
            strSql.AppendLine("     , '' AS TEMP                                                                                     ");
            strSql.AppendLine("     , '' AS TEMP2                                                                                     ");
            strSql.AppendLine("  FROM(                                                                                                                         ");
            strSql.AppendLine("         SELECT A.J_ID1--거래처코드                                                                                             ");
            strSql.AppendLine("              , A.KeraType--거래구분                                                                                            ");
            strSql.AppendLine("              , A.J_Date--작성일자                                                                                              ");
            strSql.AppendLine("              , A.JUNPYOID--전표ID                                                                                              ");
            strSql.AppendLine("              , B.DEALER_NM--거래처명                                                                                           ");
            strSql.AppendLine("              , B.IDT_NO--사업자번호                                                                                            ");
            strSql.AppendLine("              , B.REP_NM--대표자명                                                                                              ");
            strSql.AppendLine("              , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE--연락처 ");
            strSql.AppendLine("              , CAST(0 AS NUMERIC) AS TOT_WGT--총중량                                                                           ");
            strSql.AppendLine("              , CAST(0 AS NUMERIC) AS AMT--공급가액                                                                             ");
            strSql.AppendLine("              , CAST(0 AS NUMERIC) AS VAT--SEAK_POHAM: Y일 경우 부가세 미포함, N은 포함                                         ");
            strSql.AppendLine("              , 0 AS CNT                                                                                                        ");
            strSql.AppendLine("           FROM INLIST A                                                                                                        ");
            strSql.AppendLine("           LEFT JOIN ACC_DEALER_CD B                                                                                            ");
            strSql.AppendLine("             ON A.J_ID1 = B.DEALER_CD                                                                                           ");
            strSql.AppendLine("           LEFT JOIN MESURING C                                                                                                 ");
            strSql.AppendLine("             ON A.J_ID = C.IPCHULGO_MACHULID                                                                                    ");
            strSql.AppendLine("           LEFT JOIN TAXF D                                                                                                     ");
            strSql.AppendLine("             ON A.TAXNO = D.TAXNO                                                                                               ");
            strSql.AppendLine("          WHERE A.J_DATE BETWEEN  '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "'   ");
            strSql.AppendLine("            AND ISNULL(D.TAXNO,'') = ''");
            strSql.AppendLine("            AND C.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("            AND A.KERATYPE = '매출' ");
            strSql.AppendLine("            AND(('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '0' AND(B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '2' AND A.TAXNO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '3' AND 1 = 1) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '4' AND 1 = 1)) ");
            strSql.AppendLine("       ) X1                                                                                                                     ");
            strSql.AppendLine("   GROUP BY X1.J_ID1,X1.KeraType , X1.DEALER_NM, X1.IDT_NO, X1.REP_NM, X1.PHONE                                                 ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        /// <summary>
        ///     매입 -> 계산서발행 중 거래처 정보 (상단그리드)
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetIssueDealerInfo2(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");

            #region mariaDB
            //strSql.AppendLine(" SELECT X1.* ");
            //strSql.AppendLine("      , X1.AMT + X1.VAT AS TOT_AMT #합계금액 ");
            //strSql.AppendLine("      , 'N' AS CHK #전체선택 ");
            //strSql.AppendLine("      , @DATE_F AS DATE_F ");
            //strSql.AppendLine("      , @DATE_T AS DATE_T ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("          SELECT A.J_ID1 #거래처코드 ");
            //strSql.AppendLine("               , A.KeraType #거래구분 ");
            //strSql.AppendLine("               , A.J_Date #작성일자 ");
            //strSql.AppendLine("               , A.JUNPYOID #전표ID ");
            //strSql.AppendLine("               , B.DEALER_NM #거래처명 ");
            //strSql.AppendLine("               , B.IDT_NO #사업자번호 ");
            //strSql.AppendLine("               , B.REP_NM #대표자명 ");
            //strSql.AppendLine("               , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE #연락처 ");
            //strSql.AppendLine("               , 0 AS TOT_WGT #총중량 ");
            //strSql.AppendLine("               , 0 AS AMT #공급가액 ");
            //strSql.AppendLine("               , 0 AS VAT #SEAK_POHAM : Y일 경우 부가세 미포함, N은 포함 ");
            ////strSql.AppendLine("               , SUM(IFNULL(A.DANJUNG, 0)) AS TOT_WGT #총중량 ");
            ////strSql.AppendLine("               , SUM(IFNULL(A.IKONGKEP, 0)) AS AMT #공급가액 ");
            ////strSql.AppendLine("               , SUM(CASE C.SEAK_POHAM WHEN 'Y' THEN 0 ELSE IFNULL(A.IKONGKEP, 0) END) * 0.1 AS VAT #SEAK_POHAM : Y일 경우 부가세 미포함, N은 포함 ");
            //strSql.AppendLine("               , '' AS TEMP ");
            //strSql.AppendLine("               , 0 AS CNT ");
            ////strSql.AppendLine("               , GROUP_CONCAT(A.JUNPYOID) AS TEMP ");
            ////strSql.AppendLine("               , COUNT(*) AS CNT ");
            //strSql.AppendLine("            FROM INLIST A ");
            //strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B ");
            //strSql.AppendLine("              ON A.J_ID1 = B.DEALER_CD ");
            //strSql.AppendLine("            LEFT JOIN MESURING C ");
            //strSql.AppendLine("              ON A.J_ID = C.IPCHULGO_MAIPID ");
            //strSql.AppendLine("            LEFT JOIN TAXF D ");
            //strSql.AppendLine("              ON A.TAXNO = D.TAXNO ");
            //strSql.AppendLine("           WHERE A.J_DATE BETWEEN  @DATE_F AND @DATE_T   ");
            //strSql.AppendLine("             AND D.TAXNO IS NULL ");
            //strSql.AppendLine("             AND A.KERATYPE = '매입' ");
            //strSql.AppendLine("             AND C.JUNPYOID IS NOT NULL ");
            //strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND A.TAXNO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '3' AND 1 = 1 ) ");
            //strSql.AppendLine("                  OR ");
            //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '4' AND 1 = 1 ) ) ");
            //strSql.AppendLine("           GROUP BY A.J_ID1 ");
            //strSql.AppendLine("        ) X1        ");
            #endregion

            strSql.AppendLine("SELECT X1.J_ID1--거래처코드                                                                                                     ");
            strSql.AppendLine("     , X1.KeraType--거래구분                                                                                                    ");
            strSql.AppendLine("     , X1.DEALER_NM--거래처명                                                                                                   ");
            strSql.AppendLine("     , X1.IDT_NO--사업자번호                                                                                                    ");
            strSql.AppendLine("     , X1.REP_NM--대표자명                                                                                                      ");
            strSql.AppendLine("     , X1.PHONE--연락처                                                                                                         ");
            strSql.AppendLine("     , 'N' AS CHK --전체선택                                                                                                    ");
            strSql.AppendLine("     , '" + dicParams["DATE_F"] + "' AS DATE_F ");
            strSql.AppendLine("     , '" + dicParams["DATE_T"] + "' AS DATE_T ");
            //strSql.AppendLine("     , STRING_AGG(X1.JUNPYOID, ',') AS TEMP                                                                                     ");
            strSql.AppendLine("     , '' AS TEMP                                                                                     ");
            strSql.AppendLine("     , 0 AS TOT_WGT--총중량                                                                                                     ");
            strSql.AppendLine("     , '0' AS AMT--공급가액                                                                                                       ");
            strSql.AppendLine("     , '0' AS VAT--SEAK_POHAM: Y일 경우 부가세 미포함, N은 포함                                                                   ");
            strSql.AppendLine("     , '0' AS TOT_AMT --합계금액                                                                                                  ");
            strSql.AppendLine("     , 0 AS CNT                                                                                                                 ");
            strSql.AppendLine("  FROM(                                                                                                                         ");
            strSql.AppendLine("         SELECT A.J_ID1--거래처코드                                                                                             ");
            strSql.AppendLine("              , A.KeraType--거래구분                                                                                            ");
            strSql.AppendLine("              --, A.J_Date--작성일자                                                                                            ");
            strSql.AppendLine("              , A.JUNPYOID--전표ID                                                                                              ");
            strSql.AppendLine("              , B.DEALER_NM--거래처명                                                                                           ");
            strSql.AppendLine("              , B.IDT_NO--사업자번호                                                                                            ");
            strSql.AppendLine("              , B.REP_NM--대표자명                                                                                              ");
            strSql.AppendLine("              , CASE WHEN B.CHRG_TEL_NO IS NULL OR B.CHRG_TEL_NO = '' THEN B.CHRG_HP_NO ELSE B.CHRG_TEL_NO END AS PHONE--연락처 ");
            strSql.AppendLine("           FROM INLIST A                                                                                                        ");
            strSql.AppendLine("           LEFT JOIN ACC_DEALER_CD B                                                                                            ");
            strSql.AppendLine("             ON A.J_ID1 = B.DEALER_CD                                                                                           ");
            strSql.AppendLine("           LEFT JOIN MESURING C                                                                                                 ");
            strSql.AppendLine("             ON A.J_ID = C.IPCHULGO_MAIPID                                                                                      ");
            strSql.AppendLine("           LEFT JOIN TAXF D                                                                                                     ");
            strSql.AppendLine("             ON A.TAXNO = D.TAXNO                                                                                               ");
            strSql.AppendLine("          WHERE A.J_DATE BETWEEN  '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "'   ");
            strSql.AppendLine("            AND ISNULL(D.TAXNO,'') = ''");
            strSql.AppendLine("            AND A.KERATYPE = '매입' ");
            strSql.AppendLine("            AND C.JUNPYOID IS NOT NULL ");
            strSql.AppendLine("            AND(('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '0' AND(B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '2' AND A.TAXNO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '3' AND 1 = 1) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '4' AND 1 = 1)) ");
            strSql.AppendLine("       ) X1                                                                                                                     ");
            strSql.AppendLine("   GROUP BY X1.J_ID1, X1.KeraType, X1.DEALER_NM, X1.IDT_NO, X1.REP_NM, X1.PHONE                                                 ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetCloseData(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            string sJunpyoCnt = dicParams["GROUP_JUNPYOID"];
            if (string.IsNullOrEmpty(sJunpyoCnt.Replace(",", "").Replace(" ", "")))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT *  ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("          SELECT X1.* ");
                strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("               , 'N' AS CHK ");
                strSql.AppendLine("            FROM ( ");
                strSql.AppendLine("                   SELECT A.JUNPYOID ");
                strSql.AppendLine("                        , A.J_DATE ");
                strSql.AppendLine("                        , B.GUBUN1 ");
                strSql.AppendLine("                        , A.DANJUNG ");
                strSql.AppendLine("                        , A.DANGA ");
                strSql.AppendLine("                        , A.KONGKEP AS AMT  ");
                strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("                        , C.J_BNUM ");
                //#0003
                strSql.AppendLine("                        , NULL AS TDATE");
                strSql.AppendLine("                        , NULL AS ATGUB");
                strSql.AppendLine("                        , NULL AS SEQNO");
                strSql.AppendLine("                     FROM INLIST A ");
                strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MACHULID  ");
                strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                strSql.AppendLine("                      AND A.KERATYPE = '매출' ");
                //#0002, #0003
                strSql.AppendLine("                    UNION ALL");
                strSql.AppendLine("                   SELECT NULL AS JUNPYOID                                                     ");
                strSql.AppendLine("                           , CONVERT(VARCHAR(10), CONVERT(DATE, Z1.TDATE), 23) AS J_DATE       ");
                strSql.AppendLine("                           , '인센티브' AS GUBUN1                                              ");
                strSql.AppendLine("                           , 0 AS DANJUNG                                                      ");
                strSql.AppendLine("                           , 0 AS DANGA                                                        ");
                strSql.AppendLine("                           , SUM(Z1.AMT) AS AMT                                                ");
                strSql.AppendLine("                           , SUM(Z1.VAT) AS VAT                                                ");
                strSql.AppendLine("                           , NULL AS J_BNUM                                                    ");
                strSql.AppendLine("                           , Z1.TDATE                                                          ");
			    strSql.AppendLine("                           , Z1.ATGUB                                                          ");
			    strSql.AppendLine("                           , Z1.SEQNO                                                          ");
                strSql.AppendLine("                        FROM(SELECT CASE WHEN A.ACCOD = '0404' THEN A.ADAMT ELSE 0 END AS AMT  ");
                strSql.AppendLine("                                   , CASE WHEN A.ACCOD = '0255' THEN A.ADAMT ELSE 0 END AS VAT ");
                strSql.AppendLine("                                   , A.TDATE                                                   ");
                strSql.AppendLine("                                   , A.ATGUB                                                   ");
                strSql.AppendLine("                                   , A.SEQNO                                                   ");
                strSql.AppendLine("                                   , A.LINNO                                                   ");
                strSql.AppendLine("                                FROM ACTRAN A                                                  ");
                strSql.AppendLine("                                LEFT OUTER JOIN ACMSTF B                                       ");
                strSql.AppendLine("                                  ON A.ACCOD = B.ACCOD                                         ");
                strSql.AppendLine("                                LEFT OUTER JOIN ACC_DEALER_CD C                                ");
                strSql.AppendLine("                                  ON A.CVCOD = C.DEALER_CD                                     ");
                strSql.AppendLine("                               WHERE A.TDATE BETWEEN REPLACE(@DATE_F, '-', '') AND REPLACE(@DATE_T, '-', '')");
                strSql.AppendLine("                                 AND A.ACCOD IN('0404', '0255')--제품매출금, 부가세예수금");
                strSql.AppendLine("                                 AND A.RK LIKE '%인센티브%'                              ");
                strSql.AppendLine("                                 AND (C.DEALER_CD = @DEALER_CD                            ");
                strSql.AppendLine("                                     OR C.DEALER_CD = '6323068002')                            ");
                strSql.AppendLine("                                 AND ISNULL(A.TAXNO,'') = ''");
                strSql.AppendLine("                                 AND A.APVYN = 'Y') Z1                                   ");
                strSql.AppendLine("                       GROUP BY Z1.TDATE, Z1.ATGUB, Z1.SEQNO                             ");
                strSql.AppendLine("                       HAVING SUM(Z1.AMT) > 0                                        ");
                strSql.AppendLine("                 ) X1 ");
                strSql.AppendLine("        ) Y1 ");
                strSql.AppendLine("    ORDER BY Y1.J_DATE, Y1.JUNPYOID  ");
            }
            else
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                #region  2023-01-05 이전코드
                //strSql.AppendLine(" SELECT *  ");
                //strSql.AppendLine("   FROM ( ");
                //strSql.AppendLine("          SELECT X1.* ");
                //strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                //strSql.AppendLine("               , 'N' AS CHK "); //디폴트 미선택으로 변경
                //strSql.AppendLine("            FROM ( ");
                //strSql.AppendLine("                   SELECT A.JUNPYOID ");
                //strSql.AppendLine("                        , A.J_DATE ");
                //strSql.AppendLine("                        , B.GUBUN1 ");
                //strSql.AppendLine("                        , A.DANJUNG ");
                //strSql.AppendLine("                        , A.DANGA ");
                //strSql.AppendLine("                        , A.KONGKEP AS AMT  ");
                //strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END AS VAT  ");
                //strSql.AppendLine("                        , C.J_BNUM ");
                //strSql.AppendLine("                     FROM INLIST A ");
                //strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                //strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                //strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                //strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MACHULID  ");
                //strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                //strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                //strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                //strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                //strSql.AppendLine("                      AND A.JUNPYOID IN (" + dicParams["GROUP_JUNPYOID"] + ") ");
                //strSql.AppendLine("                 ) X1 ");
                //strSql.AppendLine("           UNION ALL ");
                //strSql.AppendLine("          SELECT X1.* ");
                //strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                //strSql.AppendLine("               , 'N' AS CHK ");
                //strSql.AppendLine("            FROM ( ");
                //strSql.AppendLine("                   SELECT A.JUNPYOID ");
                //strSql.AppendLine("                        , A.J_DATE ");
                //strSql.AppendLine("                        , B.GUBUN1 ");
                //strSql.AppendLine("                        , A.DANJUNG ");
                //strSql.AppendLine("                        , A.DANGA ");
                //strSql.AppendLine("                        , A.KONGKEP AS AMT  ");
                //strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END AS VAT  ");
                //strSql.AppendLine("                        , C.J_BNUM ");
                //strSql.AppendLine("                     FROM INLIST A ");
                //strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                //strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                //strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                //strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MACHULID  ");
                //strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                //strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                //strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                //strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                //strSql.AppendLine("                      AND A.JUNPYOID NOT IN (" + dicParams["GROUP_JUNPYOID"] + ") ");
                //strSql.AppendLine("                 ) X1 ");
                //strSql.AppendLine("        ) Y1 ");
                //strSql.AppendLine("    ORDER BY Y1.J_DATE, Y1.JUNPYOID  ");
                #endregion

                strSql.AppendLine(" SELECT *  ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("          SELECT X1.* ");
                strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("               , 'N' AS CHK "); //디폴트 미선택으로 변경
                strSql.AppendLine("            FROM ( ");
                strSql.AppendLine("                   SELECT A.JUNPYOID ");
                strSql.AppendLine("                        , A.J_DATE ");
                strSql.AppendLine("                        , B.GUBUN1 ");
                strSql.AppendLine("                        , A.DANJUNG ");
                strSql.AppendLine("                        , A.DANGA ");
                strSql.AppendLine("                        , A.KONGKEP AS AMT  ");
                strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("                        , C.J_BNUM ");
                //#0003
                strSql.AppendLine("                        , NULL AS TDATE");
                strSql.AppendLine("                        , NULL AS ATGUB");
                strSql.AppendLine("                        , NULL AS SEQNO");
                strSql.AppendLine("                     FROM INLIST A ");
                strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MACHULID  ");
                strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                strSql.AppendLine("                      AND A.JUNPYOID IN (" + dicParams["GROUP_JUNPYOID"] + ") ");
                //#0002,#0003
                strSql.AppendLine("                    UNION ALL");
                strSql.AppendLine("                   SELECT NULL AS JUNPYOID                                                     ");
                strSql.AppendLine("                           , CONVERT(VARCHAR(10), CONVERT(DATE, Z1.TDATE), 23) AS J_DATE       ");
                strSql.AppendLine("                           , '인센티브' AS GUBUN1                                              ");
                strSql.AppendLine("                           , 0 AS DANJUNG                                                      ");
                strSql.AppendLine("                           , 0 AS DANGA                                                        ");
                strSql.AppendLine("                           , SUM(Z1.AMT) AS AMT                                                ");
                strSql.AppendLine("                           , SUM(Z1.VAT) AS VAT                                                ");
                strSql.AppendLine("                           , NULL AS J_BNUM                                                    ");
                strSql.AppendLine("                           , Z1.TDATE                                                          ");
                strSql.AppendLine("                           , Z1.ATGUB                                                          ");
                strSql.AppendLine("                           , Z1.SEQNO                                                          ");
                strSql.AppendLine("                        FROM(SELECT CASE WHEN A.ACCOD = '0404' THEN A.ADAMT ELSE 0 END AS AMT  ");
                strSql.AppendLine("                                   , CASE WHEN A.ACCOD = '0255' THEN A.ADAMT ELSE 0 END AS VAT ");
                strSql.AppendLine("                                   , A.TDATE                                                   ");
                strSql.AppendLine("                                   , A.ATGUB                                                   ");
                strSql.AppendLine("                                   , A.SEQNO                                                   ");
                strSql.AppendLine("                                   , A.LINNO                                                   ");
                strSql.AppendLine("                                FROM ACTRAN A                                                  ");
                strSql.AppendLine("                                LEFT OUTER JOIN ACMSTF B                                       ");
                strSql.AppendLine("                                  ON A.ACCOD = B.ACCOD                                         ");
                strSql.AppendLine("                                LEFT OUTER JOIN ACC_DEALER_CD C                                ");
                strSql.AppendLine("                                  ON A.CVCOD = C.DEALER_CD                                     ");
                strSql.AppendLine("                               WHERE A.TDATE BETWEEN REPLACE(@DATE_F, '-', '') AND REPLACE(@DATE_T, '-', '')");
                strSql.AppendLine("                                 AND A.ACCOD IN('0404', '0255')--제품매출금, 부가세예수금");
                strSql.AppendLine("                                 AND A.RK LIKE '%인센티브%'                              ");
                strSql.AppendLine("                                 AND (C.DEALER_CD = @DEALER_CD                            ");
                strSql.AppendLine("                                     OR C.DEALER_CD = '6323068002')                            ");
                strSql.AppendLine("                                 AND ISNULL(A.TAXNO,'') = ''");
                strSql.AppendLine("                                 AND A.APVYN = 'Y') Z1                                   ");
                strSql.AppendLine("                       GROUP BY Z1.TDATE, Z1.ATGUB, Z1.SEQNO                             ");
                strSql.AppendLine("                       HAVING SUM(Z1.AMT) > 0                                        ");
                strSql.AppendLine("                 ) X1 ");
                strSql.AppendLine("        ) Y1 ");
                strSql.AppendLine("    ORDER BY Y1.J_DATE, Y1.JUNPYOID  ");
            }
            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetCloseData2(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            string sJunpyoCnt = dicParams["GROUP_JUNPYOID"];
            if (string.IsNullOrEmpty(sJunpyoCnt.Replace(",", "").Replace(" ", "")))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT *  ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("          SELECT X1.* ");
                strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("               , 'N' AS CHK ");
                strSql.AppendLine("            FROM ( ");
                strSql.AppendLine("                   SELECT A.JUNPYOID ");
                strSql.AppendLine("                        , A.J_DATE ");
                strSql.AppendLine("                        , B.GUBUN1 ");
                strSql.AppendLine("                        , A.DANJUNG ");
                strSql.AppendLine("                        , A.DANGA ");
                strSql.AppendLine("                        , A.IKONGKEP AS AMT  ");
                strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("                        , C.J_BNUM ");
                strSql.AppendLine("                     FROM INLIST A ");
                strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MAIPID  ");
                strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                strSql.AppendLine("                      AND A.KERATYPE = '매입' ");
                strSql.AppendLine("                 ) X1 ");
                strSql.AppendLine("        ) Y1 ");
                strSql.AppendLine("    ORDER BY Y1.J_DATE, Y1.JUNPYOID  ");
            }
            else
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT *  ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("          SELECT X1.* ");
                strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("               , 'N' AS CHK "); //디폴트 미선택으로 변경
                strSql.AppendLine("            FROM ( ");
                strSql.AppendLine("                   SELECT A.JUNPYOID ");
                strSql.AppendLine("                        , A.J_DATE ");
                strSql.AppendLine("                        , B.GUBUN1 ");
                strSql.AppendLine("                        , A.DANJUNG ");
                strSql.AppendLine("                        , A.DANGA ");
                strSql.AppendLine("                        , A.IKONGKEP AS AMT  ");
                strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("                        , C.J_BNUM ");
                strSql.AppendLine("                     FROM INLIST A ");
                strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MAIPID  ");
                strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                strSql.AppendLine("                      AND A.JUNPYOID IN (" + dicParams["GROUP_JUNPYOID"] + ") ");
                strSql.AppendLine("                 ) X1 ");
                strSql.AppendLine("           UNION ALL ");
                strSql.AppendLine("          SELECT X1.* ");
                strSql.AppendLine("               , X1.AMT + X1.VAT AS TOT_AMT ");
                strSql.AppendLine("               , 'N' AS CHK ");
                strSql.AppendLine("            FROM ( ");
                strSql.AppendLine("                   SELECT A.JUNPYOID ");
                strSql.AppendLine("                        , A.J_DATE ");
                strSql.AppendLine("                        , B.GUBUN1 ");
                strSql.AppendLine("                        , A.DANJUNG ");
                strSql.AppendLine("                        , A.DANGA ");
                strSql.AppendLine("                        , A.IKONGKEP AS AMT  ");
                strSql.AppendLine("                        , CASE WHEN C.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END AS VAT  ");
                strSql.AppendLine("                        , C.J_BNUM ");
                strSql.AppendLine("                     FROM INLIST A ");
                strSql.AppendLine("                     LEFT JOIN JAJAE B  ");
                strSql.AppendLine("                       ON A.J_SERIAL = B.J_SERIAL ");
                strSql.AppendLine("                     LEFT JOIN MESURING C  ");
                strSql.AppendLine("                       ON A.J_ID = C.IPCHULGO_MAIPID  ");
                strSql.AppendLine("                    WHERE A.J_DATE BETWEEN @DATE_F AND @DATE_T    ");
                strSql.AppendLine("                      AND A.J_ID1 = @DEALER_CD ");
                strSql.AppendLine("                      AND ISNULL(A.TAXNO,'') = '' ");
                strSql.AppendLine("                      AND C.JUNPYOID IS NOT NULL ");
                strSql.AppendLine("                      AND A.JUNPYOID NOT IN (" + dicParams["GROUP_JUNPYOID"] + ") ");
                strSql.AppendLine("                 ) X1 ");
                strSql.AppendLine("        ) Y1 ");
                strSql.AppendLine("    ORDER BY Y1.J_DATE, Y1.JUNPYOID  ");
            }
            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable SetGridData(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();
            
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("  SELECT MMDD1 AS MMDD, ITNM1 AS ITNM, SPEC1 AS SPEC, TWGT1 AS TWGT ");
            strSql.AppendLine("       , COST1 AS COST, TAMT1 AS TAMT, TVAT1 AS TVAT, TRMK1 AS TRMK ");
            strSql.AppendLine("       , TAMT1 + TVAT1 AS TOT_AMT ");
            strSql.AppendLine("    FROM TAXF X1 ");
            strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            strSql.AppendLine("   UNION ALL ");
            strSql.AppendLine("  SELECT MMDD2, ITNM2, SPEC2, TWGT2 ");
            strSql.AppendLine("       , COST2, TAMT2, TVAT2, TRMK2 ");
            strSql.AppendLine("       , TAMT2 + TVAT2 AS TOT_AMT ");
            strSql.AppendLine("    FROM TAXF X1 ");
            strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            strSql.AppendLine("   UNION ALL ");
            strSql.AppendLine("  SELECT MMDD3, ITNM3, SPEC3, TWGT3 ");
            strSql.AppendLine("       , COST3, TAMT3, TVAT3, TRMK3 ");
            strSql.AppendLine("       , TAMT3 + TVAT3 AS TOT_AMT ");
            strSql.AppendLine("    FROM TAXF X1 ");
            strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            strSql.AppendLine("   UNION ALL ");
            strSql.AppendLine("  SELECT MMDD4, ITNM4, SPEC4, TWGT4 ");
            strSql.AppendLine("       , COST4, TAMT4, TVAT4, TRMK4 ");
            strSql.AppendLine("       , TAMT4 + TVAT2 AS TOT_AMT ");
            strSql.AppendLine("    FROM TAXF X1 ");
            strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);            
        }

        #region[GridViewL1 관련 이벤트]

        private void GridViewL1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewL1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewL1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
            {
                GridRetrL2.DataSource = null;
                return;
            }

            string sTaxNo = GridViewL1.GetRowCellValue(e.FocusedRowHandle, GridColL1TaxNo)?.ToString();
            if (string.IsNullOrEmpty(sTaxNo))
            {
                return;
            }
            string sKeraType = GridViewL1.GetRowCellValue(e.FocusedRowHandle, GridColL1TaxGu1)?.ToString();


            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("TAXNO", sTaxNo);

            Cursor = Cursors.WaitCursor;
            if (TabControlSub.SelectedTabPage == TabPageSubClose)
            {
                GridRetrL2.DataSource = null;
                GridRetrL2.DataSource = GetCloseInfo(dicParams, sKeraType);
            }
            else if(TabControlSub.SelectedTabPage == TabPageSubTax)
            {
                GridRetrL3.DataSource = null;
                GridRetrL3.DataSource = SetGridData(dicParams);
            }

            Cursor = Cursors.Default;
        }

        #endregion[GridViewL1 관련 이벤트]

        private void TabControlSub_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            string sTaxNo = GridViewL1.GetFocusedRowCellValue(GridColL1TaxNo)?.ToString();
            if (string.IsNullOrEmpty(sTaxNo))
            {
                return;
            }

            string sKeraType = GridViewL1.GetFocusedRowCellValue(GridColL1TaxGu1)?.ToString();
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("TAXNO", sTaxNo);

            if (TabControlSub.SelectedTabPage == TabPageSubTax)
            {
                GridRetrL3.DataSource = null;
                GridRetrL3.DataSource = SetGridData(dicParams);
            }
            else if (TabControlSub.SelectedTabPage == TabPageSubClose)
            {
                GridRetrL2.DataSource = null;
                GridRetrL2.DataSource = GetCloseInfo(dicParams, sKeraType);
            }
        }

        #region[GridViewL2 관련 이벤트]

        private void GridViewL2_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewL2_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridViewL2 관련 이벤트]

        #region[GridViewR1 관련 이벤트]

        private void GridViewR1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewR1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewR1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
            {
                GridRetrR2.DataSource = null;
                return;
            }

            Cursor = Cursors.WaitCursor;
            DataRow drFocusedRow = GridViewR1.GetFocusedDataRow();
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            string[] sTemps = RemoveAt(drFocusedRow["TEMP"].ToString().Replace(" ", "").Split(','), "인센티브");

            dicParams.Add("DATE_F", drFocusedRow["DATE_F"].ToString());
            dicParams.Add("DATE_T", drFocusedRow["DATE_T"].ToString());
            dicParams.Add("DEALER_CD", drFocusedRow["J_ID1"].ToString());
            dicParams.Add("GROUP_JUNPYOID", sTemps.Aggregate((text, next) => text + "," + next));

            GridRetrR2.DataSource = GetCloseData(dicParams);
            Cursor = Cursors.Default;
        }

        private void RepoChkR1Yn_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            if (chk.Checked)
            {
                GridViewR1.SetFocusedRowCellValue(GridColR1Chk, "Y");
                SetCheckDetail(GridViewR2, GridColR2Chk, "Y");
            }
            else
            {
                GridViewR1.SetFocusedRowCellValue(GridColR1Chk, "N");
                GridViewR1.SetFocusedRowCellValue(GridColR1Temp, "");
                GridViewR1.SetFocusedRowCellValue(GridColR1Temp2, "");
                SetCheckDetail(GridViewR2, GridColR2Chk, "N");
            }

            ApplyTempString(GridViewR2);
        }
        
        private void SetCheckDetail(GridView view, GridColumn col, string sYn)
        {
            for(int i = 0; i < view.RowCount; i++)
            {
                view.SetRowCellValue(i, col, sYn);
            }
        }

        private void GridViewR1_MouseDown(object sender, MouseEventArgs e)
        {
            GridView view = sender as GridView;
            GridHitInfo hitiInfo = view.CalcHitInfo(e.Location);
            if (hitiInfo.HitTest == GridHitTest.Column)
            {
                if (hitiInfo.Column == GridColR1Chk)
                {
                    if (GridViewR1.RowCount == 0)
                        return;

                    string sYn = GridViewR1.GetRowCellValue(0, GridColR1Chk)?.ToString();
                    if (sYn.Equals("Y"))
                    {
                        SetValueZero();
                        SetCheckDetail(view, GridColR1Chk, "N");
                        SetCheckDetail(view, GridColR1Temp, string.Empty);
                        SetCheckDetail(view, GridColR1Temp2, string.Empty);
                        SetCheckDetail(GridViewR2, GridColR2Chk, "N");
                    }
                    else if (sYn.Equals("N"))
                    {
                        SetCheckDetail(view, GridColR1Chk, "Y");
                        SetCheckDetail(GridViewR2, GridColR2Chk, "Y");
                        //GridViewR1.SetRowCellValue(0, GridColR1Chk, "Y");
                        //SetCheckDetail(view, GridColR1Chk, "Y");
                        //TEMP 값 DB 태우기
                        GetGroupJunpyoIDs(view);
                    }
                    
                    //ApplyTempString(GridViewR2);
                }
            }
        }

        private void GridViewP1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
            {
                GridRetrP2.DataSource = null;
                return;
            }

            Cursor = Cursors.WaitCursor;
            DataRow drFocusedRow = GridViewP1.GetFocusedDataRow();
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", drFocusedRow["DATE_F"].ToString());
            dicParams.Add("DATE_T", drFocusedRow["DATE_T"].ToString());
            dicParams.Add("DEALER_CD", drFocusedRow["J_ID1"].ToString());
            dicParams.Add("GROUP_JUNPYOID", drFocusedRow["TEMP"].ToString());

            GridRetrP2.DataSource = GetCloseData2(dicParams);
            Cursor = Cursors.Default;
        }

        private void GridViewP1_MouseDown(object sender, MouseEventArgs e)
        {
            GridView view = sender as GridView;
            GridHitInfo hitiInfo = view.CalcHitInfo(e.Location);
            if (hitiInfo.HitTest == GridHitTest.Column)
            {
                if (hitiInfo.Column == GridColP1Chk)
                {
                    if (GridViewP1.RowCount == 0)
                        return;

                    string sYn = GridViewP1.GetRowCellValue(0, GridColP1Chk)?.ToString();
                    if (sYn.Equals("Y"))
                    {
                        SetValueZero2();
                        SetCheckDetail(view, GridColP1Chk, "N");
                        SetCheckDetail(view, GridColP1Temp, string.Empty);
                        SetCheckDetail(GridViewP2, GridColP2Chk, "N");
                    }
                    else if (sYn.Equals("N"))
                    {
                        SetCheckDetail(view, GridColP1Chk, "Y");
                        SetCheckDetail(GridViewP2, GridColP2Chk, "Y");
                        //GridViewR1.SetRowCellValue(0, GridColR1Chk, "Y");
                        //SetCheckDetail(view, GridColR1Chk, "Y");
                        //TEMP 값 DB 태우기
                        GetGroupJunpyoIDs2(view);
                    }

                    //ApplyTempString(GridViewR2);
                }
            }
        }

        private void RepoChkP2Yn_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            if (chk.Checked)
            {
                GridViewP2.SetFocusedRowCellValue(GridColP2Chk, "Y");
            }
            else
            {
                GridViewP2.SetFocusedRowCellValue(GridColP2Chk, "N");
            }
            
            ApplyTempString2(GridViewP2);
        }

        private void SetValueZero()
        {
            for(int i = 0; i < GridViewR1.RowCount; i++)
            {
                GridViewR1.SetRowCellValue(i, GridColR1TotWeight, 0);
                GridViewR1.SetRowCellValue(i, GridColR1TotAmt, 0);
                GridViewR1.SetRowCellValue(i, GridColR1TotVat, 0);
                GridViewR1.SetRowCellValue(i, GridColR1TotSumAmt, 0);
                GridViewR1.SetRowCellValue(i, GridColR1Cnt, 0);
            }
        }

        private void GridViewP2_MouseDown(object sender, MouseEventArgs e)
        {
            GridView view = sender as GridView;
            GridHitInfo hitiInfo = view.CalcHitInfo(e.Location);
            if (hitiInfo.HitTest == GridHitTest.Column)
            {
                if (hitiInfo.Column == GridColP2Chk)
                {
                    if (view.RowCount == 0)
                        return;

                    string sYn = view.GetRowCellValue(0, GridColP2Chk)?.ToString();
                    if (sYn.Equals("Y"))
                    {
                        SetCheckDetail(view, GridColP2Chk, "N");
                    }
                    else if (sYn.Equals("N"))
                    {
                        SetCheckDetail(view, GridColP2Chk, "Y");
                    }

                    ApplyTempString2(view);
                }
            }
        }

        private void SetValueZero2()
        {
            for (int i = 0; i < GridViewP1.RowCount; i++)
            {
                GridViewP1.SetRowCellValue(i, GridColP1TotWeight, 0);
                GridViewP1.SetRowCellValue(i, GridColP1TotAmt, 0);
                GridViewP1.SetRowCellValue(i, GridColP1TotVat, 0);
                GridViewP1.SetRowCellValue(i, GridColP1TotSumAmt, 0);
                GridViewP1.SetRowCellValue(i, GridColP1Cnt, 0);
            }
        }

        private void GetGroupJunpyoIDs(GridView view)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                StringBuilder strSql = new StringBuilder();
                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                for (int i = 0; i < view.RowCount; i++)
                {
                    dicParams.Clear();
                    dicParams.Add("CVCOD", view.GetRowCellValue(i, GridColR1CvCod)?.ToString());
                    dicParams.Add("DATE_F", view.GetRowCellValue(i, GridColR1DateF)?.ToString());
                    dicParams.Add("DATE_T", view.GetRowCellValue(i, GridColR1DateT)?.ToString());

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" SELECT SUM(A.DANJUNG) AS DANJUNG ");
                    //strSql.AppendLine("      , SUM(A.KONGKEP) AS AMT  ");
                    //strSql.AppendLine("      , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END ) AS VAT ");
                    //strSql.AppendLine("      , GROUP_CONCAT(A.JUNPYOID) AS TEMP ");
                    //strSql.AppendLine("      , COUNT(*) AS CNT ");
                    //strSql.AppendLine("   FROM INLIST A ");
                    //strSql.AppendLine("   LEFT JOIN MESURING B  ");
                    //strSql.AppendLine("     ON A.J_ID = B.IPCHULGO_MACHULID ");
                    //strSql.AppendLine("  WHERE A.J_ID1 = @CVCOD ");
                    //strSql.AppendLine("    AND A.KERATYPE = '매출' ");
                    //strSql.AppendLine("    AND A.J_DATE BETWEEN @DATE_F AND @DATE_T ");
                    #endregion

                    strSql.AppendLine("SELECT SUM(A.DANJUNG) AS DANJUNG                                                ");
                    strSql.AppendLine("     , SUM(A.KONGKEP) AS AMT                                                    ");
                    strSql.AppendLine("     , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.KONGKEP * 0.1 END ) AS VAT");
                    strSql.AppendLine("     , STRING_AGG(A.JUNPYOID, ',') AS TEMP                                      ");
                    strSql.AppendLine("     , COUNT(*) AS CNT                                                          ");
                    strSql.AppendLine("  FROM INLIST A                                                                 ");
                    strSql.AppendLine("  LEFT JOIN MESURING B                                                          ");
                    strSql.AppendLine("    ON A.J_ID = B.IPCHULGO_MACHULID                                             ");
                    strSql.AppendLine(" WHERE A.J_ID1 = @CVCOD                                                   ");
                    strSql.AppendLine("       A.KERATYPE = '매출'                                                      ");
                    strSql.AppendLine("   AND A.J_DATE BETWEEN @DATE_F AND @DATE_T                           ");

                    DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                    double dDanjung = Convert.ToDouble(dt.Rows[0]["DANJUNG"]);
                    double dAmt = Convert.ToDouble(dt.Rows[0]["AMT"]);
                    double dVat = Convert.ToDouble(dt.Rows[0]["VAT"]);
                    double dTotAmt = dAmt + dVat;
                    string sTemp = dt.Rows[0]["TEMP"]?.ToString();
                    string sTemp2 = dt.Rows[0]["TEMP2"]?.ToString();
                    double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);

                    view.SetRowCellValue(i, GridColR1TotWeight, dt.Rows[0]["DANJUNG"]);
                    view.SetRowCellValue(i, GridColR1TotAmt, dt.Rows[0]["AMT"]);
                    view.SetRowCellValue(i, GridColR1TotVat, dt.Rows[0]["VAT"]);
                    view.SetRowCellValue(i, GridColR1TotSumAmt, dTotAmt);
                    view.SetRowCellValue(i, GridColR1Temp, sTemp);
                    view.SetRowCellValue(i, GridColR1Temp2, sTemp2);
                    view.SetRowCellValue(i, GridColR1Cnt, dCnt);

                }

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
            }
        }

        private void GetGroupJunpyoIDs2(GridView view)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                StringBuilder strSql = new StringBuilder();
                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                for (int i = 0; i < view.RowCount; i++)
                {
                    dicParams.Clear();
                    dicParams.Add("CVCOD", view.GetRowCellValue(i, GridColP1CvCod)?.ToString());
                    dicParams.Add("DATE_F", view.GetRowCellValue(i, GridColP1DateF)?.ToString());
                    dicParams.Add("DATE_T", view.GetRowCellValue(i, GridColP1DateT)?.ToString());

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" SELECT SUM(A.DANJUNG) AS DANJUNG ");
                    //strSql.AppendLine("      , SUM(A.IKONGKEP) AS AMT  ");
                    //strSql.AppendLine("      , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END ) AS VAT ");
                    //strSql.AppendLine("      , GROUP_CONCAT(A.JUNPYOID) AS TEMP ");
                    //strSql.AppendLine("      , COUNT(*) AS CNT ");
                    //strSql.AppendLine("   FROM INLIST A ");
                    //strSql.AppendLine("   LEFT JOIN MESURING B  ");
                    //strSql.AppendLine("     ON A.J_ID = B.IPCHULGO_MAIPID ");
                    //strSql.AppendLine("  WHERE A.J_ID1 = @CVCOD ");
                    //strSql.AppendLine("    AND A.KERATYPE = '매입' ");
                    //strSql.AppendLine("    AND A.J_DATE BETWEEN @DATE_F AND @DATE_T ");
                    #endregion

                    strSql.AppendLine(" SELECT SUM(A.DANJUNG) AS DANJUNG ");
                    strSql.AppendLine("      , SUM(A.IKONGKEP) AS AMT  ");
                    strSql.AppendLine("      , SUM(CASE WHEN B.SEAK_POHAM = 'Y' THEN 0 ELSE A.IKONGKEP * 0.1 END ) AS VAT ");
                    strSql.AppendLine("      , STRING_AGG(A.JUNPYOID,',') AS TEMP ");
                    strSql.AppendLine("      , COUNT(*) AS CNT ");
                    strSql.AppendLine("   FROM INLIST A ");
                    strSql.AppendLine("   LEFT JOIN MESURING B  ");
                    strSql.AppendLine("     ON A.J_ID = B.IPCHULGO_MAIPID ");
                    strSql.AppendLine("  WHERE A.J_ID1 = @CVCOD ");
                    strSql.AppendLine("    AND A.KERATYPE = '매입' ");
                    strSql.AppendLine("    AND A.J_DATE BETWEEN @DATE_F AND @DATE_T ");

                    DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                    double dDanjung = Convert.ToDouble(dt.Rows[0]["DANJUNG"]);
                    double dAmt = Convert.ToDouble(dt.Rows[0]["AMT"]);
                    double dVat = Convert.ToDouble(dt.Rows[0]["VAT"]);
                    double dTotAmt = dAmt + dVat;
                    string sTemp = dt.Rows[0]["TEMP"]?.ToString();
                    double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);

                    view.SetRowCellValue(i, GridColP1TotWeight, dt.Rows[0]["DANJUNG"]);
                    view.SetRowCellValue(i, GridColP1TotAmt, dt.Rows[0]["AMT"]);
                    view.SetRowCellValue(i, GridColP1TotVat, dt.Rows[0]["VAT"]);
                    view.SetRowCellValue(i, GridColP1TotSumAmt, dTotAmt);
                    view.SetRowCellValue(i, GridColP1Temp, sTemp);
                    view.SetRowCellValue(i, GridColP1Cnt, dCnt);
                }

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
            }
        }

        #endregion[GridViewR1 관련 이벤트]

        #region[GridViewR2 관련 이벤트]

        private void GridViewR2_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewR2_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RepoChkR2Yn_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            if (chk.Checked)
            {
                GridViewR2.SetFocusedRowCellValue(GridColR2Chk, "Y");
            }
            else
            {
                GridViewR2.SetFocusedRowCellValue(GridColR2Chk, "N");
            }

            ApplyTempString(GridViewR2);
        }

        private void GridViewR2_MouseDown(object sender, MouseEventArgs e)
        {
            GridView view = sender as GridView;
            GridHitInfo hitiInfo = view.CalcHitInfo(e.Location);
            if (hitiInfo.HitTest == GridHitTest.Column)
            {
                if (hitiInfo.Column == GridColR2Chk)
                { 
                    if (view.RowCount == 0)
                        return;

                    string sYn = view.GetRowCellValue(0, GridColR2Chk)?.ToString();
                    if (sYn.Equals("Y"))
                    {
                        SetCheckDetail(view, GridColR2Chk, "N");
                    }
                    else if (sYn.Equals("N"))
                    {
                        SetCheckDetail(view, GridColR2Chk, "Y");
                    }

                    ApplyTempString(view);
                }
            }
        }

        #endregion[GridViewR2 관련 이벤트]

        private void ApplyTempString(GridView view)
        {
            double dTotAmt = 0;
            double dVat = 0;
            double dWeight = 0;
            string sTemp = string.Empty;

            Dictionary<string, double[]> dicParams = new Dictionary<string, double[]>();
            Dictionary<string, string> dicInsen = new Dictionary<string, string>();

            int iInsen = 1;
            //#0003
            string strJson = GridViewR1.GetFocusedRowCellValue("TEMP2")?.ToString();

            DataTable dtInsen = new DataTable();

            dtInsen.Columns.Add("TDATE");
            dtInsen.Columns.Add("ATGUB");
            dtInsen.Columns.Add("SEQNO");

            for (int i = 0; i < view.RowCount; i++)
            {
                string sChk = view.GetRowCellValue(i, GridColR2Chk)?.ToString();
                
                if (sChk.Equals("Y"))
                {
                    string sJunpyoId = view.GetRowCellValue(i, GridColR2JunpyoId)?.ToString();
                    string sGubun1 = view.GetRowCellValue(i, GridColR2Gubun1)?.ToString();

                    if (sJunpyoId.Equals("") && sGubun1.Equals("인센티브"))
                    {
                        sJunpyoId = "인센티브" + iInsen++;

                        string sTDATE = view.GetRowCellValue(i,"TDATE")?.ToString();
                        string sATGUB = view.GetRowCellValue(i,"ATGUB")?.ToString();
                        string sSEQNO = view.GetRowCellValue(i,"SEQNO")?.ToString();

                        DataRow row = dtInsen.NewRow();

                        row["TDATE"] = sTDATE;
                        row["ATGUB"] = sATGUB;
                        row["SEQNO"] = sSEQNO;

                        dtInsen.Rows.Add(row);
                    }

                    string sAmt = view.GetRowCellValue(i, GridColR2Amt)?.ToString();
                    string sVat = view.GetRowCellValue(i, GridColR2Vat)?.ToString();
                    string sWeight = view.GetRowCellValue(i, GridColR2Danjung)?.ToString();

                    double[] dArr = new double[3];
                    dArr[0] = string.IsNullOrEmpty(sAmt) ? 0 : Convert.ToDouble(sAmt);
                    dArr[1] = string.IsNullOrEmpty(sVat) ? 0 : Convert.ToDouble(sVat);
                    dArr[2] = string.IsNullOrEmpty(sWeight) ? 0 : Convert.ToDouble(sWeight);

                    dicParams.Add(sJunpyoId, dArr);
                }
                //else
                //{
                //    if (dtInsen != null && dtInsen.Rows.Count > 0)
                //    {
                //        string sTDATE = view.GetRowCellValue(i, "TDATE")?.ToString();
                //        string sATGUB = view.GetRowCellValue(i, "ATGUB")?.ToString();
                //        string sSEQNO = view.GetRowCellValue(i, "SEQNO")?.ToString();

                //        dtInsen.AcceptChanges();
                //        foreach (DataRow row in dtInsen.Rows)
                //        {
                //            string sTDATE2 = row["TDATE"]?.ToString();
                //            string sATGUB2 = row["ATGUB"]?.ToString();
                //            string sSEQNO2 = row["SEQNO"]?.ToString();

                //            if (sTDATE.Equals(sTDATE2) && sATGUB.Equals(sATGUB2) && sSEQNO.Equals(sSEQNO2))
                //            {
                //                row.Delete();
                //            }
                //        }
                //        dtInsen.AcceptChanges();
                //    }
                //}
            }

            string sResult = string.Empty;
            int iCnt = 1;
            foreach(KeyValuePair<string, double[]> param in dicParams)
            {
                if (iCnt == dicParams.Count)
                {
                    sResult += param.Key;
                    //if (dicParams.Count == 1)
                    //{
                    //    sResult += param.Key;
                    //}
                    //else
                    //{
                    //    sResult += param.Key;
                    //}
                }
                else
                {
                    sResult += string.Format("{0},", param.Key);
                }

                dTotAmt += param.Value[0];
                dVat += param.Value[1];
                dWeight += param.Value[2];

                iCnt++;
            }

            GridViewR1.SetFocusedRowCellValue(GridColR1TotWeight, dWeight);
            GridViewR1.SetFocusedRowCellValue(GridColR1TotAmt, doubleComma(dTotAmt));
            GridViewR1.SetFocusedRowCellValue(GridColR1TotVat, doubleComma(dVat));
            GridViewR1.SetFocusedRowCellValue(GridColR1TotSumAmt, doubleComma(dTotAmt + dVat));
            GridViewR1.SetFocusedRowCellValue(GridColR1Temp, sResult);
            if(dtInsen != null && dtInsen.Rows.Count > 0)
            {
                GridViewR1.SetFocusedRowCellValue(GridColR1Temp2, ComnEtcFunc.DataTableToJsonWithNewtonSoft(dtInsen));
            }
            else
            {
                GridViewR1.SetFocusedRowCellValue(GridColR1Temp2, "");
            }
            GridViewR1.SetFocusedRowCellValue(GridColR1Cnt, iCnt - 1);

            if (view.RowCount == dicParams.Count)
                GridViewR1.SetFocusedRowCellValue(GridColR1Chk, "Y");
            else
                GridViewR1.SetFocusedRowCellValue(GridColR1Chk, "N");
        }
        
        /// <summary>
        /// 숫자세자리 콤마해서 문자열 반환
        /// </summary>
        /// <param name="dVal"></param>
        /// <returns></returns>
        private string doubleComma(double dVal)
        {
            string sTemp = "";
            string sVal = dVal.ToString();
            char[] cVal = sVal.ToCharArray().Reverse().ToArray();

            int iCnt = 1;
            foreach(char str in cVal)
            {
                if(iCnt % 3 == 0 && iCnt < cVal.Length)
                {
                    sTemp += str + ",";
                }
                else
                {
                    sTemp += str;
                }

                iCnt++;
            }

            char[] cResult = sTemp.ToCharArray().Reverse().ToArray();

            string sResult = new string(cResult);

            return sResult;
        }

        private void ApplyTempString2(GridView view)
        {
            double dTotAmt = 0;
            double dVat = 0;
            double dWeight = 0;
            string sTemp = string.Empty;
            Dictionary<string, double[]> dicParams = new Dictionary<string, double[]>();

            for (int i = 0; i < view.RowCount; i++)
            {
                string sChk = view.GetRowCellValue(i, GridColP2Chk)?.ToString();

                if (sChk.Equals("Y"))
                {
                    string sJunpyoId = view.GetRowCellValue(i, GridColP2JunpyoId)?.ToString();
                    string sAmt = view.GetRowCellValue(i, GridColP2Amt)?.ToString();
                    string sVat = view.GetRowCellValue(i, GridColP2Vat)?.ToString();
                    string sWeight = view.GetRowCellValue(i, GridColP2Danjung)?.ToString();

                    double[] dArr = new double[3];
                    dArr[0] = string.IsNullOrEmpty(sAmt) ? 0 : Convert.ToDouble(sAmt);
                    dArr[1] = string.IsNullOrEmpty(sVat) ? 0 : Convert.ToDouble(sVat);
                    dArr[2] = string.IsNullOrEmpty(sWeight) ? 0 : Convert.ToDouble(sWeight);

                    dicParams.Add(sJunpyoId, dArr);
                }
            }

            string sResult = string.Empty;
            int iCnt = 1;
            foreach (KeyValuePair<string, double[]> param in dicParams)
            {
                if (iCnt == dicParams.Count)
                {
                    sResult += param.Key;
                    //if (dicParams.Count == 1)
                    //{
                    //    sResult += param.Key;
                    //}
                    //else
                    //{
                    //    sResult += param.Key;
                    //}
                }
                else
                {
                    sResult += string.Format("{0},", param.Key);
                }

                dTotAmt += param.Value[0];
                dVat += param.Value[1];
                dWeight += param.Value[2];

                iCnt++;
            }

            GridViewP1.SetFocusedRowCellValue(GridColP1TotWeight, dWeight);
            GridViewP1.SetFocusedRowCellValue(GridColP1TotAmt, doubleComma(dTotAmt));
            GridViewP1.SetFocusedRowCellValue(GridColP1TotVat, doubleComma(dVat));
            GridViewP1.SetFocusedRowCellValue(GridColP1TotSumAmt, doubleComma(dTotAmt + dVat));
            GridViewP1.SetFocusedRowCellValue(GridColP1Temp, sResult);
            GridViewP1.SetFocusedRowCellValue(GridColP1Cnt, iCnt - 1);

            if (view.RowCount == dicParams.Count)
                GridViewP1.SetFocusedRowCellValue(GridColP1Chk, "Y");
            else
                GridViewP1.SetFocusedRowCellValue(GridColP1Chk, "N");

        }
        private void TabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (e.Page == TabPageTax)
            {
                LayoutBtnIssue.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnExcel.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutBtnAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutBtnDel.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutPrint.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                LayoutDropBtn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else if(e.Page == TabPageIssueSale)
            {
                LayoutBtnIssue.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutBtnExcel.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnDel.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutPrint.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                LayoutDropBtn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                InitControls();
                UpdateDropDownButton(BtnOutMaster);
            }
            else if (e.Page == TabPageIssuePurc)
            {
                LayoutBtnIssue.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutBtnExcel.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnDel.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutPrint.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                LayoutDropBtn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                InitControls2();
                UpdateDropDownButton(BtnInMaster);
            }

            BtnRetr.PerformClick();
        }

        private void GridViewL3_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewL3_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewL1_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if(e.Column == GridColL1AutGb2)
            {
                string sVal = GridViewL1.GetRowCellValue(e.RowHandle, GridColL1AutGb)?.ToString();
                if (!string.IsNullOrEmpty(sVal))
                {
                    if (sVal.Equals("1")) //자동
                    {
                        e.Appearance.BackColor = Color.PaleGreen;
                    }
                    else if (sVal.Equals("2")) //수동
                    {
                        e.Appearance.BackColor = Color.LightSkyBlue;
                    }
                }
            }
        }

        private void GridViewL1_RowClick(object sender, RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                AC18001F02 frm = new AC18001F02();
                frm.DR_INFO = GridViewL1.GetFocusedDataRow();
                if(frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr.PerformClick();
                }
            }
        }

        private void RepoChkP2Yn_EditValueChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            GridViewP2.SetFocusedRowCellValue(GridColP2Chk, chk.EditValue);
        }

        private void RepoChkP1Yn_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            if (chk.Checked)
            {
                GridViewP1.SetFocusedRowCellValue(GridColP1Chk, "Y");
                SetCheckDetail(GridViewP2, GridColP2Chk, "Y");
            }
            else
            {
                GridViewR1.SetFocusedRowCellValue(GridColP1Chk, "N");
                GridViewR1.SetFocusedRowCellValue(GridColP1Temp, "");
                SetCheckDetail(GridViewP2, GridColP2Chk, "N");
            }

            ApplyTempString2(GridViewP2);
        }

        private void GridViewP1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewP1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewP2_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewP2_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AC18001F01_TextChanged(object sender, EventArgs e)
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
                layoutControl2.SaveLayoutToXml(path + @"\" + this.Name + "_Layout1.xaml");
                layoutControl3.SaveLayoutToXml(path + @"\" + this.Name + "_Layout2.xaml");
                layoutControl6.SaveLayoutToXml(path + @"\" + this.Name + "_Layout3.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnOutMaster;
        BarButtonItem BtnOutDetail;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnOutMaster = new BarButtonItem(barManager1, "매출거래처");
            BtnOutDetail = new BarButtonItem(barManager1, "매출상세내역");
            popupMenu1.AddItem(BtnOutMaster);
            popupMenu1.AddItem(BtnOutDetail);

            DropBtnExcel.DropDownControl = popupMenu1;

            BtnOutMaster.Tag = "매출거래처";
            BtnOutMaster.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnOutMaster_ItemClick);

            BtnOutDetail.Tag = "매출상세내역";
            BtnOutDetail.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnOutDetail_ItemClick);
        }

        BarButtonItem BtnInMaster;
        BarButtonItem BtnInDetail;
        private void InitControls2()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnInMaster = new BarButtonItem(barManager1, "매입거래처");
            BtnInDetail = new BarButtonItem(barManager1, "매입상세내역");
            popupMenu1.AddItem(BtnInMaster);
            popupMenu1.AddItem(BtnInDetail);

            DropBtnExcel.DropDownControl = popupMenu1;

            BtnInMaster.Tag = "매입거래처";
            BtnInMaster.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnInMaster_ItemClick);

            BtnInDetail.Tag = "매입상세내역";
            BtnInDetail.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnInDetail_ItemClick);
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnExcel.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnExcel.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnExcel.Tag = submenuItem.Tag;
            DropBtnExcel.Text = string.Format("{0}", submenuItem.Tag);
        }

        private void BtnOutMaster_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            
        }

        private void BtnOutDetail_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
        }

        private void BtnInMaster_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);

        }

        private void BtnInDetail_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
        }

        private void DropBtnExcel_Click(object sender, EventArgs e)
        {
            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "매출거래처")
            {
                ComnEtcFunc.ExportExcelFile(this.Text + "_" + tag, GridRetrR1);
            }
            else if (tag == "매출상세내역")
            {
                ComnEtcFunc.ExportExcelFile(this.Text + "_" + tag, GridRetrR2);
            }
            else if (tag == "매입거래처")
            {
                ComnEtcFunc.ExportExcelFile(this.Text + "_" + tag, GridRetrP1);
            }
            else if (tag == "매입상세내역")
            {
                ComnEtcFunc.ExportExcelFile(this.Text + "_" + tag, GridRetrP2);
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        double tot_Danga = 0;
        double tot_Danga_cnt = 0;
        private void GridViewL2_CustomSummaryCalculate(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            GridSummaryItem gridSummaryItem = e.Item as GridSummaryItem;

            GridView gridView = sender as GridView;

            switch (e.SummaryProcess)
            {
                //calculation entry point
                case CustomSummaryProcess.Start:
                    tot_Danga = 0;
                    tot_Danga_cnt = 0;
                    break;
                //consequent calculations
                case CustomSummaryProcess.Calculate:
                    string sGubun = gridView.GetRowCellValue(e.RowHandle, "GUBUN1")?.ToString();
                    string sDanga = gridView.GetRowCellValue(e.RowHandle, "DANGA")?.ToString();

                    if (!string.IsNullOrEmpty(sGubun) && !sGubun.Equals("인센티브"))
                    {
                        double dDanga = 0;
                        double.TryParse(sDanga, out dDanga);

                        tot_Danga += dDanga;
                        tot_Danga_cnt++;
                    }

                    break;
                //final summary value
                case CustomSummaryProcess.Finalize:

                    double dTot = Math.Round(tot_Danga / tot_Danga_cnt);

                    e.TotalValue = string.Format("{0}", dTot.ToString("0,#.#"));
                    break;
            }
        }

        private void GridViewL3_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e)
        {
            GridSummaryItem gridSummaryItem = e.Item as GridSummaryItem;

            GridView gridView = sender as GridView;

            switch (e.SummaryProcess)
            {
                //calculation entry point
                case CustomSummaryProcess.Start:
                    tot_Danga = 0;
                    tot_Danga_cnt = 0;
                    break;
                //consequent calculations
                case CustomSummaryProcess.Calculate:
                    string sDanga = gridView.GetRowCellValue(e.RowHandle, "COST")?.ToString();

                    if (!string.IsNullOrEmpty(sDanga))
                    {
                        double dDanga = 0;
                        double.TryParse(sDanga, out dDanga);

                        tot_Danga += dDanga;
                        tot_Danga_cnt++;
                    }

                    break;
                //final summary value
                case CustomSummaryProcess.Finalize:

                    double dTot = Math.Round(tot_Danga / tot_Danga_cnt);

                    e.TotalValue = string.Format("{0}", dTot.ToString("0,#.#"));
                    break;
            }
        }
    }
}