using ComLib;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-07
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            거래처 초성검색 기능 추가 (쿼리확인)
 * 
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 * 
 * 수정일자 : 2021-03-12
 * 수정자   : 고혜성
 * Reference Key : #00001
 * 수정내용 : (현업요청)
 *            1. 로그적용
 *            
 * 수정일자 : 2021-03-12
 * 수정자   : 고혜성 
 * Reference Key : #00002
 * 수정내용 : 1. 현업에서는 금액부분과 관련된 로직체크를 원함
 *            기초정보인 거래처관리는 로그처리 우선 주석처리 진행
 *            추후 요청 시 아래 INSERT 로그로직 주석해제하여 사용할 것 (Reference Key 참조) 
 * 
 * 수정일자 : 2021-08-06
 * 수정자   : 정은영
 * Reference Key : #00003
 * 수정내용 : (현업요청)
 *            1. 사업자 등록번호 확인 후 (폐업, 비폐업 구분X) 중복 메세지 띄운 후 계속 진행할건지 확인.
 */
namespace AccAdm
{
    public partial class AccDealerCdDev : DevExpress.XtraEditors.XtraForm
    {
        public AccDealerCdDev()
        {
            InitializeComponent();
        }
        public string _dealercd;
        private string PROCEDURE_ID = "DP_IMAGE";
        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccDealerCdDev_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            SetLookUpEdit(LkupBankCd, "1", "", "Y");
            SetLookUpEdit(LkupEditBillKind, "2", "Y", "Y");
            SetLookUpEdit(LkupBIssueGB, "3", "Y", "Y");
            SetLookUpEdit(LkupEditPayGB, "4", "Y", "Y");
            SetLookUpEdit(LkupEditChrgID, "5", "Y", "Y");

            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            arrGrdView = new GridView[] { GridViewDealerMng };
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
            
            DataTable dt = GetLookUpData("1");
            ComGrid.SetLookUpEdit(LkupDealerRetr, dt, "CD", "NM", "Y");
            //LkupDealerRetr.Properties.Columns[LkupDealerRetr.Properties.ValueMember].Visible = false; 
            //ComGrid.SetLookUpEdit(LkupDealerRetr, dt, "CD", "CD", "Y");
            //LkupDealerRetr.Properties.PopulateColumns();
            //LkupDealerRetr.Properties.Columns[1].Visible = false;

            BtnRetr_Click(null, null);

            Cursor = Cursors.Default;
        }

        private void AccDealerCdDev_Shown(object sender, EventArgs e)
        {
            //LkupDealerRetr.Properties.PopulateColumns();
            //LkupDealerRetr.Properties.Columns[LkupDealerRetr.Properties.ValueMember].Visible = false;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ClearFps();

            StringBuilder strSql = new StringBuilder();

            string sDealerCd = LkupDealerRetr.EditValue?.ToString();
            string sEobYn = RdgbMenuEobYn.EditValue.ToString();
            string sSelectedIdx = CboFindSbj.SelectedIndex.ToString();
            string sFindWord = TxtFindWord.EditValue?.ToString().Trim();

            Cursor = Cursors.WaitCursor;

            strSql.AppendLine(" SELECT A.DEALER_CD ");
            strSql.AppendLine("      , A.IDT_NO");
            strSql.AppendLine("      , A.CORP_NO ");
            #region mariaDB 
            //strSql.AppendLine("      , DATE_FORMAT(A.STRT_YMD, '%Y-%m-%d') AS STRT_YMD ");
            //strSql.AppendLine("      , DATE_FORMAT(A.END_YMD, '%Y-%m-%d') AS END_YMD ");
            #endregion
            strSql.AppendLine("      , CONVERT(VARCHAR(10), CONVERT(DATE, A.STRT_YMD), 23) AS STRT_YMD");
            strSql.AppendLine("      , CONVERT(VARCHAR(10), CONVERT(DATE, A.END_YMD), 23) AS END_YMD");
            strSql.AppendLine("      , A.DEALER_GB");
            strSql.AppendLine("      , A.DEALER_NM ");
            strSql.AppendLine("      , A.INITIAL_NM");
            strSql.AppendLine("      , A.REP_NM ");
            strSql.AppendLine("      , A.BIZ_NM ");
            strSql.AppendLine("      , A.TYPE_NM ");
            strSql.AppendLine("      , A.ZIP ");
            strSql.AppendLine("      , A.ZIP_SEQ ");
            strSql.AppendLine("      , A.ADDR ");
            strSql.AppendLine("      , A.DTL_ADDR ");
            strSql.AppendLine("      , A.J_REGION ");
            strSql.AppendLine("      , A.BANK_CD ");
            strSql.AppendLine("      , A.BANK_ACNT_NO ");
            strSql.AppendLine("      , A.ACNT_HOLDER ");
            strSql.AppendLine("      , A.EMAIL ");
            strSql.AppendLine("      , A.HOMEPG ");
            strSql.AppendLine("      , A.FAX");
            strSql.AppendLine("      , A.WEB_FAX_YN ");
            strSql.AppendLine("      , A.SCM_PSWD ");
            strSql.AppendLine("      , A.CHRG_ID");
            strSql.AppendLine("      , A.CHRG_NM ");
            strSql.AppendLine("      , A.CHRG_TEL_NO ");
            strSql.AppendLine("      , A.CHRG_RGN_NO ");
            strSql.AppendLine("      , A.CHRG_HP_NO ");
            strSql.AppendLine("      , A.CHRG_EMAIL ");
            strSql.AppendLine("      , A.BFR_DEALER_CD ");
            strSql.AppendLine("      , A.BFR_DEALER_NM ");
            strSql.AppendLine("      , A.BILL_KIND ");
            strSql.AppendLine("      , A.BILL_ISSUE_GB ");
            strSql.AppendLine("      , A.PAY_GB ");
            strSql.AppendLine("      , A.PAY_TERM ");
            strSql.AppendLine("      , A.SPPL_LIMIT_DCNT ");
            strSql.AppendLine("      , A.EOB_YN ");
            strSql.AppendLine("      , A.NOTE ");
            strSql.AppendLine("      , A2.IMTXT1 ");
            strSql.AppendLine("      , A2.IMTXT2");
            strSql.AppendLine("      , A2.IMTXT3 ");
            strSql.AppendLine("      , A.BANKYN");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID  AS NUMERIC) IS NULL THEN A.MFY_ID  ELSE DBO.FN_USRNM(A.MFY_ID ) END AS MFY_ID ");strSql.AppendLine("      , B.EMP_NM ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A ");
            strSql.AppendLine("   LEFT JOIN  HR_EMP_BASIS B ON A.CHRG_ID = B.EMP_ID");
            strSql.AppendLine("   LEFT JOIN  ACC_DEALER_IMAGE A2 ON A.DEALER_CD = A2.DEALER_CD");
            strSql.AppendLine("   WHERE 1 = 1 ");
            strSql.AppendLine("     AND (('" + sFindWord + "' = '' AND 1 = 1) ");
            strSql.AppendLine("          OR   ");
            strSql.AppendLine("          (('" + sFindWord + "' <> '' AND '" + sSelectedIdx + "' = '0') AND (A.DEALER_NM LIKE '%" + sFindWord + "%' OR A.INITIAL_NM LIKE '%" + sFindWord + "%' ))   ");
            strSql.AppendLine("          OR   ");
            strSql.AppendLine("          (('" + sFindWord + "' <> '' AND '" + sSelectedIdx + "' = '1') AND A.DEALER_CD LIKE '%" + sFindWord + "%' )   ");
            strSql.AppendLine("          OR   ");
            strSql.AppendLine("          (('" + sFindWord + "' <> '' AND '" + sSelectedIdx + "' = '2') AND A.IDT_NO LIKE '%" + sFindWord + "%' )   ");
            strSql.AppendLine("          OR   ");
            strSql.AppendLine("          (('" + sFindWord + "' <> '' AND '" + sSelectedIdx + "' = '3') AND A.REP_NM LIKE '%" + sFindWord + "%' )   ");
            strSql.AppendLine("          OR   ");
            strSql.AppendLine("          (('" + sFindWord + "' <> '' AND '" + sSelectedIdx + "' = '4') AND B.EMP_NM LIKE '%" + sFindWord + "%' ) )  ");
            if (!string.IsNullOrEmpty(sDealerCd))
            {
                strSql.AppendLine("  AND A.DEALER_CD = " + sDealerCd + " ");
            }
            if (!sEobYn.Equals("ALL"))
            {
                strSql.AppendLine("    AND A.EOB_YN = '" + sEobYn + "' ");
            }

            strSql.AppendLine("  ORDER BY A.DEALER_CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridDealerMng.DataSource = null;
            GridHistoryRetr.DataSource = null;
            GridDealerMng.DataSource = dt;
            if (dt.Rows.Count == 0)
            {
                SetClear();
                TxtFindWord.SelectAll();
                TxtFindWord.Focus();
            }
            //GridViewDealerMng.BestFitColumns();

            // GridViewDealerMng.RestoreLayoutFromXml(@"C:\Users\USER2\source\repos\SteelAccSln\AccAdm\xaml\AccDealerCdDev.xaml");
            if (!string.IsNullOrEmpty(_dealercd))
            {
                GridViewDealerMng.FocusedRowHandle = GridViewDealerMng.LocateByDisplayText(0, GridColDealerCd, _dealercd);
                _dealercd = string.Empty;
            }


            Cursor = Cursors.Default;
        }

        private void BtnCrete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            SetInit();

            //ClearAllForm(layoutControl2);

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT MAX(A.DEALER_CD) AS MAX_CD ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            DataRow[] row = dt.Select();
            double dDealerCd = Convert.ToDouble(row[0]["MAX_CD"]) + 1000;

            TxtDealerCd.EditValue = dDealerCd.ToString();

            TxtIdtNo.ReadOnly = false;
            TxtIdtNo.Focus();
            TxtIdtNo.SelectAll();
            DateEditTo.EditValue = "2999-12-31";

            TxtDealerNM.SelectAll();
            TxtDealerNM.Focus();
            ChkAutoNo.ReadOnly = false;
        }

        private void SetInit()
        {
            TxtDealerCd.EditValue = null;
            TxtIdtNo.EditValue = null;
            TxtCorpNo.EditValue = null;

            TxtDealerNM.EditValue = null;
            CbxDealerGB.SelectedIndex = 0;
            TxtInitialNm.EditValue = null;
            TxtBizNM.EditValue = null;
            TxtRepNM.EditValue = null;
            TxtTypeNM.EditValue = null;
            TxtZipNum.EditValue = null;
            TxtAddr.EditValue = null;
            TxtDtlAddr.EditValue = null;
            TxtRegion.EditValue = null;
            TxtEmail.EditValue = null;
            TxtHomePG.EditValue = null;
            TxtChrgRgnNo.EditValue = null;
            TxtFaxNum.EditValue = null;
            RdgbWebFaxGb.EditValue = "N";
            TxtScmPw.EditValue = null;
            LkupBankCd.EditValue = null;
            TxtHolder.EditValue = null;
            TxtBankNum.EditValue = null;
            LkupEditChrgID.EditValue = null;
            TxtChrgNM.EditValue = null;
            TxtChrgHP.EditValue = null;
            TxtChrgEmail.EditValue = null;
            TxtChrgTel.EditValue = null;
            LkupEditBillKind.EditValue = "****";
            LkupBIssueGB.EditValue = "****";
            LkupEditPayGB.EditValue = "****";
            TxtPayTerm.EditValue = 0;
            TxtLtDcnt.EditValue = 0;
            RdgbEobYn.EditValue = "N";
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = "2999-12-31";
            TxtBfrDlrCd.EditValue = null;
            TxtBfrDlrNM.EditValue = null;
            TxtNote.EditValue = null;

            //히스토리 삭제 추가
            GridHistoryRetr.DataSource = null;

            buttonEdit1.EditValue = null;
            buttonEdit2.EditValue = null;
            buttonEdit3.EditValue = null;

            //수정 필요
            //if (sBankCd == "****") sBankCd = "";
            //string sBankHolder = TxtHolder.Text;
            //string sBankNM = TxtBankNum.Text;
            //string sChrgID = LkupEditChrgID.EditValue?.ToString();
            //if (sChrgID == "****") sChrgID = "";
            //string sChrgNM = TxtChrgNM.Text;
            //string sChrgHP = TxtChrgHP.Text;
            //string sChrgEmail = TxtChrgEmail.Text;
            //string sChrgTel = TxtChrgTel.Text;
            //string sBillKind = LkupEditBillKind.EditValue?.ToString();
            //if (sBillKind == "****") sBillKind = "";
            //string sBlIssue = LkupBIssueGB.EditValue?.ToString();
            //if (sBlIssue == "****") sBlIssue = "";
            //string sPayGB = LkupEditPayGB.EditValue?.ToString();
            //if (sPayGB == "****") sPayGB = "";
            //string sPayTerm = TxtPayTerm.Text;
            //if (TxtPayTerm.Text == "") sPayTerm = "0";
            //string sLimitDcnt = TxtLtDcnt.Text;
            //if (TxtLtDcnt.Text == "") sLimitDcnt = "0";
            //string sEobYn = RdgbEobYn.EditValue?.ToString();
            //string sNote = TxtNote.Text;
            //string sStartDay = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            //string sEndDay = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            //string sPreDealerCd = TxtBfrDlrCd.Text;
            //string sPreDealerNM = TxtBfrDlrNM.Text;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (TxtDealerNM.Text == "")
            {
                XtraMessageBox.Show("거래처명을 쓰세요.");
                TxtDealerNM.Focus();
                return;
            }
            else if (DateEditFrom.EditValue.ToString() == "" || DateEditTo.EditValue.ToString() == "")
            {
                XtraMessageBox.Show("시작일자 및 종료일자를 선택하세요.");
                DateEditFrom.Focus();
                return;
            }
            else if (TxtPayTerm.Text.Length > 2)
            {
                XtraMessageBox.Show("결제기간 개월 수는 두 자리를 초과할 수 없습니다");
                return;
            }
            else if (TxtLtDcnt.Text.Length > 4)
            {
                XtraMessageBox.Show("납품기한일수는 네 자리를 초과할 수 없습니다. ");
                return;
            }
            else if (string.IsNullOrEmpty(TxtDealerCd.EditValue?.ToString()))
            {
                XtraMessageBox.Show("거래처코드가 존재하지 않습니다.");
                TxtDealerCd.Focus();
                return;
            }
            else
            {
                Cursor = Cursors.WaitCursor;
                if (!ChkAutoNo.Checked)
                {
                    if (FindDealerCd(TxtDealerCd.EditValue?.ToString()))
                    {
                        Cursor = Cursors.Default;
                        string sMsg = string.Format("거래처코드 : {0}\r\n해당 거래처코드는 이미 존재합니다.\r\n다른 거래처코드를 입력하세요.", TxtDealerCd.EditValue?.ToString());
                        XtraMessageBox.Show(sMsg);
                        TxtDealerCd.SelectAll();
                        TxtDealerCd.Focus();
                        return;
                    }
                }

                /*
                * #00003
                */
                StringBuilder strSql = new StringBuilder();

                string sDealerCd = TxtDealerCd.EditValue?.ToString();
                string sIdtNo = TxtIdtNo.Text.Replace("-", "");
                sIdtNo = string.IsNullOrEmpty(sIdtNo) ? string.Empty : sIdtNo;

                if (!sIdtNo.Equals(""))
                {
                    strSql.Clear();
                    strSql.AppendLine("SELECT DEALER_NM");
                    strSql.AppendLine("FROM ACC_DEALER_CD");
                    strSql.AppendLine("WHERE IDT_NO = '" + sIdtNo + "'");
                    if (!sDealerCd.Equals(""))
                        strSql.AppendLine("AND DEALER_CD !=" + sDealerCd);

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            string sDealerNm = string.Empty;

                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (i != 0)
                                {
                                    sDealerNm += ", ";
                                }

                                sDealerNm += dt.Rows[i]["DEALER_NM"]?.ToString();
                            }

                            if (DialogResult.Yes == XtraMessageBox.Show(sDealerNm + "의 거래처와 사업자번호가 중복입니다. 계속 진행하시겠습니까?",
                                                                        "사업자번호 중복", MessageBoxButtons.YesNo))
                            {

                            }
                            else
                            {
                                /*
                                 * #00004
                                 */
                                int idxFoc = GridViewDealerMng.FocusedRowHandle;

                                BtnRetr.PerformClick();
                                GridViewDealerMng.FocusedRowHandle = idxFoc;

                                Cursor = Cursors.Default;
                                return;
                            }

                        }
                    }
                }
                string sCorpNo = TxtCorpNo.EditValue?.ToString().Trim();

                string sDealerName = TxtDealerNM.Text;
                string sDealerGB = CbxDealerGB.EditValue.ToString();
                string sInitialNm = TxtInitialNm.EditValue?.ToString();
                string sBizName = TxtBizNM.Text;
                string sRepName = TxtRepNM.Text;
                string sTypeName = TxtTypeNM.Text;
                string sZipNum = TxtZipNum.Text;
                string sZipSeq = string.Empty;
                string sAddr = TxtAddr.Text;
                string sDtlAddr = TxtDtlAddr.Text;
                string sRegion = TxtRegion.Text;
                string sEMail = TxtEmail.Text;
                string sHomePG = TxtHomePG.Text;
                string sFax = TxtFaxNum.Text;
                string sWebFaxYn = RdgbWebFaxGb.EditValue?.ToString();
                string sSCMPswd = TxtScmPw.EditValue?.ToString();
                string sBankCd = LkupBankCd.EditValue?.ToString();
                string sBankHolder = TxtHolder.Text;
                string sBankNM = TxtBankNum.Text;
                string sChrgID = LkupEditChrgID.EditValue?.ToString();
                if (sChrgID == "****") sChrgID = "";
                string sChrgNM = TxtChrgNM.Text;
                string sChrgHP = TxtChrgHP.Text;
                string sChrgEmail = TxtChrgEmail.Text;
                string sChrgTel = TxtChrgTel.Text;
                string sChrgRgnTel = TxtChrgRgnNo.EditValue?.ToString().Trim();
                string sBillKind = LkupEditBillKind.EditValue?.ToString();
                if (sBillKind == "****") sBillKind = "";
                string sBlIssue = LkupBIssueGB.EditValue?.ToString();
                if (sBlIssue == "****") sBlIssue = "";
                string sPayGB = LkupEditPayGB.EditValue?.ToString();
                if (sPayGB == "****") sPayGB = "";
                string sPayTerm = TxtPayTerm.Text;
                if (TxtPayTerm.Text == "") sPayTerm = "0";
                string sLimitDcnt = TxtLtDcnt.Text;
                if (TxtLtDcnt.Text == "") sLimitDcnt = "0";
                string sEobYn = RdgbEobYn.EditValue?.ToString();
                string sNote = TxtNote.Text;
                string sStartDay = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                string sEndDay = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                string sPreDealerCd = TxtBfrDlrCd.Text;
                string sPreDealerNM = TxtBfrDlrNM.Text;
                string sBANKYN = ChkBANKYN.EditValue?.ToString();
                string sCurTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string sEntID = FmMainToolBar2.UserID;
                string sMfyId = FmMainToolBar2.UserID;

                string sLogRmk = string.Empty;
                
                //거래처구분 계좌 체크 
                //23.09.25요청으로 인한 계좌체크 해제
                //if(!string.IsNullOrEmpty(sDealerGB) && sDealerGB.Equals("계좌"))
                //{
                //    if (string.IsNullOrEmpty(sBankCd))
                //    {
                //        Cursor = Cursors.Default;
                //        XtraMessageBox.Show("은행을 선택해주세요.");
                //        LkupBankCd.Focus();
                //        return;
                //    }
                //
                //    if (string.IsNullOrEmpty(sBankNM))
                //    {
                //        Cursor = Cursors.Default;
                //        XtraMessageBox.Show("계좌번호를 입력해주세요.");
                //        TxtBankNum.Focus();
                //        return;
                //    }
                //}

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                #region 이전코드(기간모름)
                //dicParams.Add("DEALER_CD", sDealerCd);
                //dicParams.Add("IDT_NO", sIdtNo);
                //dicParams.Add("CORP_NO", sCorpNo);
                //dicParams.Add("STRT_YMD", sStartDay);
                //dicParams.Add("END_YMD", sEndDay);
                //dicParams.Add("DEALER_GB", sDealerGB);
                //dicParams.Add("DEALER_NM", sDealerName);
                //dicParams.Add("INITIAL_NM", sInitialNm);
                //dicParams.Add("REP_NM", sRepName);
                //dicParams.Add("BIZ_NM", sBizName);
                //dicParams.Add("TYPE_NM", sTypeName);
                //dicParams.Add("ZIP", sZipNum);
                //dicParams.Add("ZIP_SEQ", sZipSeq);
                //dicParams.Add("ADDR", sAddr);
                //dicParams.Add("DTL_ADDR", sDtlAddr);
                //dicParams.Add("J_REGION", sRegion);
                //dicParams.Add("BANK_CD", sBankCd);
                //dicParams.Add("BANK_ACNT_NO", sBankNM);
                //dicParams.Add("ACNT_HOLDER", sBankHolder);
                //dicParams.Add("EMAIL", sEMail);
                //dicParams.Add("HOMEPG", sHomePG);
                //dicParams.Add("FAX", sFax);
                //dicParams.Add("WEB_FAX_YN", sWebFaxYn);
                //dicParams.Add("SCM_PSWD", sSCMPswd);
                //dicParams.Add("CHRG_ID", sChrgID);
                //dicParams.Add("CHRG_NM", sChrgNM);
                //dicParams.Add("CHRG_TEL_NO", sChrgTel);
                //dicParams.Add("CHRG_RGN_NO", sChrgRgnTel);
                //dicParams.Add("CHRG_HP_NO", sChrgHP);
                //dicParams.Add("CHRG_EMAIL", sChrgEmail);
                //dicParams.Add("BFR_DEALER_CD", sPreDealerCd);
                //dicParams.Add("BFR_DEALER_NM", sPreDealerNM);
                //dicParams.Add("BILL_KIND", sBillKind);
                //dicParams.Add("BILL_ISSUE_GB", sBlIssue);
                //dicParams.Add("PAY_GB", sPayGB);
                //dicParams.Add("PAY_TERM", sPayTerm);
                //dicParams.Add("SPPL_LIMIT_DCNT", sLimitDcnt);
                //dicParams.Add("EOB_YN", sEobYn);
                //dicParams.Add("NOTE", sNote);
                //dicParams.Add("ENT_DT", sCurTime);
                //dicParams.Add("ENT_ID", sEntID);
                //dicParams.Add("MFY_DT", sCurTime);
                //dicParams.Add("MFY_ID", sMfyId);

                ////LOG 변수부분
                //dicParams.Add("CUR_TIME", sCurTime);
                //dicParams.Add("USRCD", FmMainToolBar2.UserID);
                //dicParams.Add("PGMID", this.Name);
                //dicParams.Add("IP", ComnEtcFunc.GetLocalIP());
                #endregion

                try
                {

                    #region 이전코드(기간 모름)
                    //strSql.Clear();
                    //strSql.AppendLine(" ");
                    //strSql.AppendLine(" SELECT A.* ");
                    //strSql.AppendLine("      , B.COM_NM AS BANK_NM ");
                    //strSql.AppendLine("      , C.EMP_NM ");
                    //strSql.AppendLine("   FROM ACC_DEALER_CD A ");
                    //strSql.AppendLine("   LEFT JOIN COM_BASE_CD B ");
                    //strSql.AppendLine("     ON A.BANK_CD = B.COM_CD ");
                    //strSql.AppendLine("    AND B.CD_GB = 'BANK_CD'");
                    //strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C ");
                    //strSql.AppendLine("     ON A.CHRG_ID = C.EMP_ID ");
                    //strSql.AppendLine("  WHERE A.DEALER_CD = " + sDealerCd + " ");

                    //DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    //sLogRmk = string.Format("[거래처관리]거래처코드 : {0}, 거래처명 : {1} -> "
                    //    , dtPrv.Rows[0]["DEALER_CD"]
                    //    , dtPrv.Rows[0]["DEALER_NM"]);

                    //int iLogCnt = 0;
                    //if(dtPrv.Rows.Count > 0)
                    //{
                    //    string sPRV_DEALER_NM = dtPrv.Rows[0]["DEALER_NM"]?.ToString();
                    //    string sPRV_IDT_NO = dtPrv.Rows[0]["IDT_NO"]?.ToString();
                    //    string sPRV_CORP_NO = dtPrv.Rows[0]["CORP_NO"]?.ToString();
                    //    string sPRV_STRT_YMD = dtPrv.Rows[0]["STRT_YMD"]?.ToString();
                    //    string sPRV_END_YMD = dtPrv.Rows[0]["END_YMD"]?.ToString();
                    //    string sPRV_DEALER_GB = dtPrv.Rows[0]["DEALER_GB"]?.ToString();
                    //    string sPRV_INITIAL_NM = dtPrv.Rows[0]["INITIAL_NM"]?.ToString();
                    //    string sPRV_REP_NM = dtPrv.Rows[0]["REP_NM"]?.ToString();
                    //    string sPRV_BIZ_NM = dtPrv.Rows[0]["BIZ_NM"]?.ToString();
                    //    string sPRV_TYPE_NM = dtPrv.Rows[0]["TYPE_NM"]?.ToString();
                    //    string sPRV_ADDR = dtPrv.Rows[0]["ADDR"]?.ToString();
                    //    string sPRV_DTL_ADDR = dtPrv.Rows[0]["DTL_ADDR"]?.ToString();
                    //    string sPRV_J_REGION = dtPrv.Rows[0]["J_REGION"]?.ToString();
                    //    string sPRV_BANK_CD = dtPrv.Rows[0]["BANK_CD"]?.ToString();
                    //    string sPRV_BANK_NM = dtPrv.Rows[0]["BANK_NM"]?.ToString();
                    //    string sPRV_BANK_ACNT_NO = dtPrv.Rows[0]["BANK_ACNT_NO"]?.ToString();
                    //    string sPRV_ACNT_HOLDER = dtPrv.Rows[0]["ACNT_HOLDER"]?.ToString();
                    //    string sPRV_EMAIL = dtPrv.Rows[0]["EMAIL"]?.ToString();
                    //    string sPRV_FAX = dtPrv.Rows[0]["FAX"]?.ToString();
                    //    string sPRV_WEB_FAX_YN = dtPrv.Rows[0]["WEB_FAX_YN"]?.ToString();
                    //    string sPRV_SCM_PSWD = dtPrv.Rows[0]["SCM_PSWD"]?.ToString();
                    //    string sPRV_CHRG_ID = dtPrv.Rows[0]["CHRG_ID"]?.ToString();
                    //    string sPRV_CHRG_NM = dtPrv.Rows[0]["EMP_NM"]?.ToString();
                    //    string sPRV_CHRG_TEL_NO = dtPrv.Rows[0]["CHRG_TEL_NO"]?.ToString();
                    //    string sPRV_CHRG_RGN_NO = dtPrv.Rows[0]["CHRG_RGN_NO"]?.ToString();
                    //    string sPRV_CHRG_HP_NO = dtPrv.Rows[0]["CHRG_HP_NO"]?.ToString();
                    //    string sPRV_CHRG_EMAIL = dtPrv.Rows[0]["CHRG_EMAIL"]?.ToString();

                    //    if (!sPRV_DEALER_NM.Equals(sDealerName))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 거래처명 : {0} ▶ {1} ", sPRV_DEALER_NM, sDealerName);
                    //    }

                    //    if (!sPRV_IDT_NO.Equals(sIdtNo))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 사업자번호 : {0} ▶ {1} ", sPRV_IDT_NO, sIdtNo);
                    //    }

                    //    if (!sPRV_CORP_NO.Equals(sCorpNo))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 법인번호 : {0} ▶ {1} ", sPRV_CORP_NO, sCorpNo);
                    //    }

                    //    if (!sPRV_STRT_YMD.Equals(sStartDay))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 시작일자 : {0} ▶ {1} ", sPRV_STRT_YMD, sStartDay);
                    //    }

                    //    if (!sPRV_END_YMD.Equals(sEndDay))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 종료일자 : {0} ▶ {1} ", sPRV_END_YMD, sEndDay);
                    //    }

                    //    if (!sPRV_DEALER_GB.Equals(sDealerGB))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 거래처구분 : {0} ▶ {1} ", sPRV_DEALER_GB, sDealerGB);
                    //    }

                    //    if (!sPRV_INITIAL_NM.Equals(sInitialNm))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 초성명 : {0} ▶ {1} ", sPRV_INITIAL_NM, sInitialNm);
                    //    }

                    //    if (!sPRV_REP_NM.Equals(sRepName))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 대표자 : {0} ▶ {1} ", sPRV_REP_NM, sRepName);
                    //    }

                    //    if (!sPRV_BIZ_NM.Equals(sBizName))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 업태 : {0} ▶ {1} ", sPRV_BIZ_NM, sBizName);
                    //    }

                    //    if (!sPRV_TYPE_NM.Equals(sTypeName))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 종목 : {0} ▶ {1} ", sPRV_TYPE_NM, sTypeName);
                    //    }

                    //    if (!sPRV_ADDR.Equals(sAddr))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 상위주소 : {0} ▶ {1} ", sPRV_ADDR, sAddr);
                    //    }

                    //    if (!sPRV_DTL_ADDR.Equals(sDtlAddr))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 하위주소 : {0} ▶ {1} ", sPRV_DTL_ADDR, sDtlAddr);
                    //    }

                    //    if (!sPRV_J_REGION.Equals(sRegion))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("l 지역 : {0} ▶ {1} ", sPRV_J_REGION, sRegion);
                    //    }

                    //    if (!sPRV_BANK_CD.Equals(sBankCd))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("은행 : {0} ▶ {1} ", sPRV_BANK_CD, sBankCd);
                    //    }

                    //    if (!sPRV_BANK_ACNT_NO.Equals(sBankNM))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("계좌번호 : {0} ▶ {1} ", sPRV_BANK_ACNT_NO, sBankNM);
                    //    }

                    //    if (!sPRV_ACNT_HOLDER.Equals(sBankHolder))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("예금주 : {0} ▶ {1} ", sPRV_ACNT_HOLDER, sBankHolder);
                    //    }

                    //    if (!sPRV_EMAIL.Equals(sEMail))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("예금주 : {0} ▶ {1} ", sPRV_EMAIL, sEMail);
                    //    }

                    //    if (!sPRV_FAX.Equals(sFax))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("팩스 : {0} ▶ {1} ", sPRV_FAX, sFax);
                    //    }

                    //    if (!sPRV_WEB_FAX_YN.Equals(sWebFaxYn))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("팩스사용여부 : {0} ▶ {1} ", sPRV_WEB_FAX_YN, sWebFaxYn);
                    //    }

                    //    if (!sPRV_SCM_PSWD.Equals(sSCMPswd))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("SCM패스워드 : {0} ▶ {1} ", sPRV_SCM_PSWD, sSCMPswd);
                    //    }

                    //    if (!sPRV_CHRG_ID.Equals(sChrgID))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("담당자 : {0} ▶ {1} ", sPRV_CHRG_NM, LkupEditChrgID.Text);
                    //    }

                    //    if (!sPRV_CHRG_TEL_NO.Equals(sChrgTel))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("담당자전화번호 : {0} ▶ {1} ", sPRV_CHRG_TEL_NO, sChrgTel);
                    //    }

                    //    if (!sPRV_CHRG_EMAIL.Equals(sChrgEmail))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("담당자이메일 : {0} ▶ {1} ", sPRV_CHRG_EMAIL, sChrgEmail);
                    //    }

                    //    if (!sPRV_CHRG_RGN_NO.Equals(sChrgRgnTel))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("지역번호 : {0} ▶ {1} ", sPRV_CHRG_RGN_NO, sChrgRgnTel);
                    //    }

                    //    if (!sPRV_CHRG_HP_NO.Equals(sChrgHP))
                    //    {
                    //        iLogCnt++;
                    //        sLogRmk += string.Format("담당자휴대폰 : {0} ▶ {1} ", sPRV_CHRG_HP_NO, sChrgHP);
                    //    }
                    //}

                    //sLogRmk = sLogRmk.Length > 500 ? sLogRmk.Substring(0, 500) : sLogRmk;

                    //dicParams.Add("EDIT_RMK", sLogRmk);
                    #endregion

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    strSql.Clear();
                    strSql.AppendLine("SET IDENTITY_INSERT ACC_DEALER_CD ON");
                    strSql.AppendLine("IF EXISTS(SELECT* FROM ACC_DEALER_CD WHERE DEALER_CD ="+ sDealerCd + " )");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         UPDATE ACC_DEALER_CD");
                    strSql.AppendLine("            SET IDT_NO = '" + sIdtNo + "' ");
                    strSql.AppendLine("              , STRT_YMD = '" + sStartDay + "' ");
                    strSql.AppendLine("              , END_YMD = '" + sEndDay + "' ");
                    strSql.AppendLine("              , DEALER_GB = '" + sDealerGB + "' ");
                    strSql.AppendLine("              , DEALER_NM = '" + sDealerName + "' ");
                    strSql.AppendLine("              , INITIAL_NM = '" + sInitialNm + "' ");
                    strSql.AppendLine("              , REP_NM = '" + sRepName + "' ");
                    strSql.AppendLine("              , BIZ_NM   = '" + sBizName + "' ");
                    strSql.AppendLine("              , TYPE_NM   = '" + sTypeName + "' ");
                    strSql.AppendLine("              , ZIP   = '" + sZipNum + "' ");
                    strSql.AppendLine("              , ZIP_SEQ = '" + sZipSeq + "' ");
                    strSql.AppendLine("              , ADDR   = '" + sAddr + "' ");
                    strSql.AppendLine("              , DTL_ADDR   = '" + sDtlAddr + "' ");
                    strSql.AppendLine("              , J_REGION = '" + sRegion + "' ");
                    strSql.AppendLine("              , BANK_CD   = '" + sBankCd + "' ");
                    strSql.AppendLine("              , BANK_ACNT_NO   = '" + sBankNM + "' ");
                    strSql.AppendLine("              , ACNT_HOLDER   = '" + sBankHolder + "' ");
                    strSql.AppendLine("              , EMAIL   = '" + sEMail + "' ");
                    strSql.AppendLine("              , HOMEPG   = '" + sHomePG + "' ");
                    strSql.AppendLine("              , FAX = '" + sFax + "' ");
                    strSql.AppendLine("              , WEB_FAX_YN = '" + sWebFaxYn + "' ");
                    strSql.AppendLine("              , SCM_PSWD = '" + sSCMPswd + "' ");
                    strSql.AppendLine("              , CHRG_ID   = '" + sChrgID + "' ");
                    strSql.AppendLine("              , CHRG_NM   = '" + sChrgNM + "' ");
                    strSql.AppendLine("              , CHRG_TEL_NO   = '" + sChrgTel + "' ");
                    strSql.AppendLine("              , CHRG_RGN_NO = '" + sChrgRgnTel + "' ");
                    strSql.AppendLine("              , CHRG_HP_NO   = '" + sChrgHP + "' ");
                    strSql.AppendLine("              , CHRG_EMAIL   = '" + sChrgEmail + "' ");
                    strSql.AppendLine("              , BFR_DEALER_CD   = '" + sPreDealerCd + "' ");
                    strSql.AppendLine("              , BFR_DEALER_NM   = '" + sPreDealerNM + "' ");
                    strSql.AppendLine("              , BILL_KIND   = '" + sBillKind + "' ");
                    strSql.AppendLine("              , BILL_ISSUE_GB   = '" + sBlIssue + "' ");
                    strSql.AppendLine("              , PAY_GB   = '" + sPayGB + "' ");
                    strSql.AppendLine("              , PAY_TERM   = '" + sPayTerm + "' ");
                    strSql.AppendLine("              , SPPL_LIMIT_DCNT   = '" + sLimitDcnt + "' ");
                    strSql.AppendLine("              , EOB_YN = '" + sEobYn + "' ");
                    strSql.AppendLine("              , NOTE   = '" + sNote + "' ");
                    strSql.AppendLine("              , BANKYN = '" + sBANKYN + "'");
                    strSql.AppendLine("              , MFY_DT = CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine("              , MFY_ID   = '" + sMfyId + "' ");
                    strSql.AppendLine("          WHERE DEALER_CD = "+ sDealerCd + "");
                    strSql.AppendLine("   END");
                    strSql.AppendLine("ELSE");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine(" INSERT INTO ACC_DEALER_CD ");
                    strSql.AppendLine("           ( DEALER_CD");
                    strSql.AppendLine("           , IDT_NO");
                    strSql.AppendLine("           , CORP_NO");
                    strSql.AppendLine("           , STRT_YMD ");
                    strSql.AppendLine("           , END_YMD ");
                    strSql.AppendLine("           , DEALER_GB ");
                    strSql.AppendLine("           , DEALER_NM ");
                    strSql.AppendLine("           , INITIAL_NM ");
                    strSql.AppendLine("           , REP_NM ");
                    strSql.AppendLine("           , BIZ_NM ");
                    strSql.AppendLine("           , TYPE_NM ");
                    strSql.AppendLine("           , ZIP ");
                    strSql.AppendLine("           , ZIP_SEQ ");
                    strSql.AppendLine("           , ADDR ");
                    strSql.AppendLine("           , DTL_ADDR ");
                    strSql.AppendLine("           , J_REGION ");
                    strSql.AppendLine("           , BANK_CD ");
                    strSql.AppendLine("           , BANK_ACNT_NO ");
                    strSql.AppendLine("           , ACNT_HOLDER ");
                    strSql.AppendLine("           , EMAIL ");
                    strSql.AppendLine("           , HOMEPG ");
                    strSql.AppendLine("           , FAX");
                    strSql.AppendLine("           , WEB_FAX_YN ");
                    strSql.AppendLine("           , SCM_PSWD ");
                    strSql.AppendLine("           , CHRG_ID ");
                    strSql.AppendLine("           , CHRG_NM ");
                    strSql.AppendLine("           , CHRG_TEL_NO ");
                    strSql.AppendLine("           , CHRG_RGN_NO ");
                    strSql.AppendLine("           , CHRG_HP_NO ");
                    strSql.AppendLine("           , CHRG_EMAIL ");
                    strSql.AppendLine("           , BFR_DEALER_CD ");
                    strSql.AppendLine("           , BFR_DEALER_NM ");
                    strSql.AppendLine("           , BILL_KIND ");
                    strSql.AppendLine("           , BILL_ISSUE_GB ");
                    strSql.AppendLine("           , PAY_GB ");
                    strSql.AppendLine("           , PAY_TERM ");
                    strSql.AppendLine("           , SPPL_LIMIT_DCNT ");
                    strSql.AppendLine("           , EOB_YN ");
                    strSql.AppendLine("           , NOTE ");
                    strSql.AppendLine("           , BANKYN");
                    strSql.AppendLine("           , ENT_DT ");
                    strSql.AppendLine("           , ENT_ID ");
                    strSql.AppendLine("           ) ");
                    strSql.AppendLine("     VALUES( " + sDealerCd + " ");
                    strSql.AppendLine("           , '" + sIdtNo + "'");
                    strSql.AppendLine("           , '" + sCorpNo + "'");
                    strSql.AppendLine("           , '" + sStartDay + "' ");
                    strSql.AppendLine("           , '" + sEndDay + "' ");
                    strSql.AppendLine("           , '" + sDealerGB + "' ");
                    strSql.AppendLine("           , '" + sDealerName + "' ");
                    strSql.AppendLine("           , '" + sInitialNm + "' ");
                    strSql.AppendLine("           , '" + sRepName + "' ");
                    strSql.AppendLine("           , '" + sBizName + "' ");
                    strSql.AppendLine("           , '" + sTypeName + "' ");
                    strSql.AppendLine("           , '" + sZipNum + "' ");
                    strSql.AppendLine("           , '" + sZipSeq + "' ");
                    strSql.AppendLine("           , '" + sAddr + "' ");
                    strSql.AppendLine("           , '" + sDtlAddr + "' ");
                    strSql.AppendLine("           , '" + sRegion + "' ");
                    strSql.AppendLine("           , '" + sBankCd + "' ");
                    strSql.AppendLine("           , '" + sBankNM + "' ");
                    strSql.AppendLine("           , '" + sBankHolder + "' ");
                    strSql.AppendLine("           , '" + sEMail + "' ");
                    strSql.AppendLine("           , '" + sHomePG + "' ");
                    strSql.AppendLine("           , '" + sFax + "' ");
                    strSql.AppendLine("           , '" + sWebFaxYn + "' ");
                    strSql.AppendLine("           , '" + sSCMPswd + "' ");
                    strSql.AppendLine("           , '" + sChrgID + "' ");
                    strSql.AppendLine("           , '" + sChrgNM + "' ");
                    strSql.AppendLine("           , '" + sChrgTel + "' ");
                    strSql.AppendLine("           , '" + sChrgRgnTel + "'  ");
                    strSql.AppendLine("           , '" + sChrgHP + "' ");
                    strSql.AppendLine("           , '" + sChrgEmail + "' ");
                    strSql.AppendLine("           , '" + sPreDealerCd + "' ");
                    strSql.AppendLine("           , '" + sPreDealerNM + "' ");
                    strSql.AppendLine("           , '" + sBillKind + "' ");
                    strSql.AppendLine("           , '" + sBlIssue + "' ");
                    strSql.AppendLine("           , '" + sPayGB + "' ");
                    strSql.AppendLine("           , '" + sPayTerm + "' ");
                    strSql.AppendLine("           , '" + sLimitDcnt + "' ");
                    strSql.AppendLine("           , '" + sEobYn + "' ");
                    strSql.AppendLine("           , '" + sNote + "' ");
                    strSql.AppendLine("           , '" + sBANKYN + "'");
                    strSql.AppendLine("           , CONVERT(VARCHAR(19),GETDATE(),20) ");
                    strSql.AppendLine("           , '" + sEntID + "') ");
                    strSql.AppendLine("   END");
                    strSql.AppendLine("SET IDENTITY_INSERT ACC_DEALER_CD OFF");

                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO ACC_DEALER_CD ");
                    //strSql.AppendLine("           ( DEALER_CD");
                    //strSql.AppendLine("           , IDT_NO");
                    //strSql.AppendLine("           , CORP_NO");
                    //strSql.AppendLine("           , STRT_YMD ");
                    //strSql.AppendLine("           , END_YMD ");
                    //strSql.AppendLine("           , DEALER_GB ");
                    //strSql.AppendLine("           , DEALER_NM ");
                    //strSql.AppendLine("           , INITIAL_NM ");
                    //strSql.AppendLine("           , REP_NM ");
                    //strSql.AppendLine("           , BIZ_NM ");
                    //strSql.AppendLine("           , TYPE_NM ");
                    //strSql.AppendLine("           , ZIP ");
                    //strSql.AppendLine("           , ZIP_SEQ ");
                    //strSql.AppendLine("           , ADDR ");
                    //strSql.AppendLine("           , DTL_ADDR ");
                    //strSql.AppendLine("           , J_REGION ");
                    //strSql.AppendLine("           , BANK_CD ");
                    //strSql.AppendLine("           , BANK_ACNT_NO ");
                    //strSql.AppendLine("           , ACNT_HOLDER ");
                    //strSql.AppendLine("           , EMAIL ");
                    //strSql.AppendLine("           , HOMEPG ");
                    //strSql.AppendLine("           , FAX");
                    //strSql.AppendLine("           , WEB_FAX_YN ");
                    //strSql.AppendLine("           , SCM_PSWD ");
                    //strSql.AppendLine("           , CHRG_ID ");
                    //strSql.AppendLine("           , CHRG_NM ");
                    //strSql.AppendLine("           , CHRG_TEL_NO ");
                    //strSql.AppendLine("           , CHRG_RGN_NO ");
                    //strSql.AppendLine("           , CHRG_HP_NO ");
                    //strSql.AppendLine("           , CHRG_EMAIL ");
                    //strSql.AppendLine("           , BFR_DEALER_CD ");
                    //strSql.AppendLine("           , BFR_DEALER_NM ");
                    //strSql.AppendLine("           , BILL_KIND ");
                    //strSql.AppendLine("           , BILL_ISSUE_GB ");
                    //strSql.AppendLine("           , PAY_GB ");
                    //strSql.AppendLine("           , PAY_TERM ");
                    //strSql.AppendLine("           , SPPL_LIMIT_DCNT ");
                    //strSql.AppendLine("           , EOB_YN ");
                    //strSql.AppendLine("           , NOTE ");
                    //strSql.AppendLine("           , ENT_DT ");
                    //strSql.AppendLine("           , ENT_ID ");
                    //strSql.AppendLine("           , MFY_DT ");
                    //strSql.AppendLine("           , MFY_ID ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("     VALUES( '" + sDealerCd + "' ");
                    //strSql.AppendLine("           , '" + sIdtNo + "' ");
                    //strSql.AppendLine("           , '"+ sCorpNo + "'");
                    //strSql.AppendLine("           , '" + sStartDay + "' ");
                    //strSql.AppendLine("           , '" + sEndDay + "' ");
                    //strSql.AppendLine("           , '" + sDealerGB + "' ");
                    //strSql.AppendLine("           , '" + sDealerName + "' ");
                    //strSql.AppendLine("           , '" + sInitialNm + "' ");
                    //strSql.AppendLine("           , '" + sRepName + "' ");
                    //strSql.AppendLine("           , '" + sBizName + "' ");
                    //strSql.AppendLine("           , '" + sTypeName + "' ");
                    //strSql.AppendLine("           , '" + sZipNum + "' ");
                    //strSql.AppendLine("           , '" + sZipSeq + "' ");
                    //strSql.AppendLine("           , '" + sAddr + "' ");
                    //strSql.AppendLine("           , '" + sDtlAddr + "' ");
                    //strSql.AppendLine("           , '" + sRegion + "' ");
                    //strSql.AppendLine("           , '" + sBankCd + "' ");
                    //strSql.AppendLine("           , '" + sBankNM + "' ");
                    //strSql.AppendLine("           , '" + sBankHolder + "' ");
                    //strSql.AppendLine("           , '" + sEMail + "' ");
                    //strSql.AppendLine("           , '" + sHomePG + "' ");
                    //strSql.AppendLine("           , '" + sFax + "' ");
                    //strSql.AppendLine("           , '" + sWebFaxYn + "' ");
                    //strSql.AppendLine("           , '" + sSCMPswd + "' ");
                    //strSql.AppendLine("           , '" + sChrgID + "' ");
                    //strSql.AppendLine("           , '" + sChrgNM + "' ");
                    //strSql.AppendLine("           , '" + sChrgTel + "' ");
                    //strSql.AppendLine("           , '" + sChrgRgnTel + "'  ");
                    //strSql.AppendLine("           , '" + sChrgHP + "' ");
                    //strSql.AppendLine("           , '" + sChrgEmail + "' ");
                    //strSql.AppendLine("           , '" + sPreDealerCd + "' ");
                    //strSql.AppendLine("           , '" + sPreDealerNM + "' ");
                    //strSql.AppendLine("           , '" + sBillKind + "' ");
                    //strSql.AppendLine("           , '" + sBlIssue + "' ");
                    //strSql.AppendLine("           , '" + sPayGB + "' ");
                    //strSql.AppendLine("           , '" + sPayTerm + "' ");
                    //strSql.AppendLine("           , '" + sLimitDcnt + "' ");
                    //strSql.AppendLine("           , '" + sEobYn + "' ");
                    //strSql.AppendLine("           , '" + sNote + "' ");
                    //strSql.AppendLine("           , NOW() ");
                    //strSql.AppendLine("           , '" + sEntID + "' ");
                    //strSql.AppendLine("           , NOW() ");
                    //strSql.AppendLine("           , '" + sMfyId + "' )");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE IDT_NO = '" + sIdtNo + "' ");
                    //strSql.AppendLine("           , STRT_YMD = '" + sStartDay + "' ");
                    //strSql.AppendLine("           , END_YMD = '" + sEndDay + "' ");
                    //strSql.AppendLine("           , DEALER_GB = '" + sDealerGB + "' ");
                    //strSql.AppendLine("           , DEALER_NM = '" + sDealerName + "' ");
                    //strSql.AppendLine("           , INITIAL_NM = '" + sInitialNm + "' ");
                    //strSql.AppendLine("           , REP_NM = '" + sRepName + "' ");
                    //strSql.AppendLine("           , BIZ_NM   = '" + sBizName + "' ");
                    //strSql.AppendLine("           , TYPE_NM   = '" + sTypeName + "' ");
                    //strSql.AppendLine("           , ZIP   = '" + sZipNum + "' ");
                    //strSql.AppendLine("           , ZIP_SEQ = '" + sZipSeq + "' ");
                    //strSql.AppendLine("           , ADDR   = '" + sAddr + "' ");
                    //strSql.AppendLine("           , DTL_ADDR   = '" + sDtlAddr + "' ");
                    //strSql.AppendLine("           , J_REGION = '" + sRegion + "' ");
                    //strSql.AppendLine("           , BANK_CD   = '" + sBankCd + "' ");
                    //strSql.AppendLine("           , BANK_ACNT_NO   = '" + sBankNM + "' ");
                    //strSql.AppendLine("           , ACNT_HOLDER   = '" + sBankHolder + "' ");
                    //strSql.AppendLine("           , EMAIL   = '" + sEMail + "' ");
                    //strSql.AppendLine("           , HOMEPG   = '" + sHomePG + "' ");
                    //strSql.AppendLine("           , FAX = '" + sFax + "' ");
                    //strSql.AppendLine("           , WEB_FAX_YN = '" + sWebFaxYn + "' ");
                    //strSql.AppendLine("           , SCM_PSWD = '" + sSCMPswd + "' ");
                    //strSql.AppendLine("           , CHRG_ID   = '" + sChrgID + "' ");
                    //strSql.AppendLine("           , CHRG_NM   = '" + sChrgNM + "' ");
                    //strSql.AppendLine("           , CHRG_TEL_NO   = '" + sChrgTel + "' ");
                    //strSql.AppendLine("           , CHRG_RGN_NO = '" + sChrgRgnTel + "' ");
                    //strSql.AppendLine("           , CHRG_HP_NO   = '" + sChrgHP + "' ");
                    //strSql.AppendLine("           , CHRG_EMAIL   = '" + sChrgEmail + "' ");
                    //strSql.AppendLine("           , BFR_DEALER_CD   = '" + sPreDealerCd + "' ");
                    //strSql.AppendLine("           , BFR_DEALER_NM   = '" + sPreDealerNM + "' ");
                    //strSql.AppendLine("           , BILL_KIND   = '" + sBillKind + "' ");
                    //strSql.AppendLine("           , BILL_ISSUE_GB   = '" + sBlIssue + "' ");
                    //strSql.AppendLine("           , PAY_GB   = '" + sPayGB + "' ");
                    //strSql.AppendLine("           , PAY_TERM   = '" + sPayTerm + "' ");
                    //strSql.AppendLine("           , SPPL_LIMIT_DCNT   = '" + sLimitDcnt + "' ");
                    //strSql.AppendLine("           , EOB_YN = '" + sEobYn + "' ");
                    //strSql.AppendLine("           , NOTE   = '" + sNote + "' ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID   = '" + sMfyId + "' ");
                    #endregion

                    #region[2021-03-12 로그 로직 변경 이전쿼리]

                    //strSql.AppendLine(" IF ( SELECT 1 = 1 FROM ACC_DEALER_CD WHERE DEALER_CD = @DEALER_CD ) THEN ");
                    //strSql.AppendLine("    BEGIN #EXISTS ");
                    //strSql.AppendLine("           ");
                    //strSql.AppendLine("          UPDATE ACC_DEALER_CD ");
                    //strSql.AppendLine("             SET IDT_NO          = @IDT_NO          ");
                    //strSql.AppendLine("               , CORP_NO         = @CORP_NO ");
                    //strSql.AppendLine("               , STRT_YMD        = @STRT_YMD        ");
                    //strSql.AppendLine("               , END_YMD         = @END_YMD         ");
                    //strSql.AppendLine("               , DEALER_GB       = @DEALER_GB       ");
                    //strSql.AppendLine("               , DEALER_NM       = @DEALER_NM       ");
                    //strSql.AppendLine("               , INITIAL_NM      = @INITIAL_NM      ");
                    //strSql.AppendLine("               , REP_NM          = @REP_NM          ");
                    //strSql.AppendLine("               , BIZ_NM          = @BIZ_NM          ");
                    //strSql.AppendLine("               , TYPE_NM         = @TYPE_NM         ");
                    //strSql.AppendLine("               , ZIP             = @ZIP             ");
                    //strSql.AppendLine("               , ZIP_SEQ         = @ZIP_SEQ         ");
                    //strSql.AppendLine("               , ADDR            = @ADDR            ");
                    //strSql.AppendLine("               , DTL_ADDR        = @DTL_ADDR        ");
                    //strSql.AppendLine("               , J_REGION        = @J_REGION        ");
                    //strSql.AppendLine("               , BANK_CD         = @BANK_CD         ");
                    //strSql.AppendLine("               , BANK_ACNT_NO    = @BANK_ACNT_NO    ");
                    //strSql.AppendLine("               , ACNT_HOLDER     = @ACNT_HOLDER     ");
                    //strSql.AppendLine("               , EMAIL           = @EMAIL           ");
                    //strSql.AppendLine("               , HOMEPG          = @HOMEPG          ");
                    //strSql.AppendLine("               , FAX             = @FAX             ");
                    //strSql.AppendLine("               , WEB_FAX_YN      = @WEB_FAX_YN      ");
                    //strSql.AppendLine("               , SCM_PSWD        = @SCM_PSWD        ");
                    //strSql.AppendLine("               , CHRG_ID         = @CHRG_ID         ");
                    //strSql.AppendLine("               , CHRG_NM         = @CHRG_NM         ");
                    //strSql.AppendLine("               , CHRG_TEL_NO     = @CHRG_TEL_NO     ");
                    //strSql.AppendLine("               , CHRG_RGN_NO     = @CHRG_RGN_NO     ");
                    //strSql.AppendLine("               , CHRG_HP_NO      = @CHRG_HP_NO      ");
                    //strSql.AppendLine("               , CHRG_EMAIL      = @CHRG_EMAIL      ");
                    //strSql.AppendLine("               , BFR_DEALER_CD   = @BFR_DEALER_CD   ");
                    //strSql.AppendLine("               , BFR_DEALER_NM   = @BFR_DEALER_NM   ");
                    //strSql.AppendLine("               , BILL_KIND       = @BILL_KIND       ");
                    //strSql.AppendLine("               , BILL_ISSUE_GB   = @BILL_ISSUE_GB   ");
                    //strSql.AppendLine("               , PAY_GB          = @PAY_GB          ");
                    //strSql.AppendLine("               , PAY_TERM        = @PAY_TERM        ");
                    //strSql.AppendLine("               , SPPL_LIMIT_DCNT = @SPPL_LIMIT_DCNT ");
                    //strSql.AppendLine("               , EOB_YN          = @EOB_YN          ");
                    //strSql.AppendLine("               , NOTE            = @NOTE            ");
                    //strSql.AppendLine("               , MFY_DT          = @MFT_DT  ");
                    //strSql.AppendLine("               , MFY_ID          = @MFY_ID ");
                    //strSql.AppendLine("           WHERE DEALER_CD = @DEALER_CD; ");
                    //strSql.AppendLine("     ");
                    //strSql.AppendLine("          INSERT INTO ZSYS_LOG ");
                    //strSql.AppendLine("          	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                    //strSql.AppendLine("          	      VALUES ");
                    //strSql.AppendLine("          	           ( @CUR_TIME ");
                    //strSql.AppendLine("          	           , @USRCD ");
                    //strSql.AppendLine("          	           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    //strSql.AppendLine("                               FROM DAEJIERP.ZSYS_LOG X1 ");
                    //strSql.AppendLine("                              WHERE X1.OCCUR_DATE = @CUR_TIME ");
                    //strSql.AppendLine("                                AND X1.USRCD = @USRCD ) #LOG_SEQ(구분자) ");
                    //strSql.AppendLine("          	           , @PGMID ");
                    //strSql.AppendLine("          	           , 'U' ");
                    //strSql.AppendLine("          	           , @IP ");
                    //strSql.AppendLine("          	           , @EDIT_RMK ); ");
                    //strSql.AppendLine("      END; ");
                    //strSql.AppendLine(" ELSE ");
                    //strSql.AppendLine("    BEGIN #NON_EXISTS ");
                    //strSql.AppendLine("          INSERT INTO ACC_DEALER_CD  ");
                    //strSql.AppendLine("                    ( DEALER_CD      , IDT_NO      , STRT_YMD      , END_YMD   , DEALER_GB  ");
                    //strSql.AppendLine("                    , CORP_NO "); //2021-01-20 (법인번호 추가)
                    //strSql.AppendLine("                    , DEALER_NM      , INITIAL_NM  , REP_NM        , BIZ_NM    , TYPE_NM  ");
                    //strSql.AppendLine("                    , ZIP            , ZIP_SEQ     , ADDR          , DTL_ADDR  , J_REGION  ");
                    //strSql.AppendLine("                    , BANK_CD        , BANK_ACNT_NO, ACNT_HOLDER   , EMAIL     , HOMEPG  ");
                    //strSql.AppendLine("                    , FAX            , WEB_FAX_YN  , SCM_PSWD      , CHRG_ID   , CHRG_NM  ");
                    //strSql.AppendLine("                    , CHRG_TEL_NO    , CHRG_RGN_NO , CHRG_HP_NO    , CHRG_EMAIL, BFR_DEALER_CD  ");
                    //strSql.AppendLine("                    , BFR_DEALER_NM  , BILL_KIND   , BILL_ISSUE_GB , PAY_GB    , PAY_TERM  ");
                    //strSql.AppendLine(" 			       , SPPL_LIMIT_DCNT, EOB_YN      , NOTE          , ENT_DT    , ENT_ID )  ");
                    //strSql.AppendLine("              VALUES( @DEALER_CD      , @IDT_NO      , @STRT_YMD      , @END_YMD   , @DEALER_GB  ");
                    //strSql.AppendLine("                    , @CORP_NO ");
                    //strSql.AppendLine("                    , @DEALER_NM      , @INITIAL_NM  , @REP_NM        , @BIZ_NM    , @TYPE_NM  ");
                    //strSql.AppendLine("                    , @ZIP            , @ZIP_SEQ     , @ADDR          , @DTL_ADDR  , @J_REGION  ");
                    //strSql.AppendLine("                    , @BANK_CD        , @BANK_ACNT_NO, @ACNT_HOLDER   , @EMAIL     , @HOMEPG  ");
                    //strSql.AppendLine("                    , @FAX            , @WEB_FAX_YN  , @SCM_PSWD      , @CHRG_ID   , @CHRG_NM  ");
                    //strSql.AppendLine("                    , @CHRG_TEL_NO    , @CHRG_RGN_NO , @CHRG_HP_NO    , @CHRG_EMAIL, @BFR_DEALER_CD  ");
                    //strSql.AppendLine("                    , @BFR_DEALER_NM  , @BILL_KIND   , @BILL_ISSUE_GB , @PAY_GB    , @PAY_TERM  ");
                    //strSql.AppendLine(" 	 	       	   , @SPPL_LIMIT_DCNT, @EOB_YN      , @NOTE          , @ENT_DT    , @ENT_ID ); ");
                    //strSql.AppendLine("           ");
                    //strSql.AppendLine("          INSERT INTO ZSYS_LOG ");
                    //strSql.AppendLine("          	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                    //strSql.AppendLine("          	      VALUES ");
                    //strSql.AppendLine("          	           ( @CUR_TIME ");
                    //strSql.AppendLine("          	           , @USRCD ");
                    //strSql.AppendLine("          	           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    //strSql.AppendLine("                               FROM DAEJIERP.ZSYS_LOG X1 ");
                    //strSql.AppendLine("                              WHERE X1.OCCUR_DATE = @CUR_TIME ");
                    //strSql.AppendLine("                                AND X1.USRCD = @USRCD ) #LOG_SEQ(구분자) ");
                    //strSql.AppendLine("          	           , @PGMID ");
                    //strSql.AppendLine("          	           , 'I' ");
                    //strSql.AppendLine("          	           , @IP ");
                    //strSql.AppendLine("          	           , @EDIT_RMK ); ");
                    //strSql.AppendLine("      END; ");
                    //strSql.AppendLine(" END IF; ");

                    #endregion[2021-03-12 로그 로직 변경 이전쿼리]

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.ExecuteNonQuery();

                    #region 이전코드(기간모름)
                    //LogInsert
                    //Reference : #00001
                    //if (iLogCnt > 0)
                    //{
                    //    /*
                    //     * 2021-03-12
                    //     * #00002
                    //     * 현업에서는 금액부분과 관련된 로직체크를 원함
                    //     * 기초정보인 거래처관리는 로그처리 우선 주석처리 진행
                    //     * 추후 요청 시 아래 INSERT 로그로직 주석해제하여 사용할 것
                    //     */
                    //    //strSql.Clear();
                    //    //strSql.AppendLine(" ");
                    //    //strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                    //    //strSql.AppendLine("           ( OCCUR_DATE ");
                    //    //strSql.AppendLine("           , USRCD ");
                    //    //strSql.AppendLine("           , LOG_SEQ ");
                    //    //strSql.AppendLine("           , EDIT_KIND ");
                    //    //strSql.AppendLine("           , PGM_ID ");
                    //    //strSql.AppendLine("           , ACS_IP ");
                    //    //strSql.AppendLine("           , EDIT_RMK ) ");
                    //    //strSql.AppendLine("     VALUES( @OCCUR_DATE ");
                    //    //strSql.AppendLine("           , @USRCD ");
                    //    //strSql.AppendLine("           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                    //    //strSql.AppendLine("           , 'U' ");
                    //    //strSql.AppendLine("           , @PGM_ID ");
                    //    //strSql.AppendLine("           , @ACS_IP ");
                    //    //strSql.AppendLine("           , @EDIT_RMK ) ");

                    //    //cmd.CommandType = CommandType.Text;
                    //    //cmd.CommandText = strSql.ToString();
                    //    //cmd.Parameters.Clear();
                    //    //cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    //    //cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                    //    //cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                    //    //cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                    //    //cmd.Parameters.AddWithValue("@EDIT_RMK", sLogRmk);
                    //    //cmd.ExecuteNonQuery();
                    //    //cmd.Parameters.Clear();
                    //}
                    #endregion

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show("저장을 완료했습니다.");
                    ChkAutoNo.ReadOnly = true;

                    ClearFps();
                    BtnRetr.PerformClick();
                    GridViewDealerMng.FocusedRowHandle = GridViewDealerMng.LocateByDisplayText(0, GridColDealerCd, sDealerCd);
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show(ex.Message);
                }

                Cursor = Cursors.Default;
            }

            #region[2020-12-29 로그테스트 이전로직]

            //if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            //{
            //    XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
            //    return;
            //}

            //if (TxtDealerNM.Text == "")
            //{
            //    XtraMessageBox.Show("거래처명을 쓰세요.");
            //    TxtDealerNM.Focus();
            //    return;
            //}
            //else if (DateEditFrom.EditValue.ToString() == "" || DateEditTo.EditValue.ToString() == "")
            //{
            //    XtraMessageBox.Show("시작일자 및 종료일자를 선택하세요.");
            //    DateEditFrom.Focus();
            //    return;
            //}
            //else if (TxtPayTerm.Text.Length > 2)
            //{
            //    XtraMessageBox.Show("결제기간 개월 수는 두 자리를 초과할 수 없습니다");
            //    return;
            //}
            //else if (TxtLtDcnt.Text.Length > 4)
            //{
            //    XtraMessageBox.Show("납품기한일수는 네 자리를 초과할 수 없습니다. ");
            //    return;
            //}
            //else if (string.IsNullOrEmpty(TxtDealerCd.EditValue?.ToString()))
            //{
            //    XtraMessageBox.Show("거래처코드가 존재하지 않습니다.");
            //    TxtDealerCd.Focus();
            //    return;
            //}
            //else
            //{
            //    Cursor = Cursors.WaitCursor;
            //    if (!ChkAutoNo.Checked)
            //    {
            //        if (FindDealerCd(TxtDealerCd.EditValue?.ToString()))
            //        {
            //            Cursor = Cursors.Default;
            //            string sMsg = string.Format("거래처코드 : {0}\r\n해당 거래처코드는 이미 존재합니다.\r\n다른 거래처코드를 입력하세요.", TxtDealerCd.EditValue?.ToString());
            //            XtraMessageBox.Show(sMsg);
            //            TxtDealerCd.SelectAll();
            //            TxtDealerCd.Focus();
            //            return;
            //        }
            //    }

            //    string sDealerCd = TxtDealerCd.EditValue?.ToString();
            //    string sIdtNo = TxtIdtNo.EditValue?.ToString().Replace("-", "");
            //    //if (!CheckSameChar(TxtIdtNo.Text.Replace("-", "").ToCharArray()))
            //    //    sIdtNo = string.Empty;

            //    string sDealerName = TxtDealerNM.Text;
            //    string sDealerGB = CbxDealerGB.EditValue.ToString();
            //    string sInitialNm = TxtInitialNm.EditValue?.ToString();
            //    string sBizName = TxtBizNM.Text;
            //    string sRepName = TxtRepNM.Text;
            //    string sTypeName = TxtTypeNM.Text;
            //    string sZipNum = TxtZipNum.Text;
            //    string sZipSeq = string.Empty;
            //    string sAddr = TxtAddr.Text;
            //    string sDtlAddr = TxtDtlAddr.Text;
            //    string sRegion = TxtRegion.Text;
            //    string sEMail = TxtEmail.Text;
            //    string sHomePG = TxtHomePG.Text;
            //    string sFax = TxtFaxNum.Text;
            //    string sWebFaxYn = RdgbWebFaxGb.EditValue?.ToString();
            //    string sSCMPswd = TxtScmPw.EditValue?.ToString();
            //    string sBankCd = LkupBankNM.EditValue?.ToString();
            //    if (sBankCd == "****") sBankCd = "";
            //    string sBankHolder = TxtHolder.Text;
            //    string sBankNM = TxtBankNum.Text;
            //    string sChrgID = LkupEditChrgID.EditValue?.ToString();
            //    if (sChrgID == "****") sChrgID = "";
            //    string sChrgNM = TxtChrgNM.Text;
            //    string sChrgHP = TxtChrgHP.Text;
            //    string sChrgEmail = TxtChrgEmail.Text;
            //    string sChrgTel = TxtChrgTel.Text;
            //    string sChrgRgnTel = TxtChrgRgnNo.EditValue?.ToString().Trim();
            //    string sBillKind = LkupEditBillKind.EditValue?.ToString();
            //    if (sBillKind == "****") sBillKind = "";
            //    string sBlIssue = LkupBIssueGB.EditValue?.ToString();
            //    if (sBlIssue == "****") sBlIssue = "";
            //    string sPayGB = LkupEditPayGB.EditValue?.ToString();
            //    if (sPayGB == "****") sPayGB = "";
            //    string sPayTerm = TxtPayTerm.Text;
            //    if (TxtPayTerm.Text == "") sPayTerm = "0";
            //    string sLimitDcnt = TxtLtDcnt.Text;
            //    if (TxtLtDcnt.Text == "") sLimitDcnt = "0";
            //    string sEobYn = RdgbEobYn.EditValue?.ToString();
            //    string sNote = TxtNote.Text;
            //    string sStartDay = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            //    string sEndDay = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            //    string sPreDealerCd = TxtBfrDlrCd.Text;
            //    string sPreDealerNM = TxtBfrDlrNM.Text;
            //    string sEntID = FmMainToolBar2.UserID;
            //    string sMfyId = FmMainToolBar2.UserID;

            //    ClearFps();

            //    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            //    SqlCommand cmd = DBConn.dbCon.CreateCommand();
            //    cmd.Transaction = DBConn.dbTran;

            //    StringBuilder strSql = new StringBuilder();

            //    strSql.AppendLine(" INSERT INTO ACC_DEALER_CD ");
            //    strSql.AppendLine("           ( DEALER_CD");
            //    strSql.AppendLine("           , IDT_NO");
            //    strSql.AppendLine("           , STRT_YMD ");
            //    strSql.AppendLine("           , END_YMD ");
            //    strSql.AppendLine("           , DEALER_GB ");
            //    strSql.AppendLine("           , DEALER_NM ");
            //    strSql.AppendLine("           , INITIAL_NM ");
            //    strSql.AppendLine("           , REP_NM ");
            //    strSql.AppendLine("           , BIZ_NM ");
            //    strSql.AppendLine("           , TYPE_NM ");
            //    strSql.AppendLine("           , ZIP ");
            //    strSql.AppendLine("           , ZIP_SEQ ");
            //    strSql.AppendLine("           , ADDR ");
            //    strSql.AppendLine("           , DTL_ADDR ");
            //    strSql.AppendLine("           , J_REGION ");
            //    strSql.AppendLine("           , BANK_CD ");
            //    strSql.AppendLine("           , BANK_ACNT_NO ");
            //    strSql.AppendLine("           , ACNT_HOLDER ");
            //    strSql.AppendLine("           , EMAIL ");
            //    strSql.AppendLine("           , HOMEPG ");
            //    strSql.AppendLine("           , FAX");
            //    strSql.AppendLine("           , WEB_FAX_YN ");
            //    strSql.AppendLine("           , SCM_PSWD ");
            //    strSql.AppendLine("           , CHRG_ID ");
            //    strSql.AppendLine("           , CHRG_NM ");
            //    strSql.AppendLine("           , CHRG_TEL_NO ");
            //    strSql.AppendLine("           , CHRG_RGN_NO ");
            //    strSql.AppendLine("           , CHRG_HP_NO ");
            //    strSql.AppendLine("           , CHRG_EMAIL ");
            //    strSql.AppendLine("           , BFR_DEALER_CD ");
            //    strSql.AppendLine("           , BFR_DEALER_NM ");
            //    strSql.AppendLine("           , BILL_KIND ");
            //    strSql.AppendLine("           , BILL_ISSUE_GB ");
            //    strSql.AppendLine("           , PAY_GB ");
            //    strSql.AppendLine("           , PAY_TERM ");
            //    strSql.AppendLine("           , SPPL_LIMIT_DCNT ");
            //    strSql.AppendLine("           , EOB_YN ");
            //    strSql.AppendLine("           , NOTE ");
            //    strSql.AppendLine("           , ENT_DT ");
            //    strSql.AppendLine("           , ENT_ID ");
            //    strSql.AppendLine("           , MFY_DT ");
            //    strSql.AppendLine("           , MFY_ID ");
            //    strSql.AppendLine("           ) ");
            //    strSql.AppendLine("     VALUES( '" + sDealerCd + "' ");
            //    strSql.AppendLine("           , '" + sIdtNo + "' ");
            //    strSql.AppendLine("           , '" + sStartDay + "' ");
            //    strSql.AppendLine("           , '" + sEndDay + "' ");
            //    strSql.AppendLine("           , '" + sDealerGB + "' ");
            //    strSql.AppendLine("           , '" + sDealerName + "' ");
            //    strSql.AppendLine("           , '" + sInitialNm + "' ");
            //    strSql.AppendLine("           , '" + sRepName + "' ");
            //    strSql.AppendLine("           , '" + sBizName + "' ");
            //    strSql.AppendLine("           , '" + sTypeName + "' ");
            //    strSql.AppendLine("           , '" + sZipNum + "' ");
            //    strSql.AppendLine("           , '" + sZipSeq + "' ");
            //    strSql.AppendLine("           , '" + sAddr + "' ");
            //    strSql.AppendLine("           , '" + sDtlAddr + "' ");
            //    strSql.AppendLine("           , '" + sRegion + "' ");
            //    strSql.AppendLine("           , '" + sBankCd + "' ");
            //    strSql.AppendLine("           , '" + sBankNM + "' ");
            //    strSql.AppendLine("           , '" + sBankHolder + "' ");
            //    strSql.AppendLine("           , '" + sEMail + "' ");
            //    strSql.AppendLine("           , '" + sHomePG + "' ");
            //    strSql.AppendLine("           , '" + sFax + "' ");
            //    strSql.AppendLine("           , '" + sWebFaxYn + "' ");
            //    strSql.AppendLine("           , '" + sSCMPswd + "' ");
            //    strSql.AppendLine("           , '" + sChrgID + "' ");
            //    strSql.AppendLine("           , '" + sChrgNM + "' ");
            //    strSql.AppendLine("           , '" + sChrgTel + "' ");
            //    strSql.AppendLine("           , '" + sChrgRgnTel + "'  ");
            //    strSql.AppendLine("           , '" + sChrgHP + "' ");
            //    strSql.AppendLine("           , '" + sChrgEmail + "' ");
            //    strSql.AppendLine("           , '" + sPreDealerCd + "' ");
            //    strSql.AppendLine("           , '" + sPreDealerNM + "' ");
            //    strSql.AppendLine("           , '" + sBillKind + "' ");
            //    strSql.AppendLine("           , '" + sBlIssue + "' ");
            //    strSql.AppendLine("           , '" + sPayGB + "' ");
            //    strSql.AppendLine("           , '" + sPayTerm + "' ");
            //    strSql.AppendLine("           , '" + sLimitDcnt + "' ");
            //    strSql.AppendLine("           , '" + sEobYn + "' ");
            //    strSql.AppendLine("           , '" + sNote + "' ");
            //    strSql.AppendLine("           , NOW() ");
            //    strSql.AppendLine("           , '" + sEntID + "' ");
            //    strSql.AppendLine("           , NOW() ");
            //    strSql.AppendLine("           , '" + sMfyId + "' )");
            //    strSql.AppendLine("          ON DUPLICATE KEY UPDATE IDT_NO = '" + sIdtNo + "' ");
            //    strSql.AppendLine("           , STRT_YMD = '" + sStartDay + "' ");
            //    strSql.AppendLine("           , END_YMD = '" + sEndDay + "' ");
            //    strSql.AppendLine("           , DEALER_GB = '" + sDealerGB + "' ");
            //    strSql.AppendLine("           , DEALER_NM = '" + sDealerName + "' ");
            //    strSql.AppendLine("           , INITIAL_NM = '" + sInitialNm + "' ");
            //    strSql.AppendLine("           , REP_NM = '" + sRepName + "' ");
            //    strSql.AppendLine("           , BIZ_NM   = '" + sBizName + "' ");
            //    strSql.AppendLine("           , TYPE_NM   = '" + sTypeName + "' ");
            //    strSql.AppendLine("           , ZIP   = '" + sZipNum + "' ");
            //    strSql.AppendLine("           , ZIP_SEQ = '" + sZipSeq + "' ");
            //    strSql.AppendLine("           , ADDR   = '" + sAddr + "' ");
            //    strSql.AppendLine("           , DTL_ADDR   = '" + sDtlAddr + "' ");
            //    strSql.AppendLine("           , J_REGION = '" + sRegion + "' ");
            //    strSql.AppendLine("           , BANK_CD   = '" + sBankCd + "' ");
            //    strSql.AppendLine("           , BANK_ACNT_NO   = '" + sBankNM + "' ");
            //    strSql.AppendLine("           , ACNT_HOLDER   = '" + sBankHolder + "' ");
            //    strSql.AppendLine("           , EMAIL   = '" + sEMail + "' ");
            //    strSql.AppendLine("           , HOMEPG   = '" + sHomePG + "' ");
            //    strSql.AppendLine("           , FAX = '" + sFax + "' ");
            //    strSql.AppendLine("           , WEB_FAX_YN = '" + sWebFaxYn + "' ");
            //    strSql.AppendLine("           , SCM_PSWD = '" + sSCMPswd + "' ");
            //    strSql.AppendLine("           , CHRG_ID   = '" + sChrgID + "' ");
            //    strSql.AppendLine("           , CHRG_NM   = '" + sChrgNM + "' ");
            //    strSql.AppendLine("           , CHRG_TEL_NO   = '" + sChrgTel + "' ");
            //    strSql.AppendLine("           , CHRG_RGN_NO = '" + sChrgRgnTel + "' ");
            //    strSql.AppendLine("           , CHRG_HP_NO   = '" + sChrgHP + "' ");
            //    strSql.AppendLine("           , CHRG_EMAIL   = '" + sChrgEmail + "' ");
            //    strSql.AppendLine("           , BFR_DEALER_CD   = '" + sPreDealerCd + "' ");
            //    strSql.AppendLine("           , BFR_DEALER_NM   = '" + sPreDealerNM + "' ");
            //    strSql.AppendLine("           , BILL_KIND   = '" + sBillKind + "' ");
            //    strSql.AppendLine("           , BILL_ISSUE_GB   = '" + sBlIssue + "' ");
            //    strSql.AppendLine("           , PAY_GB   = '" + sPayGB + "' ");
            //    strSql.AppendLine("           , PAY_TERM   = '" + sPayTerm + "' ");
            //    strSql.AppendLine("           , SPPL_LIMIT_DCNT   = '" + sLimitDcnt + "' ");
            //    strSql.AppendLine("           , EOB_YN = '" + sEobYn + "' ");
            //    strSql.AppendLine("           , NOTE   = '" + sNote + "' ");
            //    strSql.AppendLine("           , MFY_DT = NOW() ");
            //    strSql.AppendLine("           , MFY_ID   = '" + sMfyId + "' ");

            //    try
            //    {
            //        cmd.CommandType = CommandType.Text;
            //        cmd.CommandText = strSql.ToString();
            //        cmd.ExecuteNonQuery();

            //        string sLogRmk = "Table:ACC_DEALER_CD -> DEALER_CD:" + sDealerCd + ",DEALER_NM:" + sDealerName;
            //        ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "S", this.Name, sLogRmk, cmd);

            //        DBConn.dbTran.Commit();
            //        DBConn.dbTran = null;
            //        XtraMessageBox.Show("저장을 완료했습니다.");
            //        ChkAutoNo.ReadOnly = true;
            //        BtnRetr.PerformClick();
            //        GridViewDealerMng.FocusedRowHandle = GridViewDealerMng.LocateByDisplayText(0, GridColDealerCd, sDealerCd);
            //    }
            //    catch (Exception ex)
            //    {
            //        DBConn.dbTran.Rollback();
            //        DBConn.dbTran = null;
            //        XtraMessageBox.Show(ex.Message);
            //    }

            //    Cursor = Cursors.Default;
            //}

            #endregion[2020-12-29 로그테스트 이전로직]
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            int iIDX = GridViewDealerMng.FocusedRowHandle;

            string sDealerCd = TxtDealerCd.EditValue?.ToString();
            string sDealerNm = TxtDealerNM.EditValue?.ToString();
            string sIdtNo = GridViewDealerMng.GetFocusedRowCellValue(GridColIdtNO)?.ToString();
            if (string.IsNullOrEmpty(sDealerCd))
            {
                XtraMessageBox.Show("올바른 데이터를 선택하세요.");
                return;
            }

            string sMsg = string.Format("거래처번호 : {0}\r\n거래처명 : {1}\r\n해당 거래처를 삭제하시겠습니까?", sDealerCd, sDealerNm);
            if (XtraMessageBox.Show(sMsg, "거래처 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }
            
            try
            {
                Cursor = Cursors.WaitCursor;
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM ACC_DEALER_CD ");
                strSql.AppendLine("       WHERE DEALER_CD = @DEALER_CD ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@DEALER_CD", sDealerCd);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                //LOG
                string sCurTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string sLogRmk = "[거래처삭제] 거래처코드 :" + sDealerCd + ", 거래처명 :" + sDealerNm + ", 사업자번호 : " + sIdtNo;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("CUR_TIME", sCurTime);
                dicParams.Add("USRCD", FmMainToolBar2.UserID);
                dicParams.Add("PGMID", this.Name);
                dicParams.Add("IP", ComnEtcFunc.GetLocalIP());
                dicParams.Add("EDIT_RMK", sLogRmk);

                strSql.Clear();
                #region mariaDB
                //strSql.AppendLine("          INSERT INTO ZSYS_LOG ");
                //strSql.AppendLine("          	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                //strSql.AppendLine("          	      VALUES ");
                //strSql.AppendLine("          	           ( @CUR_TIME ");
                //strSql.AppendLine("          	           , @USRCD ");
                //strSql.AppendLine("          	           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                //strSql.AppendLine("                               FROM ZSYS_LOG X1 ");
                //strSql.AppendLine("                              WHERE X1.OCCUR_DATE = @CUR_TIME ");
                //strSql.AppendLine("                                AND X1.USRCD = @USRCD ) #LOG_SEQ(구분자) ");
                //strSql.AppendLine("          	           , @PGMID ");
                //strSql.AppendLine("          	           , 'D' ");
                //strSql.AppendLine("          	           , @IP ");
                //strSql.AppendLine("          	           , @EDIT_RMK ); ");
                #endregion

                strSql.AppendLine("          INSERT INTO ZSYS_LOG ");
                strSql.AppendLine("          	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                strSql.AppendLine("          	      VALUES ");
                strSql.AppendLine("          	           ( @CUR_TIME ");
                strSql.AppendLine("          	           , @USRCD ");
                strSql.AppendLine("          	           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                strSql.AppendLine("                               FROM ZSYS_LOG X1 ");
                strSql.AppendLine("                              WHERE X1.OCCUR_DATE = @CUR_TIME ");
                strSql.AppendLine("                                AND X1.USRCD = @USRCD ) --LOG_SEQ(구분자) ");
                strSql.AppendLine("          	           , @PGMID ");
                strSql.AppendLine("          	           , 'D' ");
                strSql.AppendLine("          	           , @IP ");
                strSql.AppendLine("          	           , @EDIT_RMK ); ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                foreach(KeyValuePair<string, string> param in dicParams)
                {
                    cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                }
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                
                //2020-12-18
                //History도 수정진행
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM ACC_DEALER_CD ");
                strSql.AppendLine("       WHERE DEALER_CD = @DEALER_CD ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.AddWithValue("@DEALER_CD", sDealerCd);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                
                //계좌삭제
                strSql.Clear();
                strSql.AppendLine(" DELETE FROM ACC_ACNT_CD");
                strSql.AppendLine("  WHERE ACNT_CD = @DEALER_CD     ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.AddWithValue("@DEALER_CD", sDealerCd);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                Cursor = Cursors.Default;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewDealerMng.FocusedRowHandle = iIDX-1;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SetLookUpEdit(DevExpress.XtraEditors.LookUpEdit lkup, string sGb, string sNullYn, string sSetIdx)
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
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM  ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'BANK_CD' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'BILL_KIND'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'BILL_ISSUE_GB'");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'PAY_GB'");
            }
            else if (sGb.Equals("5"))
            {
                //strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                //strSql.AppendLine("      , A.COM_NM AS NM ");
                //strSql.AppendLine("   FROM COM_BASE_CD A ");
                //strSql.AppendLine("  WHERE A.CD_GB = 'CHRG_ID'");
                strSql.AppendLine(" SELECT A.EMP_ID AS CD ");
                strSql.AppendLine("      , A.EMP_NM AS NM ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A ");

            }

            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            lkup.Properties.DataSource = dt;
            lkup.Properties.DisplayMember = "NM";
            lkup.Properties.ValueMember = "CD";

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }

        private void GridViewDealerMng_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            CultureInfo enUS = new CultureInfo("en-US");

            ChkAutoNo.ReadOnly = true;
            ChkAutoNo.Properties.ValueChecked = true;
            ChkAutoNo.Checked = true;

            TxtDealerCd.EditValue = GridViewDealerMng.GetFocusedRowCellValue("DEALER_CD") ?? string.Empty;
            //TxtIdtNo.ReadOnly = true;
            TxtIdtNo.EditValue = GridViewDealerMng.GetFocusedRowCellValue("IDT_NO") ?? string.Empty;
            TxtCorpNo.EditValue = GridViewDealerMng.GetFocusedRowCellValue(GridColCorpNo)?.ToString();
            TxtDealerNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("DEALER_NM") ?? string.Empty;
            TxtInitialNm.EditValue = GridViewDealerMng.GetFocusedRowCellValue("INITIAL_NM") ?? string.Empty;
            CbxDealerGB.EditValue = GridViewDealerMng.GetFocusedRowCellValue("DEALER_GB") ?? string.Empty;
            TxtBizNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BIZ_NM") ?? string.Empty;
            TxtRepNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("REP_NM") ?? string.Empty;
            TxtTypeNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("TYPE_NM") ?? string.Empty;
            TxtZipNum.EditValue = GridViewDealerMng.GetFocusedRowCellValue("ZIP") ?? string.Empty;
            TxtAddr.EditValue = GridViewDealerMng.GetFocusedRowCellValue("ADDR") ?? string.Empty;
            TxtDtlAddr.EditValue = GridViewDealerMng.GetFocusedRowCellValue("DTL_ADDR") ?? string.Empty;
            TxtRegion.EditValue = GridViewDealerMng.GetFocusedRowCellValue("J_REGION") ?? string.Empty;
            TxtEmail.EditValue = GridViewDealerMng.GetFocusedRowCellValue("EMAIL") ?? string.Empty;
            TxtHomePG.EditValue = GridViewDealerMng.GetFocusedRowCellValue("HOMEPG") ?? string.Empty;
            TxtFaxNum.EditValue = GridViewDealerMng.GetFocusedRowCellValue("FAX") ?? string.Empty;
            RdgbWebFaxGb.EditValue = GridViewDealerMng.GetFocusedRowCellValue("WEB_FAX_YN") ?? string.Empty;
            TxtScmPw.EditValue = GridViewDealerMng.GetFocusedRowCellValue("SCM_PSWD") ?? string.Empty;

            LkupBankCd.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BANK_CD")?.ToString().Trim();
            //if(string.IsNullOrEmpty(sBankCdChk))
            //{
            //    LkupBankNM.EditValue = "****";
            //}
            //else
            //{
            //    LkupBankNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BANK_CD") ?? string.Empty;
            //}

            TxtHolder.EditValue = GridViewDealerMng.GetFocusedRowCellValue("ACNT_HOLDER") ?? string.Empty;
            TxtBankNum.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BANK_ACNT_NO") ?? string.Empty;
            LkupEditChrgID.EditValue = GridViewDealerMng.GetFocusedRowCellValue("CHRG_ID") ?? string.Empty;
            TxtChrgNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("CHRG_NM") ?? string.Empty;
            TxtChrgHP.EditValue = GridViewDealerMng.GetFocusedRowCellValue("CHRG_HP_NO") ?? string.Empty;
            TxtChrgEmail.EditValue = GridViewDealerMng.GetFocusedRowCellValue("CHRG_EMAIL") ?? string.Empty;
            TxtChrgTel.EditValue = GridViewDealerMng.GetFocusedRowCellValue("CHRG_TEL_NO") ?? string.Empty;
            TxtChrgRgnNo.EditValue = GridViewDealerMng.GetFocusedRowCellValue(GridColChrgRgnNo)?.ToString();
            LkupEditBillKind.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BILL_KIND") ?? string.Empty;
            LkupBIssueGB.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BILL_ISSUE_GB") ?? string.Empty;
            TxtPayTerm.EditValue = GridViewDealerMng.GetFocusedRowCellValue("PAY_TERM") ?? string.Empty;
            LkupEditPayGB.EditValue = GridViewDealerMng.GetFocusedRowCellValue("PAY_GB") ?? string.Empty;
            TxtLtDcnt.EditValue = GridViewDealerMng.GetFocusedRowCellValue("SPPL_LIMIT_DCNT")?? string.Empty;
            RdgbEobYn.EditValue = GridViewDealerMng.GetFocusedRowCellValue("EOB_YN") ?? string.Empty;
            TxtNote.EditValue = GridViewDealerMng.GetFocusedRowCellValue("NOTE") ?? string.Empty;
            ChkBANKYN.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BANKYN") ?? string.Empty;

            buttonEdit1.EditValue = GridViewDealerMng.GetFocusedRowCellValue("IMTXT1") ?? string.Empty;
            buttonEdit2.EditValue = GridViewDealerMng.GetFocusedRowCellValue("IMTXT2") ?? string.Empty;
            buttonEdit3.EditValue = GridViewDealerMng.GetFocusedRowCellValue("IMTXT3") ?? string.Empty;

            string sYmdFromChk = GridViewDealerMng.GetFocusedRowCellValue("STRT_YMD")?.ToString();
            if (!string.IsNullOrEmpty(sYmdFromChk))
            {
                //DateEditFrom.EditValue = DateTime.ParseExact(GridViewDealerMng.GetFocusedRowCellValue("STRT_YMD").ToString(), "yyyyMMdd", enUS);
                DateEditFrom.EditValue = GridViewDealerMng.GetFocusedRowCellValue("STRT_YMD").ToString();
            }
            else
            {
                DateEditFrom.EditValue = string.Empty;
            }
            
            string sYmdEndChk = GridViewDealerMng.GetFocusedRowCellValue("END_YMD")?.ToString();
            if (!string.IsNullOrEmpty(sYmdEndChk))
            {
                //DateEditTo.EditValue = DateTime.ParseExact(GridViewDealerMng.GetFocusedRowCellValue("END_YMD").ToString(), "yyyyMMdd", enUS);
                DateEditTo.EditValue = GridViewDealerMng.GetFocusedRowCellValue("END_YMD").ToString();
            }
            else
            {
                DateEditTo.EditValue = string.Empty;
            }

            TxtBfrDlrCd.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BFR_DEALER_CD") ?? string.Empty;
            TxtBfrDlrNM.EditValue = GridViewDealerMng.GetFocusedRowCellValue("BFR_DEALER_NM") ?? string.Empty;

            string sDealerCd = GridViewDealerMng.GetFocusedRowCellValue("DEALER_CD")?.ToString();

            for (int i = 0; i < GridViewHistoryRetr.RowCount;)
            {
                GridViewHistoryRetr.DeleteRow(i);
            }

            GetDealerHistory(sDealerCd);
        }

        private void SetClear()
        {
            ChkAutoNo.ReadOnly = true;
            ChkAutoNo.Properties.ValueChecked = true;
            ChkAutoNo.Checked = true;

            TxtDealerCd.EditValue = null;
            //TxtIdtNo.ReadOnly = true;
            TxtIdtNo.EditValue = null;
            TxtDealerNM.EditValue = null;
            TxtInitialNm.EditValue = null;
            CbxDealerGB.EditValue = null;
            TxtBizNM.EditValue = null;
            TxtRepNM.EditValue = null;
            TxtTypeNM.EditValue =null;
            TxtZipNum.EditValue = null;
            TxtAddr.EditValue = null;
            TxtDtlAddr.EditValue =null;
            TxtRegion.EditValue = null;
            TxtEmail.EditValue = null;
            TxtHomePG.EditValue = null;
            TxtFaxNum.EditValue = null;
            RdgbWebFaxGb.EditValue = null;
            TxtScmPw.EditValue = null;

            string sBankCdChk = null;
            if (string.IsNullOrEmpty(sBankCdChk))
            {
                LkupBankCd.EditValue = null;
            }
            else
            {
                LkupBankCd.EditValue = null;
            }

            TxtHolder.EditValue = null;
            TxtBankNum.EditValue = null;
            LkupEditChrgID.EditValue = null;
            TxtChrgNM.EditValue = null;
            TxtChrgHP.EditValue = null;
            TxtChrgEmail.EditValue = null;
            TxtChrgTel.EditValue = null;
            TxtChrgRgnNo.EditValue = null;
            LkupEditBillKind.EditValue = null;
            LkupBIssueGB.EditValue = null;
            TxtPayTerm.EditValue = null;
            LkupEditPayGB.EditValue = null;
            TxtLtDcnt.EditValue = null;
            RdgbEobYn.EditValue = null;
            TxtNote.EditValue = null;

            string sYmdFromChk = null;
            if (!string.IsNullOrEmpty(sYmdFromChk))
            {
                //DateEditFrom.EditValue = DateTime.ParseExact(GridViewDealerMng.GetFocusedRowCellValue("STRT_YMD").ToString(), "yyyyMMdd", enUS);
                DateEditFrom.EditValue = GridViewDealerMng.GetFocusedRowCellValue("STRT_YMD").ToString();
            }
            else
            {
                DateEditFrom.EditValue = string.Empty;
            }

            string sYmdEndChk = null;
            if (!string.IsNullOrEmpty(sYmdEndChk))
            {
                //DateEditTo.EditValue = DateTime.ParseExact(GridViewDealerMng.GetFocusedRowCellValue("END_YMD").ToString(), "yyyyMMdd", enUS);
                DateEditTo.EditValue = GridViewDealerMng.GetFocusedRowCellValue("END_YMD").ToString();
            }
            else
            {
                DateEditTo.EditValue = string.Empty;
            }

            TxtBfrDlrCd.EditValue = null;
            TxtBfrDlrNM.EditValue = null;

            string sDealerCd = null;

            for (int i = 0; i < GridViewHistoryRetr.RowCount;)
            {
                GridViewHistoryRetr.DeleteRow(i);
            }

            GetDealerHistory(sDealerCd);
        }

        private void GetDealerHistory(string sDealerCd)
        {
            if (!string.IsNullOrEmpty(sDealerCd))
            {
                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT A.DEALER_CD ");
                strSql.AppendLine("      , A.HIS_SEQ");
                strSql.AppendLine("      , A.REMARK ");
                strSql.AppendLine("      , A.USE_YN ");
                strSql.AppendLine("      , A.ENT_DT ");
                strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
                strSql.AppendLine("      , A.MFY_DT ");
                strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
                strSql.AppendLine("   FROM ACC_DEALER_HISTORY A");
                strSql.AppendLine("  WHERE A.DEALER_CD = '" + sDealerCd + "' ");
                strSql.AppendLine("    AND A.USE_YN = 'Y' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                GridHistoryRetr.DataSource = dt;
            }
        }

        private void BtnHistoryNew_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sDealerCd = GridViewDealerMng.GetFocusedRowCellValue("DEALER_CD")?.ToString();

            if (sDealerCd != null)
            {
                GridViewHistoryRetr.AddNewRow();
                GridViewHistoryRetr.ShowEditForm();
                GridViewHistoryRetr.SetFocusedRowCellValue("DEALER_CD", sDealerCd);
            }
            else
            {
                XtraMessageBox.Show("추가하고자하는 거래처를 정확히 선택해주세요.");
            }
        }
        
        private void BtnHisSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            DataTable dt = (DataTable)GridHistoryRetr.DataSource;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            double dDealerCd = 0;
            string sStrtYmd = string.Empty;
            string sEndYmd = string.Empty;
            double dSeq = 0;
            string sRemark = string.Empty;
            string sEntId = string.Empty;
            string sMfyId = string.Empty;

            try
            {
                if (dtMerge.Rows.Count == 0)
                {
                    XtraMessageBox.Show("추가 및 입력사항이 없습니다.");
                    return;
                }
                else if (dtMerge.Rows.Count > 0)
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        dDealerCd = Convert.ToDouble(dtMerge.Rows[j]["DEALER_CD"]);
                        sRemark = Convert.ToString(dtMerge.Rows[j]["REMARK"]);
                        sEntId = FmMainToolBar2.UserID;
                        sMfyId = FmMainToolBar2.UserID;

                        StringBuilder strSql = new StringBuilder();

                        DataTable dtResult = (DataTable)GridHistoryRetr.DataSource;

                        bool bChk = String.IsNullOrEmpty(dtMerge.Rows[j]["HIS_SEQ"].ToString()) ? true : false;

                        if (bChk)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" SELECT MAX(A.HIS_SEQ) AS MAX_SEQ ");
                            strSql.AppendLine("   FROM ACC_DEALER_HISTORY A");
                            strSql.AppendLine("  WHERE A.DEALER_CD = " + dDealerCd + " ");

                            dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                            DataRow[] row = dtResult.Select();

                            if (row[0]["MAX_SEQ"].ToString() == string.Empty)
                            {
                                dSeq = 1;
                            }
                            else
                            {
                                dSeq = Convert.ToDouble(row[0]["MAX_SEQ"].ToString()) + 1;
                            }
                            dtResult = null;
                            row = null;
                        }
                        else
                        {
                            dSeq = Convert.ToDouble(dtMerge.Rows[j]["HIS_SEQ"].ToString());
                            dtResult = null;
                        }

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO ACC_DEALER_HISTORY ");
                        //strSql.AppendLine("           ( DEALER_CD ");
                        //strSql.AppendLine("           , HIS_SEQ ");
                        //strSql.AppendLine("           , REMARK ");
                        //strSql.AppendLine("           , USE_YN ");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( " + dDealerCd + " ");
                        //strSql.AppendLine("           , " + dSeq + " ");
                        //strSql.AppendLine("           , '" + sRemark + "' ");
                        //strSql.AppendLine("           , 'Y'");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sMfyId + "' ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("           ON DUPLICATE KEY UPDATE");
                        //strSql.AppendLine("             REMARK = '" + sRemark + "' ");
                        //strSql.AppendLine("           , USE_YN = 'Y' ");
                        //strSql.AppendLine("           , MFY_DT = NOW() ");
                        //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        #endregion
                        strSql.AppendLine("SET IDENTITY_INSERT ACC_DEALER_HISTORY ON");
                        strSql.AppendLine("IF EXISTS(SELECT* FROM ACC_DEALER_HISTORY WHERE DEALER_CD = "+ dDealerCd + " AND HIS_SEQ = "+ dSeq + ")");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         UPDATE ACC_DEALER_HISTORY");
                        strSql.AppendLine("            SET REMARK = '" + sRemark + "' ");
                        strSql.AppendLine("           , USE_YN = 'Y' ");
                        strSql.AppendLine("           , MFY_DT = CONVERT(VARCHAR(19),GETDATE(),20) ");
                        strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        strSql.AppendLine("          WHERE DEALER_CD = " + dDealerCd + " AND HIS_SEQ = " + dSeq);
                        strSql.AppendLine("   END");
                        strSql.AppendLine("ELSE");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("          INSERT INTO ACC_DEALER_HISTORY");
                        strSql.AppendLine("               ( DEALER_CD ");
                        strSql.AppendLine("               , HIS_SEQ ");
                        strSql.AppendLine("               , REMARK ");
                        strSql.AppendLine("               , USE_YN ");
                        strSql.AppendLine("               , ENT_DT ");
                        strSql.AppendLine("               , ENT_ID ");
                        strSql.AppendLine("               , MFY_DT ");
                        strSql.AppendLine("               , MFY_ID ");
                        strSql.AppendLine("               ) ");
                        strSql.AppendLine("          VALUES ");
                        strSql.AppendLine("               ( " + dDealerCd + " ");
                        strSql.AppendLine("               , " + dSeq + " ");
                        strSql.AppendLine("               , '" + sRemark + "' ");
                        strSql.AppendLine("               , 'Y'");
                        strSql.AppendLine("               , CONVERT(VARCHAR(19),GETDATE(),20) ");
                        strSql.AppendLine("               , '" + sEntId + "' ");
                        strSql.AppendLine("               , CONVERT(VARCHAR(19),GETDATE(),20) ");
                        strSql.AppendLine("               , '" + sMfyId + "' ");
                        strSql.AppendLine("               ) ");
                        strSql.AppendLine("   END");
                        strSql.AppendLine("SET IDENTITY_INSERT ACC_DEALER_HISTORY OFF");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "Table:ACC_DEALER_HISTORY -> DEALER_CD:" + dDealerCd + ",HIS_SEQ:" + dSeq;
                        //ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, (j+1).ToString(), "S", this.Name, sLogRmk, cmd);
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show("저장을 완료했습니다.");

                    GetDealerHistory(dDealerCd.ToString());
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }

            Cursor = Cursors.Default;
        }
        
        private void BtnHisDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (GridViewDealerMng.IsGroupRow(GridViewDealerMng.FocusedRowHandle))
            {
                XtraMessageBox.Show("삭제할 내용이 없습니다. 정확한 거래처항목을 눌러주세요.");
                return;
            }

            DataTable dtChk = (DataTable)GridHistoryRetr.DataSource;

            if (dtChk.Rows.Count == 0)
            {
                XtraMessageBox.Show("삭제할 데이터가 없습니다.");
                return;
            }
            else if (GridViewHistoryRetr.GetFocusedRowCellValue("HIS_SEQ").ToString().Equals(string.Empty))
            {
                GridViewHistoryRetr.DeleteRow(GridViewHistoryRetr.FocusedRowHandle);
                return;
            }

            double dDealerCd = Convert.ToDouble(GridViewHistoryRetr.GetFocusedRowCellValue("DEALER_CD").ToString());
            double dHisSeq = Convert.ToDouble(GridViewHistoryRetr.GetFocusedRowCellValue("HIS_SEQ").ToString());

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" UPDATE ACC_DEALER_HISTORY ");
                strSql.AppendLine("    SET USE_YN = 'N' ");
                strSql.AppendLine("  WHERE DEALER_CD = '" + dDealerCd + "' ");
                strSql.AppendLine("    AND HIS_SEQ = '" + dHisSeq + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:ACC_DEALER_HISTORY -> DEALER_CD:" + dDealerCd + ",HIS_SEQ:" + dHisSeq;
                //ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                GetDealerHistory(dDealerCd.ToString());
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void RdgbMenuEobYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
            int rdgbIdx = RdgbMenuEobYn.SelectedIndex;
            DataTable dt = GetLookUpData("1");
            ComGrid.SetLookUpEdit(LkupDealerRetr, dt, "CD", "CD", "Y");
        }

        private bool FindDealerCd(string sDealerCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
            strSql.AppendLine("   FROM ACC_DEALER_CD  ");
            strSql.AppendLine("  WHERE DEALER_CD = " + sDealerCd + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);
            if (dCnt > 0)
                return true;
            else
                return false;
        }

    private DataTable GetLookUpData(string sGb)
    {
        StringBuilder strSql = new StringBuilder();

        strSql.AppendLine(" ");
        strSql.AppendLine("WITH ITEM_INFO AS (");
        strSql.AppendLine(" SELECT NULL AS CD");
        strSql.AppendLine("     , '' AS NM");
        strSql.AppendLine(" UNION ALL");

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT DEALER_CD AS CD     ");
                strSql.AppendLine(" 	 , DEALER_NM AS NM     ");
                strSql.AppendLine(" FROM ACC_DEALER_CD");
                strSql.AppendLine(" WHERE DEALER_NM != ''");
                if (!RdgbMenuEobYn.EditValue.Equals("ALL")) strSql.AppendLine(" AND EOB_YN = '" + RdgbMenuEobYn.EditValue + "'");
                //strSql.AppendLine(" GROUP BY DEALER_NM         ");
                //strSql.AppendLine(" SELECT DEALER_NM AS CD     ");
                //strSql.AppendLine(" 	 , '' AS NM     ");
                //strSql.AppendLine(" FROM ACC_DEALER_CD");
                //strSql.AppendLine(" WHERE DEALER_NM != ''");
                //if (!RdgbMenuEobYn.EditValue.Equals("ALL")) strSql.AppendLine(" AND EOB_YN = '" + RdgbMenuEobYn.EditValue + "'");
                //strSql.AppendLine(" GROUP BY DEALER_NM         ");
            }

        strSql.AppendLine(") ");
        strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            

        DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        return dt;
    }

    private void ClearFps()
        {
            GridViewDealerMng.FocusedRowChanged -= GridViewDealerMng_FocusedRowChanged;
            GridDealerMng.DataSource = null;
            GridViewDealerMng.FocusedRowChanged += GridViewDealerMng_FocusedRowChanged;
        }

        public void ClearAllForm(Control Ctrl)
        {
            string sVal = CboFindSbj.EditValue?.ToString();
            if (Ctrl.HasChildren)
            {
                foreach (Control ctrl in Ctrl.Controls)
                {
                    if (ctrl is DevExpress.XtraEditors.TextEdit)
                    {
                        if ((((DevExpress.XtraEditors.BaseEdit)ctrl).Name.Equals(LkupDealerRetr.ToString())))
                        {
                            return;
                        }
                        else
                        {
                            (ctrl as DevExpress.XtraEditors.TextEdit).ResetText();
                        }
                    }

                    if (ctrl is DevExpress.XtraEditors.LookUpEdit)
                    {
                        if (ctrl == LkupDealerRetr)
                            continue;

                        (ctrl as DevExpress.XtraEditors.LookUpEdit).EditValue = "****";
                    }

                    if (ctrl is DevExpress.XtraEditors.DateEdit)
                        (ctrl as DevExpress.XtraEditors.DateEdit).EditValue = DateTime.Now.ToString("yyyy-MM-dd");

                    if (ctrl is DevExpress.XtraEditors.ComboBoxEdit)
                    {
                        if (ctrl == CboFindSbj)
                            continue;

                        (ctrl as DevExpress.XtraEditors.ComboBoxEdit).ResetText();
                    }
                    if (ctrl is DevExpress.XtraEditors.RadioGroup)
                        (ctrl as DevExpress.XtraEditors.RadioGroup).SelectedIndex = 1;

                    if (ctrl.HasChildren)
                        ClearAllForm(ctrl);//Recursive
                }
            }

            CboFindSbj.EditValue = sVal;
        }

        private bool CheckSameChar(char[] sArrIdtNo)
        {
            bool bYn = true;
            if (sArrIdtNo.Length == 10)
            {
                //사업자번호가 동일한 숫자로 되어있는 경우 사업자번호를 등록하지 않는 것으로 판단
                char cPrvstr = '1';
                
                for (int i = 0; i < sArrIdtNo.Length; i++)
                {
                    if (i == 0)
                    {
                        cPrvstr = sArrIdtNo[i];
                    }
                    else
                    {
                        if (!cPrvstr.Equals(sArrIdtNo[i]))
                        {
                            bYn = false;
                            break;
                        }
                    }
                }
            }

            return bYn;
        }

        private void LkupDealerRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void GridViewDealerMng_Layout(object sender, EventArgs e)
        {
            //GridViewDealerMng.SaveLayoutToXml(@"C:\Users\USER2\source\repos\SteelAccSln\AccAdm\xaml\AccDealerCdDev.xaml");
        }

        private void GridViewDealerMng_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("IDT_NO"))
            {
                if(e.DisplayText.Length == 10)
                {
                    string sTemp = e.DisplayText;
                    string sResult = sTemp.Substring(0, 3) + "-" + sTemp.Substring(3, 2) + "-" + sTemp.Substring(5, 5);
                    e.DisplayText = sResult;
                }
            }
        }

        private void BtnReport_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 리포트 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            DataTable dt = (DataTable)GridDealerMng.DataSource;
            ReportViewer fm = new ReportViewer(dt, "DealerCdReport");
            fm.ShowDialog();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        
        private void AccDealerCdDev_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
            {
                
            }
            else if(e.KeyCode == Keys.F1)
            {
                BtnCrete_Click(null, null);
            }
            else if(e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if(e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F7)
            {
                BtnReport_Click(null, null);
            }
            else if(e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
            else if (e.KeyCode == Keys.F9)
            {
                BtnHistoryNew.PerformClick();
            }
            else if (e.KeyCode == Keys.F11)
            {
                BtnHisSave.PerformClick();
            }
            else if (e.KeyCode == Keys.F12)
            {
                BtnHisDelete.PerformClick();
            }
        }

        private void GridDealerMng_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewDealerMng_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccDealerCdDev_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void BtnExcel_Click(object sender, EventArgs e)
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
                string sYmd = DateTime.Today.ToString().Substring(0, 10);

                string sFileNM = "거래처관리 " + sYmd;
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    GridDealerMng.ExportToXls(FileName);
                    Process.Start(FileName);
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

        private void GridViewHistoryRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #region [정렬기능(2020-06-02 정은영)]
        //multisort
        private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        {
            /*
             * 2021-01-20
             * 그리드 헤더 사이징이 되지 않아 아래 코드 주석처리
             */
            //GridView view = (GridView)sender;
            //GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            //if (hitInfo.InColumn)
            //{
            //    if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.None)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
            //        view.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Descending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            //        view.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Ascending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.None;
            //        view.FocusedRowHandle = 0;
            //    }
            //    // if ((ModifierKeys & Keys.Control) == Keys.Control) return;
            //    //if ((ModifierKeys & Keys.Shift) != Keys.Shift) view.ClearSorting();
            //}
        }
        #endregion

        private void AccDealerCdDev_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                ////string path = @"C:\STLNT\" + sProject + @"\xaml\" + sId;
                string path = ComnEtcFunc.GetLayoutPath();
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                ////string sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sId + @"\" + sClass + ".xaml";
                //string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
                //view.SaveLayoutToXml(sFile);

                //layoutControl1.SaveLayoutToXml(path);
                //writer.Close();

                layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");
                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void ChkAutoNo_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void ChkAutoNo_CheckStateChanged(object sender, EventArgs e)
        {
            if (!ChkAutoNo.Checked)
            {
                XtraMessageBox.Show("수동으로 거래처코드를 입력할 시 존재하지 않는 코드로 입력하여야합니다.");
                TxtDealerCd.EditValue = string.Empty;
                TxtDealerCd.ReadOnly = false;
                TxtDealerCd.Focus();
            }
            else
            {
                TxtDealerCd.ReadOnly = true;
                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT MAX(A.DEALER_CD) AS MAX_CD ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                TxtDealerCd.EditValue = dt.Rows[0]["MAX_CD"];
                TxtDealerNM.Focus();
            }
        }

        private void TxtDealerNM_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //    TxtIdtNo.Focus();
        }

        private void buttonEdit1_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            PP008F00 frm = new PP008F00();

            frm.Owner = this;
            frm.DataRowSendEvent += new PP008F00.SendDataHandler(GetAddress);
            frm.ShowDialog();
        }

        public void GetAddress(DataRow row)
        {
            TxtZipNum.EditValue = row["zipNo"];
            TxtAddr.EditValue = row["roadAddr"];

            string siNm = row["siNm"]?.ToString();

            if (siNm.Contains("특별") || siNm.Contains("광역"))
            {
                TxtRegion.EditValue = siNm.Substring(0, 2);
            }
            else if (siNm.Equals("제주도") || siNm.Equals("제주특별자치도"))
            {
                TxtRegion.EditValue = siNm;
            }
            else
            {
                TxtRegion.EditValue = row["sggNm"]?.ToString();
            }
        }

        private void image1_ButtonClick(object sender, ButtonPressedEventArgs e)
        {

        }

        private void buttonEdit1_ButtonClick_1(object sender, ButtonPressedEventArgs e)
        {
            string dealercdch = TxtDealerCd.EditValue?.ToString();
            Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST3");
            dicParams.Add("DEALER_CD", dealercdch);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            
            if (dt.Rows.Count ==0)
            {
                XtraMessageBox.Show("신규등록 후 이미지파일을 등록하세요.", "");
                return;
            }
            
            AccDealerCdDevimage frm = new AccDealerCdDevimage();

            frm.Owner = this;
            //frm._ModeGubun = AccDealerCdDevimage.ModeGubun.Add;
            frm._dealercd = TxtDealerCd.EditValue?.ToString();
            frm._check = "1";
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }

        private void buttonEdit2_ButtonClick(object sender, ButtonPressedEventArgs e)
        {

            string dealercdch = TxtDealerCd.EditValue?.ToString();
            Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST3");
            dicParams.Add("DEALER_CD", dealercdch);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("신규등록 후 이미지파일을 등록하세요.", "");
                return;
            }

            AccDealerCdDevimage frm = new AccDealerCdDevimage();

            frm.Owner = this;
            //frm._ModeGubun = AccDealerCdDevimage.ModeGubun.Add;
            frm._dealercd = TxtDealerCd.EditValue?.ToString();
            frm._check = "2";
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }

        private void buttonEdit3_ButtonClick(object sender, ButtonPressedEventArgs e)
        {

            string dealercdch = TxtDealerCd.EditValue?.ToString();
            Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST3");
            dicParams.Add("DEALER_CD", dealercdch);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("신규등록 후 이미지파일을 등록하세요.", "");
                return;
            }

            AccDealerCdDevimage frm = new AccDealerCdDevimage();

            frm.Owner = this;
            //frm._ModeGubun = AccDealerCdDevimage.ModeGubun.Add;
            frm._dealercd = TxtDealerCd.EditValue?.ToString();
            frm._check = "3";
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }
    }
}
