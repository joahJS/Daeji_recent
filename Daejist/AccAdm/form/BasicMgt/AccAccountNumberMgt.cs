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
using DevExpress.XtraEditors.Repository;
using ComLib;
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.EditForm.Helpers.Controls;
using System.Diagnostics;
using System.Data.SqlClient;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
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
    public partial class AccAccountNumberMgt : DevExpress.XtraEditors.XtraForm
    {
        public AccAccountNumberMgt()
        {
            InitializeComponent();
        }

        private void AccAccountNumberMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            LkupEditAccCd.EditValueChanged -= LkupEditAccCd_EditValueChanged;
            LkupEditBankNm.EditValueChanged -= LkupEditBankNm_EditValueChanged;
            RdGrRetrCalcel.EditValueChanged -= RdGrRetrCalcel_EditValueChanged;
            DataTable dtAccCd = GetLookUpData("1", "1", "", "Y");
            DataTable dtBankNm = GetLookUpData("2", "1", "", "Y");
            DataTable dtMethod = GetLookUpData("3", "", "", "Y");
            DataTable dtOtherGea = GetLookUpData("4", "", "", "");
            DataTable dtGejagb = GetLookUpData("5", "", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditAccCd, dtAccCd, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditBankNm, dtBankNm, "CD", "NM", "Y");
            LkupEditAccCd.EditValue = "****";
            LkupEditBankNm.EditValue = "****";
            RdGrRetrCalcel.EditValue = "*";
            RepositoryItemGridLookUpEdit accLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(accLkup, dtAccCd, GridRetr, GridColAccCd, "CD", "NM", "");
            RepositoryItemGridLookUpEdit bankLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(bankLkup, dtBankNm, GridRetr, GridColBankNm, "CD", "NM", "");
            RepositoryItemGridLookUpEdit methodkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(methodkup, dtMethod, GridRetr, GridColMethod, "CD", "NM", "");
            RepositoryItemGridLookUpEdit otherGeaLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(otherGeaLkup, dtOtherGea, GridRetr, GridColOtherGea, "CD", "NM", "");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGridLkupGejagb, dtGejagb, GridRetr, GridColGejagb, "CD", "NM", "");

            //
            //rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccAccountNumberMgt");
            //ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccAccountNumberMgt", GridViewRetr);
            SetLoadFormLayout();

            BtnRetr_Click(null, null);
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr };
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

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
        }
        #endregion

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine(" UNION ALL");
            }


            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'ACNT_ACC_CD'");

            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'BANK_CD'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'PAYIN_METHOD'");
            }
            else if(sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT A.ACCOD AS CD");
                strSql.AppendLine("      , A.ACNAM AS NM");
                strSql.AppendLine("   FROM ACMSTF A");
            }
            else if (sGb.Equals("5"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'GEJAGB'");
            }

            if (sOther.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }
            else
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            }


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetGridRetr();
            LkupEditBankNm.EditValueChanged += LkupEditBankNm_EditValueChanged;
            LkupEditAccCd.EditValueChanged += LkupEditAccCd_EditValueChanged;
            RdGrRetrCalcel.EditValueChanged += RdGrRetrCalcel_EditValueChanged;

        }
        private void GetGridRetr()
        {
            Cursor = Cursors.WaitCursor;
            GridRetr.DataSource = null;
            StringBuilder strSql = new StringBuilder();
            string sAccCd = LkupEditAccCd.EditValue.ToString();
            string sBankCd = LkupEditBankNm.EditValue.ToString();
            string sCancel = RdGrRetrCalcel.EditValue.ToString();

            strSql.AppendLine("");
            strSql.AppendLine(" SELECT A1.DEALER_CD                                ");
            strSql.AppendLine("      , A1.DEALER_NM                                ");
            strSql.AppendLine("      , A1.BANK_CD                                  ");
            strSql.AppendLine("      , A1.BANK_ACNT_NO                             ");
            strSql.AppendLine("      , A1.EOB_YN                                   ");
            strSql.AppendLine("      , A2.PSSEQ--표시순번                          ");
            strSql.AppendLine("      , A2.GEJAGB--구분                             ");
            strSql.AppendLine("      , A2.ACC_CD--계정코드                         ");
            strSql.AppendLine("      , A2.RPLC_ACC_CD--대체계정코드                ");
            strSql.AppendLine("      , A2.BRANCH_NM--지점명                        ");
            strSql.AppendLine("      , A2.FINANCE_GOODS_NM--금융상품명             ");
            strSql.AppendLine("      , A2.CNTR_YMD--계약일자                       ");
            strSql.AppendLine("      , A2.CNTR_AMT--계약금액                       ");
            strSql.AppendLine("      , A2.INRST_RATE--이자율                       ");
            strSql.AppendLine("      , A2.FRST_PAYIN_YMD--최초납입일자             ");
            strSql.AppendLine("      , A2.PAYIN_AMT--납입금액                      ");
            strSql.AppendLine("      , A2.PAYIN_DD--납입일                         ");
            strSql.AppendLine("      , A2.PAYIN_METHOD--납입방법                   ");
            strSql.AppendLine("      , A2.EXPIRE_YMD--만기일자                     ");
            strSql.AppendLine("      , A2.EXPIRE_INRST_AMT--만기이자금액           ");
            strSql.AppendLine("      , A2.TRMN_YN--해지여부                        ");
            strSql.AppendLine("      , A2.TRMN_YMD--해지일자                       ");
            strSql.AppendLine("      , A2.LIQ_CPTL_YN--유동성자금여부              ");
            strSql.AppendLine("      , A2.RPLC_PMNT_YN--대체자금계좌여부           ");
            strSql.AppendLine("      , A2.RECEIVE_ACNT_YN--외상매입금 입금 계좌여부");
            strSql.AppendLine("      , A2.NOTE--비고                               ");
            strSql.AppendLine("      , A1.ENT_DT                                   ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A1.ENT_ID AS NUMERIC) IS NULL THEN A1.ENT_ID ELSE DBO.FN_USRNM(A1.ENT_ID) END AS CUSER");
            strSql.AppendLine("      , A1.MFY_DT");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A1.MFY_ID AS NUMERIC) IS NULL THEN A1.MFY_ID ELSE DBO.FN_USRNM(A1.MFY_ID) END AS MUSER");
            strSql.AppendLine("   FROM ACC_DEALER_CD A1         ");
            strSql.AppendLine("   LEFT JOIN ACC_ACNT_CD A2      ");
            strSql.AppendLine("     ON A1.DEALER_CD = A2.ACNT_CD");
            strSql.AppendLine("  WHERE A1.BANKYN = 'Y'          ");
            strSql.AppendLine("    AND (('****' = '" + sAccCd + "') OR (('****' <> '" + sAccCd + "') AND A2.ACC_CD = '" + sAccCd + "'))");
            strSql.AppendLine("    AND (('****' = '" + sBankCd + "') OR (('****' <> '" + sBankCd + "') AND A1.BANK_CD = '" + sBankCd + "'))");
            strSql.AppendLine("    AND (('*' = '" + sCancel + "') OR(('*' <> '" + sCancel + "') AND A2.TRMN_YN = '" + sCancel + "'))");
            strSql.AppendLine("  ORDER BY A2.PSSEQ");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt != null)
            {
                GridRetr.DataSource = dt;
            }
            
            Cursor = Cursors.Default;
        }
        private void RdGrRetrCalcel_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }

        private void LkupEditAccCd_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }
        private void LkupEditBankNm_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }

        private void BtnRetrClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void BtnRetrAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if(GridViewRetr.RowCount > 0)
            {
                string sDEALER_NM = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount -1, "DEALER_NM")?.ToString();

                if (string.IsNullOrEmpty(sDEALER_NM))
                {
                    XtraMessageBox.Show("계좌명을 입력해주세요.");
                    return;
                }
            }
            

            GridViewRetr.AddNewRow();
            GridViewRetr.Focus();
            GridViewRetr.FocusedColumn = GridColDealernm;

            GridViewRetr.SetFocusedRowCellValue("TRMN_YN", "N");
            GridViewRetr.SetFocusedRowCellValue("LIQ_CPTL_YN", "N");
            GridViewRetr.SetFocusedRowCellValue("RPLC_PMNT_YN", "N");
            GridViewRetr.SetFocusedRowCellValue("RECEIVE_ACNT_YN", "N");
        }

        private void BtnRetrSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("편집한 데이터를 모두 저장하시겠습니까?"), "저장확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            Cursor = Cursors.WaitCursor;

            DataTable dt = (DataTable)GridRetr.DataSource;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sDEALER_NM = dtMerge.Rows[i]["DEALER_NM"]?.ToString();

                if (string.IsNullOrEmpty(sDEALER_NM))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            if(dtMerge.Rows.Count == 0)
            {
                XtraMessageBox.Show("변경된 데이터가 없습니다.");
                Cursor = Cursors.Default;
                return;
            }

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            try
            {
                if (dtMerge.Rows.Count > 0)
                {
                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        string sDEALER_CD = dtMerge.Rows[i]["DEALER_CD"]?.ToString();
                        string sDEALER_NM = dtMerge.Rows[i]["DEALER_NM"]?.ToString();
                        string sBANK_CD = dtMerge.Rows[i]["BANK_CD"]?.ToString();
                        string sBANK_ACNT_NO = dtMerge.Rows[i]["BANK_ACNT_NO"]?.ToString();
                        string sEOB_YN = dtMerge.Rows[i]["EOB_YN"]?.ToString();
                        string sPSSEQ = dtMerge.Rows[i]["PSSEQ"]?.ToString();
                        string sGEJAGB = dtMerge.Rows[i]["GEJAGB"]?.ToString();
                        string sACC_CD = dtMerge.Rows[i]["ACC_CD"]?.ToString();
                        string sRPLC_ACC_CD = dtMerge.Rows[i]["RPLC_ACC_CD"]?.ToString();
                        string sBRANCH_NM = dtMerge.Rows[i]["BRANCH_NM"]?.ToString();
                        string sFINANCE_GOODS_NM = dtMerge.Rows[i]["FINANCE_GOODS_NM"]?.ToString();
                        string sCNTR_YMD = dtMerge.Rows[i]["CNTR_YMD"]?.ToString();
                        if (!string.IsNullOrEmpty(sCNTR_YMD))
                            sCNTR_YMD = sCNTR_YMD.Substring(0, 10);
                        string sCNTR_AMT = dtMerge.Rows[i]["CNTR_AMT"]?.ToString();
                        string sINRST_RATE = dtMerge.Rows[i]["INRST_RATE"]?.ToString();
                        string sFRST_PAYIN_YMD = dtMerge.Rows[i]["FRST_PAYIN_YMD"]?.ToString();
                        if (!string.IsNullOrEmpty(sFRST_PAYIN_YMD))
                            sFRST_PAYIN_YMD = sFRST_PAYIN_YMD.Substring(0, 10);
                        string sPAYIN_AMT = dtMerge.Rows[i]["PAYIN_AMT"]?.ToString();
                        string sPAYIN_DD = dtMerge.Rows[i]["PAYIN_DD"]?.ToString();
                        string sPAYIN_METHOD = dtMerge.Rows[i]["PAYIN_METHOD"]?.ToString();
                        string sEXPIRE_YMD = dtMerge.Rows[i]["EXPIRE_YMD"]?.ToString();
                        if (!string.IsNullOrEmpty(sEXPIRE_YMD))
                            sEXPIRE_YMD = sEXPIRE_YMD.Substring(0, 10);
                        string sEXPIRE_INRST_AMT = dtMerge.Rows[i]["EXPIRE_INRST_AMT"]?.ToString();
                        string sTRMN_YN = dtMerge.Rows[i]["TRMN_YN"]?.ToString();
                        string sTRMN_YMD = dtMerge.Rows[i]["TRMN_YMD"]?.ToString();
                        if (!string.IsNullOrEmpty(sTRMN_YMD))
                            sTRMN_YMD = sTRMN_YMD.Substring(0, 10);
                        string sLIQ_CPTL_YN = dtMerge.Rows[i]["LIQ_CPTL_YN"]?.ToString();
                        string sRPLC_PMNT_YN = dtMerge.Rows[i]["RPLC_PMNT_YN"]?.ToString();
                        string sRECEIVE_ACNT_YN = dtMerge.Rows[i]["RECEIVE_ACNT_YN"]?.ToString();
                        string sNOTE = dtMerge.Rows[i]["NOTE"]?.ToString();
                        string sUSER = FmMainToolBar2.UserID;

                        double dDEALER_CD = 0;
                        double dCNTR_AMT         = 0;
                        double dINRST_RATE       = 0;
                        double dPAYIN_AMT        = 0;
                        double dEXPIRE_INRST_AMT = 0;
                        double dPSSEQ = 0;

                        double.TryParse(sDEALER_CD, out dDEALER_CD);
                        double.TryParse(sCNTR_AMT, out dCNTR_AMT);
                        double.TryParse(sINRST_RATE, out dINRST_RATE);
                        double.TryParse(sPAYIN_AMT, out dPAYIN_AMT);
                        double.TryParse(sEXPIRE_INRST_AMT, out dEXPIRE_INRST_AMT);
                        double.TryParse(sPSSEQ, out dPSSEQ);

                        if (string.IsNullOrEmpty(sBANK_ACNT_NO))
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("계좌번호 를 입력해주세요");
                            Cursor = Cursors.Default;
                            return;
                        }

                        if (string.IsNullOrEmpty(sBANK_CD))
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("은행을 선택해주세요");
                            Cursor = Cursors.Default;
                            return;
                        }

                        if (sTRMN_YN.Equals("Y") && !DateTime.TryParse(sTRMN_YMD, out DateTime dateResult))
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("해지일자를 입력하세요");
                            return;
                        }

                        strSql.Clear();
                        strSql.AppendLine("SET IDENTITY_INSERT ACC_DEALER_CD ON");
                        strSql.AppendLine(" IF EXISTS(SELECT* FROM ACC_DEALER_CD WHERE DEALER_CD = "+ dDEALER_CD + ")        ");
                        strSql.AppendLine("     BEGIN                                                                 ");
                        strSql.AppendLine("           UPDATE ACC_DEALER_CD                                            ");
                        strSql.AppendLine("              SET DEALER_NM = '"+sDEALER_NM+"'                                   ");
                        strSql.AppendLine(" 			   , BANK_CD = '"+sBANK_CD+"'                                       ");
                        strSql.AppendLine(" 			   , BANK_ACNT_NO = '"+ sBANK_ACNT_NO + "'                                  ");
                        strSql.AppendLine(" 			   , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20)             ");
                        strSql.AppendLine(" 			   , MFY_ID = '"+sUSER+"'                                           ");
                        strSql.AppendLine("            WHERE DEALER_CD = "+dDEALER_CD+"                                   ");
                        strSql.AppendLine("                                                                           ");
                        strSql.AppendLine("            IF EXISTS(SELECT* FROM ACC_ACNT_CD WHERE ACNT_CD = "+dDEALER_CD+") ");
                        strSql.AppendLine("               BEGIN                                                       ");
                        strSql.AppendLine("                     UPDATE ACC_ACNT_CD                                    ");
                        strSql.AppendLine("                        SET PSSEQ = "+dPSSEQ+"                                 ");
                        strSql.AppendLine(" 						 , GEJAGB = '"+sGEJAGB+"'                               ");
                        strSql.AppendLine(" 						 , ACC_CD = '"+sACC_CD+"'                               ");
                        strSql.AppendLine(" 						 , RPLC_ACC_CD = '"+sRPLC_ACC_CD+"'                     ");
                        strSql.AppendLine(" 						 , BRANCH_NM = '"+sBRANCH_NM+"'                         ");
                        strSql.AppendLine(" 						 , FINANCE_GOODS_NM = '"+ sFINANCE_GOODS_NM + "'           ");
                        strSql.AppendLine(" 						 , CNTR_YMD = '"+ sCNTR_YMD + "'                           ");
                        strSql.AppendLine(" 						 , CNTR_AMT = "+ dCNTR_AMT + "                           ");
                        strSql.AppendLine(" 						 , INRST_RATE = "+ dINRST_RATE + "                       ");
                        strSql.AppendLine(" 						 , FRST_PAYIN_YMD = '"+ sFRST_PAYIN_YMD + "'              ");
                        strSql.AppendLine(" 						 , PAYIN_AMT = "+ dPAYIN_AMT + "                         ");
                        strSql.AppendLine(" 						 , PAYIN_DD = '"+ sPAYIN_DD + "'                           ");
                        strSql.AppendLine(" 						 , PAYIN_METHOD = '"+ sPAYIN_METHOD + "'                   ");
                        strSql.AppendLine(" 						 , EXPIRE_YMD = '"+ sEXPIRE_YMD + "'                       ");
                        strSql.AppendLine(" 						 , EXPIRE_INRST_AMT = "+ dEXPIRE_INRST_AMT + "           ");
                        strSql.AppendLine(" 						 , TRMN_YN = '"+ sTRMN_YN + "'                             ");
                        strSql.AppendLine(" 						 , TRMN_YMD = '"+ sTRMN_YMD + "'                           ");
                        strSql.AppendLine(" 						 , LIQ_CPTL_YN = '"+ sLIQ_CPTL_YN + "'                     ");
                        strSql.AppendLine(" 						 , RPLC_PMNT_YN = '"+ sRPLC_PMNT_YN + "'                   ");
                        strSql.AppendLine(" 						 , RECEIVE_ACNT_YN = '"+ sRECEIVE_ACNT_YN + "'             ");
                        strSql.AppendLine(" 						 , NOTE = '"+ sNOTE + "'                                   ");
                        strSql.AppendLine(" 						 , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20)   ");
                        strSql.AppendLine(" 						 , MFY_ID = '"+ sUSER + "'                                 ");
                        strSql.AppendLine("                      WHERE ACNT_CD = "+dDEALER_CD+"                           ");
                        strSql.AppendLine("                 END                                                       ");
                        strSql.AppendLine("            ELSE                                                           ");
                        strSql.AppendLine("               BEGIN                                                       ");
                        strSql.AppendLine("                     INSERT INTO ACC_ACNT_CD(ACNT_CD                       ");
                        strSql.AppendLine("                                             , PSSEQ                       ");
                        strSql.AppendLine("                                             , GEJAGB                      ");
                        strSql.AppendLine("                                             , ACC_CD                      ");
                        strSql.AppendLine("                                             , RPLC_ACC_CD                 ");
                        strSql.AppendLine("                                             , BRANCH_NM                   ");
                        strSql.AppendLine("                                             , FINANCE_GOODS_NM            ");
                        strSql.AppendLine("                                             , CNTR_YMD                    ");
                        strSql.AppendLine("                                             , CNTR_AMT                    ");
                        strSql.AppendLine("                                             , INRST_RATE                  ");
                        strSql.AppendLine("                                             , FRST_PAYIN_YMD              ");
                        strSql.AppendLine("                                             , PAYIN_AMT                   ");
                        strSql.AppendLine("                                             , PAYIN_DD                    ");
                        strSql.AppendLine("                                             , PAYIN_METHOD                ");
                        strSql.AppendLine("                                             , EXPIRE_YMD                  ");
                        strSql.AppendLine("                                             , EXPIRE_INRST_AMT            ");
                        strSql.AppendLine("                                             , TRMN_YN                     ");
                        strSql.AppendLine("                                             , TRMN_YMD                    ");
                        strSql.AppendLine("                                             , LIQ_CPTL_YN                 ");
                        strSql.AppendLine("                                             , RPLC_PMNT_YN                ");
                        strSql.AppendLine("                                             , RECEIVE_ACNT_YN             ");
                        strSql.AppendLine("                                             , NOTE                        ");
                        strSql.AppendLine("                                             , ENT_DT                      ");
                        strSql.AppendLine("                                             , ENT_ID )                    ");
                        strSql.AppendLine("                                       VALUES( "+dDEALER_CD+"                    ");
                        strSql.AppendLine("                                             , "+dPSSEQ+"                       ");
                        strSql.AppendLine("                                             , '"+sGEJAGB+"'                     ");
                        strSql.AppendLine("                                             , '"+sACC_CD+"'                     ");
                        strSql.AppendLine("                                             , '"+sRPLC_ACC_CD+"'                ");
                        strSql.AppendLine("                                             , '"+ sBRANCH_NM + "'                  ");
                        strSql.AppendLine("                                             , '"+ sFINANCE_GOODS_NM + "'           ");
                        strSql.AppendLine("                                             , '"+ sCNTR_YMD + "'                   ");
                        strSql.AppendLine("                                             , "+ dCNTR_AMT + "                   ");
                        strSql.AppendLine("                                             , "+ dINRST_RATE + "                 ");
                        strSql.AppendLine("                                             , '"+ sFRST_PAYIN_YMD + "'             ");
                        strSql.AppendLine("                                             , "+ dPAYIN_AMT + "                  ");
                        strSql.AppendLine("                                             , '"+ sPAYIN_DD + "'                   ");
                        strSql.AppendLine("                                             , '"+ sPAYIN_METHOD + "'               ");
                        strSql.AppendLine("                                             , '"+ sEXPIRE_YMD + "'                 ");
                        strSql.AppendLine("                                             , "+ dEXPIRE_INRST_AMT + "           ");
                        strSql.AppendLine("                                             , '"+ sTRMN_YN + "'                    ");
                        strSql.AppendLine("                                             , '"+ sTRMN_YMD + "'                   ");
                        strSql.AppendLine("                                             , '"+ sLIQ_CPTL_YN + "'                ");
                        strSql.AppendLine("                                             , '"+ sRPLC_PMNT_YN + "'               ");
                        strSql.AppendLine("                                             , '"+ sRECEIVE_ACNT_YN + "'            ");
                        strSql.AppendLine("                                             , '"+ sNOTE + "'                       ");
                        strSql.AppendLine("                                             , CONVERT(VARCHAR(20), GETDATE(), 20)");
                        strSql.AppendLine("                                             , '"+ sUSER + "')                             ");
                        strSql.AppendLine("                 END                               ");
                        strSql.AppendLine("       END                                         ");
                        strSql.AppendLine("                                                   ");
                        strSql.AppendLine(" ELSE                                              ");
                        strSql.AppendLine("     BEGIN                                         ");
                        strSql.AppendLine("                                                   ");
                        strSql.AppendLine("          DECLARE @DEALER_CD NUMERIC;");
                        strSql.AppendLine("           SELECT @DEALER_CD = MAX(DEALER_CD) + 1  ");
                        strSql.AppendLine("             FROM ACC_DEALER_CD                    ");
                        strSql.AppendLine("                                                   ");
                        strSql.AppendLine("           INSERT INTO ACC_DEALER_CD( DEALER_CD     ");
                        strSql.AppendLine("                                    , DEALER_NM    ");
                        strSql.AppendLine("                                    , DEALER_GB    ");
                        strSql.AppendLine("                                    , BANK_CD      ");
                        strSql.AppendLine("                                    , BANK_ACNT_NO ");
                        strSql.AppendLine("                                    , BANKYN       ");
                        strSql.AppendLine("                                    , ENT_DT       ");
                        strSql.AppendLine("                                    , ENT_ID)      ");
                        strSql.AppendLine("                              VALUES( @DEALER_CD    ");
                        strSql.AppendLine("                                    , '"+sDEALER_NM+"'   ");
                        strSql.AppendLine("                                    , '계좌'       ");
                        strSql.AppendLine("                                    , '"+sBANK_CD+"'     ");
                        strSql.AppendLine("                                    , '"+sBANK_ACNT_NO+"'");
                        strSql.AppendLine("                                    , 'Y'      ");
                        strSql.AppendLine("                                    , CONVERT(VARCHAR(20), GETDATE(), 20)  ");
                        strSql.AppendLine("                                    , '"+ sUSER + "')                               ");
                        strSql.AppendLine("                                                                           ");
                        strSql.AppendLine("                      INSERT INTO ACC_ACNT_CD( ACNT_CD                     ");
                        strSql.AppendLine("                                             , PSSEQ                       ");
                        strSql.AppendLine("                                             , GEJAGB                      ");
                        strSql.AppendLine("                                             , ACC_CD                      ");
                        strSql.AppendLine("                                             , RPLC_ACC_CD                 ");
                        strSql.AppendLine("                                             , BRANCH_NM                   ");
                        strSql.AppendLine("                                             , FINANCE_GOODS_NM            ");
                        strSql.AppendLine("                                             , CNTR_YMD                    ");
                        strSql.AppendLine("                                             , CNTR_AMT                    ");
                        strSql.AppendLine("                                             , INRST_RATE                  ");
                        strSql.AppendLine("                                             , FRST_PAYIN_YMD              ");
                        strSql.AppendLine("                                             , PAYIN_AMT                   ");
                        strSql.AppendLine("                                             , PAYIN_DD                    ");
                        strSql.AppendLine("                                             , PAYIN_METHOD                ");
                        strSql.AppendLine("                                             , EXPIRE_YMD                  ");
                        strSql.AppendLine("                                             , EXPIRE_INRST_AMT            ");
                        strSql.AppendLine("                                             , TRMN_YN                     ");
                        strSql.AppendLine("                                             , TRMN_YMD                    ");
                        strSql.AppendLine("                                             , LIQ_CPTL_YN                 ");
                        strSql.AppendLine("                                             , RPLC_PMNT_YN                ");
                        strSql.AppendLine("                                             , RECEIVE_ACNT_YN             ");
                        strSql.AppendLine("                                             , NOTE                        ");
                        strSql.AppendLine("                                             , ENT_DT                      ");
                        strSql.AppendLine("                                             , ENT_ID )                    ");
                        strSql.AppendLine("                                       VALUES( @DEALER_CD                   ");
                        strSql.AppendLine("                                             , " + dPSSEQ + "                       ");
                        strSql.AppendLine("                                             , '" + sGEJAGB + "'                     ");
                        strSql.AppendLine("                                             , '" + sACC_CD + "'                     ");
                        strSql.AppendLine("                                             , '" + sRPLC_ACC_CD + "'                ");
                        strSql.AppendLine("                                             , '" + sBRANCH_NM + "'                  ");
                        strSql.AppendLine("                                             , '" + sFINANCE_GOODS_NM + "'           ");
                        strSql.AppendLine("                                             , '" + sCNTR_YMD + "'                   ");
                        strSql.AppendLine("                                             , " + dCNTR_AMT + "                   ");
                        strSql.AppendLine("                                             , " + dINRST_RATE + "                 ");
                        strSql.AppendLine("                                             , '" + sFRST_PAYIN_YMD + "'             ");
                        strSql.AppendLine("                                             , " + dPAYIN_AMT + "                  ");
                        strSql.AppendLine("                                             , '" + sPAYIN_DD + "'                   ");
                        strSql.AppendLine("                                             , '" + sPAYIN_METHOD + "'               ");
                        strSql.AppendLine("                                             , '" + sEXPIRE_YMD + "'                 ");
                        strSql.AppendLine("                                             , " + dEXPIRE_INRST_AMT + "           ");
                        strSql.AppendLine("                                             , '" + sTRMN_YN + "'                    ");
                        strSql.AppendLine("                                             , '" + sTRMN_YMD + "'                   ");
                        strSql.AppendLine("                                             , '" + sLIQ_CPTL_YN + "'                ");
                        strSql.AppendLine("                                             , '" + sRPLC_PMNT_YN + "'               ");
                        strSql.AppendLine("                                             , '" + sRECEIVE_ACNT_YN + "'            ");
                        strSql.AppendLine("                                             , '" + sNOTE + "'                       ");
                        strSql.AppendLine("                                             , CONVERT(VARCHAR(20), GETDATE(), 20)");
                        strSql.AppendLine("                                             , '" + sUSER + "')                             ");
                        strSql.AppendLine("                                                                           ");
                        strSql.AppendLine("       END");
                        strSql.AppendLine("SET IDENTITY_INSERT ACC_DEALER_CD OFF");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");
                GetGridRetr();

                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColAccNumber, dtMerge.Rows[0]["BANK_ACNT_NO"].ToString());
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
            Cursor = Cursors.Default;
        }

        private void GridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.Equals("CNTR_YMD"))
            {
                string sYmd = e.Value.ToString().Replace("-", "").Substring(0, 8);
                e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
            }
        }

        private void repoDateEdit_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value.ToString() == "") return;
            if (e.Value.ToString() == "0") return;
            string sYmd = e.Value.ToString().Replace("-", "").Substring(0, 8);
            e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
        }

        private void GridViewRetr_ShowingPopupEditForm(object sender, DevExpress.XtraGrid.Views.Grid.ShowingPopupEditFormEventArgs e)
        {
            e.EditForm.ImeMode = ImeMode.Hangul;
            e.EditForm.StartPosition = FormStartPosition.CenterParent;

            foreach (var button in e.EditForm.Controls.OfType<EditFormContainer>()
                        .SelectMany(control => Enumerable.Cast<Control>(control.Controls)).OfType<PanelControl>()
                        .SelectMany(nestedControl => Enumerable.Cast<Control>(nestedControl.Controls)).OfType<SimpleButton>())
            {
                switch (button.Text)
                {
                    case "Update":
                        //button.Visible = false;
                        button.Text = "수정";
                        //button.Click += EditFormUpdateButton_Click;
                        break;

                    case "Cancel":
                        button.Text = "취소";
                        break;
                }
            }

        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void AccAccountNumberMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {

            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnRetrAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnRetrSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
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

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccAccountNumberMgt_FormClosed(object sender, FormClosedEventArgs e)
        {
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccAccountNumberMgt", GridViewRetr);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int idx = GridViewRetr.FocusedRowHandle;
            string sDEALER_CD = GridViewRetr.GetFocusedRowCellValue("DEALER_CD")?.ToString();
            string sDEALER_NM = GridViewRetr.GetFocusedRowCellValue("DEALER_NM")?.ToString();
            string sBankAcntNo = GridViewRetr.GetFocusedRowCellValue("BANK_ACNT_NO")?.ToString();

            if (XtraMessageBox.Show("계좌명: "+ sDEALER_NM + "\r\n계좌번호 : " + sBankAcntNo + "\r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
               , "은행계좌 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM ACC_ACNT_CD ");
                strSql.AppendLine("       WHERE ACNT_CD = " + sDEALER_CD + "");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM ACC_DEALER_CD ");
                strSql.AppendLine("       WHERE DEALER_CD = " + sDEALER_CD + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                //string sLogRmk = "Table:ACC_ACNT_CD -> BANK_ACNT_NO:" + sBankAcntNo;
                //ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");
                BtnRetr_Click(null, null);

                GridViewRetr.FocusedRowHandle = idx - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
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
                string sFileNM = "은행계좌리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridRetr.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
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

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr.UpdateCurrentRow();
        }

        private void LkupEditBankNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void AccAccountNumberMgt_TextChanged(object sender, EventArgs e)
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
    }
}