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
using System.Diagnostics;
using System.Data.SqlClient;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*/
namespace AccAdm
{
    public partial class AccPaymentCode : DevExpress.XtraEditors.XtraForm
    {
        public AccPaymentCode()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        private void AccPaymentCode_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccPaymentCode");

            SetLookUpEdit(LkupEditGb, "1", "N", "Y", "");
            SetLookUpEdit(LkupEditRGb, "1", "N", "Y", "");
            SetLookUpEdit(LkupEditCalc, "3", "N", "Y", "");
            SetLookUpEdit(LkupEditGea, "4", "N", "Y", "");

            ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccPaymentCode", GridViewRetr);
        }

        private void LkupEditGb_EditValueChanged(object sender, EventArgs e)
        {
            string sGubun = LkupEditGb.EditValue.ToString();
            SetLookUpEdit(LkupEditDetGb, "2", "N", "Y", sGubun);
            BtnRetr_Click(null, null);
        }

        private void LkupEditRGb_EditValueChanged(object sender, EventArgs e)
        {
            string sGubun = "";

            sGubun = LkupEditRGb.EditValue.ToString();

            SetLookUpEdit(LkupEditRDetGb, "2", "N", "Y", sGubun);
        }
        private void LkupEditDetGb_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }
        private void SetLookUpEdit(DevExpress.XtraEditors.LookUpEdit lkup, string sGb, string sNullYn, string sSetIdx, string sGubun)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.AppendLine("");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine("  SELECT '' AS CD ");
                strSql.AppendLine("       , '전체' AS NM ");
                strSql.AppendLine("   UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD     ");
                strSql.AppendLine("      , COM_NM AS NM     ");
                strSql.AppendLine("   FROM COM_BASE_CD      ");
                strSql.AppendLine("  WHERE CD_GB = 'ACPAYGB'");

            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD       ");
                strSql.AppendLine("      , COM_NM AS NM       ");
                strSql.AppendLine("   FROM COM_BASE_CD        ");
                strSql.AppendLine("  WHERE CD_GB = 'ACPAYDTL" + sGubun + "'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine("   SELECT COM_CD AS CD ");
                strSql.AppendLine("        , COM_NM AS NM ");
                strSql.AppendLine("     FROM COM_BASE_CD");
                strSql.AppendLine("    WHERE CD_GB = 'CALC_GB'  ");
            }
            else
            {
                strSql.AppendLine(" SELECT ACCOD AS CD    ");
                strSql.AppendLine("      , ACNAM AS NM    ");
                strSql.AppendLine("   FROM ACMSTF     ");
                strSql.AppendLine("  WHERE USEYN = 'Y'");
            }

            strSql.AppendLine("   ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            lkup.Properties.DataSource = dt;
            lkup.Properties.DisplayMember = "NM";
            lkup.Properties.ValueMember = "CD";

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            ClearGrid();

            string sGubun = LkupEditGb.EditValue.ToString();
            string sDetGubun = LkupEditDetGb.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();
            strSql.AppendLine("SELECT A.GUBUN ");//
            strSql.AppendLine("     , A.DET_GB");//
            strSql.AppendLine("     , A.PAY_CD");
            strSql.AppendLine("     , A.PAY_CD_NM");
            strSql.AppendLine("     , A.STRT_YM  ");
            strSql.AppendLine("     , A.END_YM  ");
            strSql.AppendLine("     , B.COM_NM AS CALC_GB ");
            strSql.AppendLine("     , A.PAY_BASIS_YN");
            strSql.AppendLine("     , A.TAX_EXCEPT_YN  ");
            strSql.AppendLine("     , A.TAX_EXCEPT_LIMIT ");
            strSql.AppendLine("     , A.GEA");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 1, 1) AS PAY_MON1 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 2, 1) AS PAY_MON2 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 3, 1) AS PAY_MON3 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 4, 1) AS PAY_MON4 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 5, 1) AS PAY_MON5 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 6, 1) AS PAY_MON6 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 7, 1) AS PAY_MON7 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 8, 1) AS PAY_MON8 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 9, 1) AS PAY_MON9 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 10, 1) AS PAY_MON10 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 11, 1) AS PAY_MON11 ");
            strSql.AppendLine("     , SUBSTRING(PAY_MONTH, 12, 1) AS PAY_MON12");
            strSql.AppendLine("     , A.CALC_RATE  ");
            strSql.AppendLine("     , A.RMK  ");
            strSql.AppendLine("     , A.AUTO_YN");
            strSql.AppendLine("     , A.HU_PAY_YN");
            strSql.AppendLine("     , A.PRT_GB  ");
            strSql.AppendLine("     , A.HU_PAY_BASIS_YN ");
            strSql.AppendLine("  FROM ACC_PAY_CD A ");
            strSql.AppendLine("  LEFT OUTER JOIN   COM_BASE_CD B ");
            strSql.AppendLine("    ON A.CALC_GB = B.COM_CD ");
            strSql.AppendLine("   AND B.CD_GB = 'CALC_GB' ");
            strSql.AppendLine(" WHERE A.GUBUN = '" + sGubun + "'");
            strSql.AppendLine("   AND (('' = '" + sDetGubun + "') OR(('' <> '" + sDetGubun + "') AND A.DET_GB = '" + sDetGubun + "'))");
            strSql.AppendLine(" ORDER BY  GUBUN, DET_GB, PAY_CD ");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
            Cursor = Cursors.Default;
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            LkupEditRGb.EditValue = GridViewRetr.GetFocusedRowCellValue("GUBUN").ToString();
            LkupEditRDetGb.EditValue = GridViewRetr.GetFocusedRowCellValue("DET_GB").ToString();
            TxtCd.Text = GridViewRetr.GetFocusedRowCellValue("PAY_CD").ToString();
            TxtCdName.Text = GridViewRetr.GetFocusedRowCellValue("PAY_CD_NM").ToString();
            TxtStart.Text = GridViewRetr.GetFocusedRowCellValue("STRT_YM").ToString();
            TxtEnd.Text = GridViewRetr.GetFocusedRowCellValue("END_YM").ToString();
            LkupEditCalc.Text = GridViewRetr.GetFocusedRowCellValue("CALC_GB").ToString();
            TxtCalcRate.Text = GridViewRetr.GetFocusedRowCellValue("CALC_RATE").ToString();
            LkupEditGea.EditValue = GridViewRetr.GetFocusedRowCellValue("GEA");
            TxtTaxExceptLimit.Text = GridViewRetr.GetFocusedRowCellValue("TAX_EXCEPT_LIMIT").ToString();

            ChkPayBasis.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_BASIS_YN");
            ChkTaxExcept.EditValue = GridViewRetr.GetFocusedRowCellValue("TAX_EXCEPT_YN");

            ChkPayMon1.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON1");
            ChkPayMon2.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON2");
            ChkPayMon3.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON3");
            ChkPayMon4.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON4");
            ChkPayMon5.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON5");
            ChkPayMon6.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON6");
            ChkPayMon7.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON7");
            ChkPayMon8.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON8");
            ChkPayMon9.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON9");
            ChkPayMon10.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON10");
            ChkPayMon11.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON11");
            ChkPayMon12.EditValue = GridViewRetr.GetFocusedRowCellValue("PAY_MON12");

            Memobigo.Text = GridViewRetr.GetFocusedRowCellValue("RMK").ToString();
            ChkAuto.EditValue = GridViewRetr.GetFocusedRowCellValue("AUTO_YN");
            ChkHuPayBasisYn.EditValue = GridViewRetr.GetFocusedRowCellValue("HU_PAY_BASIS_YN");
            ChkHuPayYn.EditValue = GridViewRetr.GetFocusedRowCellValue("HU_PAY_YN");
            TxtPrtGb.Text = GridViewRetr.GetFocusedRowCellValue("PRT_GB").ToString();
        }
        private void ClearGrid()
        {
            GridViewRetr.FocusedRowChanged -= GridViewRetr_FocusedRowChanged;
            GridRetr.DataSource = null;
            GridViewRetr.FocusedRowChanged += GridViewRetr_FocusedRowChanged;
        }
        public void ClearAllForm(Control Ctrl)
        {
            if (Ctrl.HasChildren)
            {
                foreach (Control ctrl in Ctrl.Controls)
                {
                    if (ctrl is DevExpress.XtraEditors.TextEdit)
                        (ctrl as DevExpress.XtraEditors.TextEdit).ResetText();

                    LkupEditRGb.EditValueChanged -= LkupEditRGb_EditValueChanged;

                    if (ctrl is DevExpress.XtraEditors.LookUpEdit)
                        (ctrl as DevExpress.XtraEditors.LookUpEdit).EditValue = null;

                    LkupEditRGb.EditValueChanged += LkupEditRGb_EditValueChanged;

                    if (ctrl is DevExpress.XtraEditors.DateEdit)
                        (ctrl as DevExpress.XtraEditors.DateEdit).ResetText();

                    if (ctrl is DevExpress.XtraEditors.ComboBoxEdit)
                        (ctrl as DevExpress.XtraEditors.ComboBoxEdit).ResetText();

                    if (ctrl is DevExpress.XtraEditors.CheckEdit)
                        (ctrl as DevExpress.XtraEditors.CheckEdit).Checked = false;
                    
                      

                    if (ctrl.HasChildren)
                        ClearAllForm(ctrl);//Recursive
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (LkupEditRGb.Text.Length < 1)
            {
                MessageBox.Show("코드구분을 선택하세요 ");
                return;
            }
            if (LkupEditRDetGb.Text.Length < 1)
            {
                MessageBox.Show("세부구분을 선택하세요 ");
                return;
            }
            if (TxtCd.Text.Length < 1)
            {
                MessageBox.Show("코드를 입력하세요 ");
                return;
            }
            if (TxtCdName.Text.Length < 1)
            {
                MessageBox.Show("코드명을 입력하세요 ");
                return;
            }
            if (TxtStart.Text.Length <= 5)
            {
                MessageBox.Show("시작월을 입력하세요 xxxx년xx월 숫자만 입력하세요 ");
                return;
            }
            if (TxtEnd.Text.Length <= 5)
            {
                MessageBox.Show("종료월을  입력하세요 xxxx년xx월 숫자만 입력하세요 ");
                return;
            }
            if (LkupEditCalc.Text.Length < 1)
            {
                MessageBox.Show("지급방법을  선택하세요 ");
                return;
            }
            if (LkupEditGea.Text.Length < 1)
            {
                MessageBox.Show("계정과목을  선택하세요 ");
                return;
            }

            if ((MessageBox.Show(this, "저장 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
            {
                return;
            }
           

            string result = string.Empty;
            for (int i = 0; i <= 11; i++)
            {
                CheckEdit[] chck = { ChkPayMon1, ChkPayMon2, ChkPayMon3, ChkPayMon4, ChkPayMon5, ChkPayMon6,
                                 ChkPayMon7, ChkPayMon8, ChkPayMon9, ChkPayMon10, ChkPayMon11, ChkPayMon12};

                if (chck[i].Checked == true) result += "1";
                else result += "0";
            }

            string sGubun = LkupEditRGb.EditValue?.ToString();
            string sDetGb = LkupEditRDetGb.EditValue?.ToString();
            string sPayCd = TxtCd.Text;
            string sPayCdNm = TxtCdName.Text;
            string sStrtYm = TxtStart.Text;
            string sEndYm = TxtEnd.Text;
            string sGubunNm = LkupEditRGb.Text;
            string sDetGbNm = LkupEditRDetGb.Text;
            string sCalcGb = LkupEditCalc.EditValue?.ToString();
            string sCalcRate = TxtCalcRate.EditValue.Equals("")? "0": TxtCalcRate.EditValue.ToString();
            string sPayBasisYn = ChkPayBasis.Checked ? "Y" : "N";
            string sTaxExceptYn = ChkTaxExcept.Checked ? "1" : "0";
            string sTaxExceptLimit = TxtTaxExceptLimit?.Text.Replace(",", "");
            string sPayMonth = result;
            string sGea = LkupEditGea.EditValue?.ToString();
            string sRmk = Memobigo.Text;
            string sId = FmMainToolBar2.drUser["USRCD"]?.ToString();
            string sAutoYn = ChkAuto.Checked ? "1" : "0";
            string sHuPayYn = ChkHuPayYn.Checked ? "Y" : "N";
            string sPrtGb = TxtPrtGb?.Text;
            string sHuPayBasisYn = ChkHuPayBasisYn.Checked ? "Y" : "N";
            Cursor = Cursors.WaitCursor;
            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();
            try
            {
               strSql.AppendLine(" MERGE INTO ACC_PAY_CD AS a                                                                          ");
               strSql.AppendLine("     USING(SELECT                                                                                    ");
               strSql.AppendLine("                 GUBUN = '" + sGubun + "'                                                                          ");
               strSql.AppendLine("               , DET_GB = '" + sDetGb + "'                                                                         ");
               strSql.AppendLine("               , PAY_CD = '" + sPayCd + "'                                                                         ");
               strSql.AppendLine("               , STRT_YM = '" + sStrtYm + "'                                                                        ");
               strSql.AppendLine("               , END_YM = '" + sEndYm + "'                                                                   ");
               strSql.AppendLine("               , GUBUN_NM = '" + sGubunNm + "'                                                                   ");
               strSql.AppendLine("               , DET_GB_NM = '" + sDetGbNm + "'                                                                  ");
               strSql.AppendLine("               , PAY_CD_NM = '" + sPayCdNm + "'                                                                     ");
               strSql.AppendLine("               , CALC_GB = '" + sCalcGb + "'                                                                       ");
               strSql.AppendLine("               , CALC_RATE = '" + sCalcRate + "'                                                                     ");
               strSql.AppendLine("               , PAY_BASIS_YN = '" + sPayBasisYn + "'                                                                  ");
               strSql.AppendLine("               , PAY_MONTH = '" + sPayMonth + "'                                                          ");
               strSql.AppendLine("               , RMK = '" + sRmk + "'                                                                            ");
               strSql.AppendLine("               , MODIFYID = '" + sId + "'                                                                    ");
               strSql.AppendLine("               , MODIFYDT = CONVERT(VARCHAR(20),GETDATE(),20)                                                                ");
               strSql.AppendLine("               , GEA = '" + sGea + "'                                                                       ");
               strSql.AppendLine("               , TAX_EXCEPT_YN = '" + sTaxExceptYn + "'                                                                 ");
               strSql.AppendLine("               , TAX_EXCEPT_LIMIT = '" + sTaxExceptLimit + "'                                                           ");
               strSql.AppendLine("               , AUTO_YN = '" + sAutoYn + "'                                                                       ");
               strSql.AppendLine("               , HU_PAY_YN = '" + sHuPayYn + "'                                                                     ");
               strSql.AppendLine("               , PRT_GB = '" + sPrtGb + "'                                                                         ");
               strSql.AppendLine("               , HU_PAY_BASIS_YN = '" + sHuPayBasisYn + "') AS b                                                         ");
               strSql.AppendLine("     ON a.GUBUN = b.GUBUN AND a.DET_GB = b.DET_GB AND a.PAY_CD = b.PAY_CD AND a.STRT_YM = b.STRT_YM  ");
               strSql.AppendLine("     WHEN MATCHED THEN UPDATE SET                                                                    ");
               strSql.AppendLine("                 GUBUN = '" + sGubun + "'                                                                          ");
               strSql.AppendLine("               , DET_GB = '" + sDetGb + "'                                                                         ");
               strSql.AppendLine("               , PAY_CD = '" + sPayCd + "'                                                                         ");
               strSql.AppendLine("               , STRT_YM = '" + sStrtYm + "'                                                                        ");
               strSql.AppendLine("               , END_YM = '" + sEndYm + "'                                                                   ");
               strSql.AppendLine("               , GUBUN_NM = '" + sGubunNm + "'                                                                   ");
               strSql.AppendLine("               , DET_GB_NM = '" + sDetGbNm + "'                                                                  ");
               strSql.AppendLine("               , PAY_CD_NM = '" + sPayCdNm + "'                                                                     ");
               strSql.AppendLine("               , CALC_GB = '" + sCalcGb + "'                                                                       ");
               strSql.AppendLine("               , CALC_RATE = '" + sCalcRate + "'                                                                     ");
               strSql.AppendLine("               , PAY_BASIS_YN = '" + sPayBasisYn + "'                                                                  ");
               strSql.AppendLine("               , PAY_MONTH = '" + sPayMonth + "'                                                          ");
               strSql.AppendLine("               , RMK = '" + sRmk + "'                                                                            ");
               strSql.AppendLine("               , MODIFYID = '" + sId + "'                                                                    ");
               strSql.AppendLine("               , MODIFYDT = CONVERT(VARCHAR(20),GETDATE(),20)                                                                   ");
               strSql.AppendLine("               , GEA = '" + sGea + "'                                                                       ");
               strSql.AppendLine("               , TAX_EXCEPT_YN = '" + sTaxExceptYn + "'                                                                 ");
               strSql.AppendLine("               , TAX_EXCEPT_LIMIT = '" + sTaxExceptLimit + "'                                                           ");
               strSql.AppendLine("               , AUTO_YN = '" + sAutoYn + "'                                                                       ");
               strSql.AppendLine("               , HU_PAY_YN = '" + sHuPayYn + "'                                                                     ");
               strSql.AppendLine("               , PRT_GB = '" + sPrtGb + "'                                                                         ");
               strSql.AppendLine("               , HU_PAY_BASIS_YN = '" + sHuPayBasisYn + "'                                                               ");
               strSql.AppendLine("     WHEN NOT MATCHED THEN INSERT(                                                                   ");
               strSql.AppendLine("                  GUBUN                                                                                       ");
               strSql.AppendLine("               , DET_GB                                                                              ");
               strSql.AppendLine("               , PAY_CD                                                                              ");
               strSql.AppendLine("               , STRT_YM                                                                             ");
               strSql.AppendLine("               , END_YM                                                                              ");
               strSql.AppendLine("               , GUBUN_NM                                                                            ");
               strSql.AppendLine("               , DET_GB_NM                                                                           ");
               strSql.AppendLine("               , PAY_CD_NM                                                                           ");
               strSql.AppendLine("               , CALC_GB                                                                             ");
               strSql.AppendLine("               , CALC_RATE                                                                           ");
               strSql.AppendLine("               , PAY_BASIS_YN                                                                        ");
               strSql.AppendLine("               , PAY_MONTH                                                                           ");
               strSql.AppendLine("               , RMK                                                                                 ");
               strSql.AppendLine("               , MODIFYID                                                                            ");
               strSql.AppendLine("               , MODIFYDT                                                                            ");
               strSql.AppendLine("               , GEA                                                                                 ");
               strSql.AppendLine("               , TAX_EXCEPT_YN                                                                       ");
               strSql.AppendLine("               , TAX_EXCEPT_LIMIT                                                                    ");
               strSql.AppendLine("               , AUTO_YN                                                                             ");
               strSql.AppendLine("               , HU_PAY_YN                                                                           ");
               strSql.AppendLine("               , PRT_GB                                                                              ");
               strSql.AppendLine("               , HU_PAY_BASIS_YN)                                                                    ");
               strSql.AppendLine("     VALUES(                                                                                         ");
               strSql.AppendLine("               '" + sGubun + "'                                                                                         ");
               strSql.AppendLine("               , '" + sDetGb + "'                                                                                ");
               strSql.AppendLine("               , '" + sPayCd + "'                                                                                ");
               strSql.AppendLine("               , '" + sStrtYm + "'                                                                            ");
               strSql.AppendLine("               , '" + sEndYm + "'                                                                            ");
               strSql.AppendLine("               , '" + sGubunNm + "'                                                                              ");
               strSql.AppendLine("               , '" + sDetGbNm + "'                                                                              ");
               strSql.AppendLine("               , '" + sPayCdNm + "'                                                                                 ");
               strSql.AppendLine("               , '" + sCalcGb + "'                                                                                 ");
               strSql.AppendLine("               , '" + sCalcRate + "'                                                                                 ");
               strSql.AppendLine("               , '" + sPayBasisYn + "'                                                                                 ");
               strSql.AppendLine("               , '" + sPayMonth + "'                                                                      ");
               strSql.AppendLine("               , '" + sRmk + "'                                                                                  ");
               strSql.AppendLine("               , '" + sId + "'                                                                               ");
               strSql.AppendLine("               , CONVERT(VARCHAR(20),GETDATE(),20)                                                                              ");
               strSql.AppendLine("               , '" + sGea + "'                                                                             ");
               strSql.AppendLine("               , '" + sTaxExceptYn + "'                                                                                 ");
               strSql.AppendLine("               , '" + sTaxExceptLimit + "'                                                                              ");
               strSql.AppendLine("               , '" + sAutoYn + "'                                                                                 ");
               strSql.AppendLine("               , '" + sHuPayYn + "'                                                                                 ");
               strSql.AppendLine("               , '" + sPrtGb + "'                                                                                  ");
               strSql.AppendLine("               , '" + sHuPayBasisYn + "'                                                                                 ");
                strSql.AppendLine("     );                                                                                              ");






                /*
                strSql.AppendLine("");
                strSql.AppendLine("INSERT INTO ACC_PAY_CD  ");
                strSql.AppendLine("          ( GUBUN");
                strSql.AppendLine("          , DET_GB");
                strSql.AppendLine("          , PAY_CD");
                strSql.AppendLine("          , STRT_YM");
                strSql.AppendLine("          , END_YM");
                strSql.AppendLine("          , GUBUN_NM");
                strSql.AppendLine("          , DET_GB_NM");
                strSql.AppendLine("          , PAY_CD_NM");
                strSql.AppendLine("          , CALC_GB");
                strSql.AppendLine("          , CALC_RATE");
                strSql.AppendLine("          , PAY_BASIS_YN");
                strSql.AppendLine("          , PAY_MONTH");
                strSql.AppendLine("          , RMK");
                strSql.AppendLine("          , MODIFYID");
                strSql.AppendLine("          , MODIFYDT");
                strSql.AppendLine("          , GEA");
                strSql.AppendLine("          , TAX_EXCEPT_YN");
                strSql.AppendLine("          , TAX_EXCEPT_LIMIT ");
                strSql.AppendLine("          , AUTO_YN ");
                strSql.AppendLine("          , HU_PAY_YN ");
                strSql.AppendLine("          , PRT_GB ");
                strSql.AppendLine("          , HU_PAY_BASIS_YN  ) ");
                strSql.AppendLine("     VALUES");
                strSql.AppendLine("          ( '" + sGubun + "'  ");
                strSql.AppendLine("          , '" + sDetGb + "' ");
                strSql.AppendLine("          , '" + sPayCd + "' ");
                strSql.AppendLine("          , '" + sStrtYm + "'  ");
                strSql.AppendLine("          , '" + sEndYm + "' ");
                strSql.AppendLine("          , '" + sGubunNm + "'  ");
                strSql.AppendLine("          , '" + sDetGbNm + "'  ");
                strSql.AppendLine("          , '" + sPayCdNm + "'  ");
                strSql.AppendLine("          , '" + sCalcGb + "'  ");
                strSql.AppendLine("          , '" + sCalcRate + "'  ");
                strSql.AppendLine("          , '" + sPayBasisYn + "'  ");
                strSql.AppendLine("          , '" + sPayMonth + "'  ");
                strSql.AppendLine("          , '" + sRmk + "'  ");
                strSql.AppendLine("          , '" + sModifyId + "'  ");
                strSql.AppendLine("          , NOW()  ");
                strSql.AppendLine("          , '" + sGea + "'  ");
                strSql.AppendLine("          , '" + sTaxExceptYn + "'  ");
                strSql.AppendLine("          , '" + sTaxExceptLimit + "'  ");
                strSql.AppendLine("          , '" + sAutoYn + "'  ");
                strSql.AppendLine("          , '" + sHuPayYn + "'  ");
                strSql.AppendLine("          , '" + sPrtGb + "'  ");
                strSql.AppendLine("          , '" + sHuPayBasisYn + "'  ");
                strSql.AppendLine("          )  ");
                strSql.AppendLine("         ON DUPLICATE KEY UPDATE ");
                strSql.AppendLine("            END_YM           = '" + sEndYm + "'");
                strSql.AppendLine("          , GUBUN_NM         = '" + sGubunNm + "'");
                strSql.AppendLine("          , DET_GB_NM        = '" + sDetGbNm + "'");
                strSql.AppendLine("          , PAY_CD_NM        = '" + sPayCdNm + "' ");
                strSql.AppendLine("          , CALC_GB          = '" + sCalcGb + "'");
                strSql.AppendLine("          , CALC_RATE        = '" + sCalcRate + "' ");
                strSql.AppendLine("          , PAY_BASIS_YN     = '" + sPayBasisYn + "'");
                strSql.AppendLine("          , PAY_MONTH        = '" + sPayMonth + "' ");
                strSql.AppendLine("          , RMK              = '" + sRmk + "'   ");
                strSql.AppendLine("          , MODIFYID         = '" + sModifyId + "'");
                strSql.AppendLine("          , MODIFYDT         = NOW() ");
                strSql.AppendLine("          , GEA              = '" + sGea + "'");
                strSql.AppendLine("          , TAX_EXCEPT_YN    = '" + sTaxExceptYn + "' ");
                strSql.AppendLine("          , TAX_EXCEPT_LIMIT = '" + sTaxExceptLimit + "'  ");
                strSql.AppendLine("          , AUTO_YN          = '" + sAutoYn + "'  ");
                strSql.AppendLine("          , HU_PAY_YN        = '" + sHuPayYn + "'  ");
                strSql.AppendLine("          , PRT_GB           = '" + sPrtGb + "'  ");
                strSql.AppendLine("          , HU_PAY_BASIS_YN  = '" + sHuPayBasisYn + "'  ");
                */

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:ACC_PAY_CD -> GUBUN:" + sGubun + ",DET_GB:" + sDetGb + ",PAY_CD:" + sPayCd + ",STRT_YM:" + sStrtYm;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "S", this.Name, sLogRmk, cmd);

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                LkupEditGb.EditValue = sGubun;
                LkupEditDetGb.EditValue = sDetGb;

                BtnRetr_Click(null, null);

                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColPayCd, sPayCd);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
            Cursor = Cursors.Default;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            LkupEditRGb.EditValueChanged -= LkupEditRGb_EditValueChanged;
            LkupEditRGb.EditValue = null;
            LkupEditRGb.EditValueChanged += LkupEditRGb_EditValueChanged;
            LkupEditRDetGb.EditValue = null;
            TxtCd.Text = "";
            TxtCdName.Text = "";
            TxtStart.Text = "";
            TxtEnd.Text = "";
            LkupEditCalc.Text = "";
            TxtCalcRate.Text = "";
            LkupEditGea.EditValue = null;
            TxtTaxExceptLimit.Text = "";

            ChkPayBasis.EditValue = null;
            ChkTaxExcept.EditValue = null;

            ChkPayMon1.EditValue= null;
            ChkPayMon2.EditValue= null;
            ChkPayMon3.EditValue= null;
            ChkPayMon4.EditValue= null;
            ChkPayMon5.EditValue= null;
            ChkPayMon6.EditValue= null;
            ChkPayMon7.EditValue= null;
            ChkPayMon8.EditValue= null;
            ChkPayMon9.EditValue = null;
            ChkPayMon10.EditValue = null;
            ChkPayMon11.EditValue = null;
            ChkPayMon12.EditValue = null;

            Memobigo.Text = "";
            ChkAuto.EditValue = null;
            ChkHuPayBasisYn.EditValue = null;
            ChkHuPayYn.EditValue = null;
            TxtPrtGb.Text = "";
        }
        private void ChkMonCalc_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkMonCalc.Checked == true)
            {
                for (int i = 0; i <= 11; i++)
                {
                    CheckEdit[] chck = { ChkPayMon1, ChkPayMon2, ChkPayMon3, ChkPayMon4, ChkPayMon5, ChkPayMon6,
                                         ChkPayMon7, ChkPayMon8, ChkPayMon9, ChkPayMon10, ChkPayMon11, ChkPayMon12};

                    chck[i].Checked = true;
                }
            }
            else
            {
                for (int i = 0; i <= 11; i++)
                {
                    CheckEdit[] chck = { ChkPayMon1, ChkPayMon2, ChkPayMon3, ChkPayMon4, ChkPayMon5, ChkPayMon6,
                                         ChkPayMon7, ChkPayMon8, ChkPayMon9, ChkPayMon10, ChkPayMon11, ChkPayMon12};

                    chck[i].Checked = false;
                }
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AccPaymentCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
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
                BtnExcel_Click(null, null);
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccPaymentCode_FormClosed(object sender, FormClosedEventArgs e)
        {
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccPaymentCode", GridViewRetr);
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
                string sFileNM = "급여코드리스트";
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

        private void LkupEditDetGb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}