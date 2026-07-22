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
using System.Data.SqlClient;

/*
 * 작성일자 : 2020-02월 초
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-07
 * 수정자 : 고혜성
 * 수정내용 : 거래처초성검색 추가 (쿼리 참조)
 */
namespace AccAdm
{
    public partial class AC07001F02 : DevExpress.XtraEditors.XtraForm
    {
        public AC07001F02()
        {
            InitializeComponent();
        }

        public AC07001F01 F_AC07001F01;
        public string AddModifyGb;
        public DataRow DrAccInfo;

        private void AC07001F02_Load(object sender, EventArgs e)
        {
            SetControls(AddModifyGb);

            if (AddModifyGb.Equals("ADD"))
            {
                DataTable dt = SetInitDataTable(SetDataList());
                //dt.Rows[0]["CVGB"] = "거래처";
                GridRetr.DataSource = dt;
            }
            else if (AddModifyGb.Equals("MODIFY"))
            {
                GridRetr.DataSource = GetDealerAccountInfo(DrAccInfo["ACYEAR"]?.ToString(), DrAccInfo["ACCOD"]?.ToString());
                Cursor = Cursors.Default;
                DateEditFrom.EditValue = DrAccInfo["ACYEAR"]?.ToString() + "-01-01";
                BtnEditAcntCd.EditValue = DrAccInfo["ACCOD"]?.ToString();
                TxtAcntNm.EditValue = DrAccInfo["ACDSP"]?.ToString();
            }
            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);
        }

        private void SetControls(string sAddModifyGb)
        {
            if (AddModifyGb.Equals("ADD"))
            {
                DateEditFrom.EditValue = DateTime.Now.ToString("yyyy") + "-01-01";

                DateEditFrom.ReadOnly = false;                
                BtnEditAcntCd.ReadOnly = false;
                BtnEditAcntCd.EditValue = "";
                BtnCntsAdd.Enabled = true;
                TxtAcntNm.EditValue = "";

                GridRetr.DataSource = GetDealerAccountInfo("9999", "9999");

            }
            else if (AddModifyGb.Equals("MODIFY"))
            {
                DateEditFrom.ReadOnly = true;
                BtnEditAcntCd.ReadOnly = true;
                BtnCntsAdd.Enabled = false;
            }
        }

        private DataTable GetDealerAccountInfo(string sAccountYmd, string sAcntCd)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.ACYEAR ");
            strSql.AppendLine(" 	 , A.ACCOD ");
            //strSql.AppendLine("      , A.CVGB");
            strSql.AppendLine(" 	 , C.DEALER_CD AS CVCOD ");
            strSql.AppendLine(" 	 , C.DEALER_NM AS CVNAM ");
            strSql.AppendLine(" 	 , C.REP_NM AS REPNM ");
            strSql.AppendLine(" 	 , CASE WHEN B.ACRDR = '1' THEN A.ACDRJN ");
            strSql.AppendLine(" 	        WHEN B.ACRDR = '2' THEN A.ACCRJN ");
            strSql.AppendLine(" 	        END AS AMT  ");
            strSql.AppendLine(" 	 , A.CUSER ");
            strSql.AppendLine(" 	 , A.CDATE  ");
            strSql.AppendLine(" 	 , A.MUSER ");
            strSql.AppendLine(" 	 , A.MDATE ");
            strSql.AppendLine("   FROM ACJANF A ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C ");
            strSql.AppendLine("     ON A.CVCOD = C.DEALER_CD ");
            strSql.AppendLine("  WHERE A.ACYEAR = '" + sAccountYmd + "' ");
            strSql.AppendLine("    AND A.ACCOD = '" + sAcntCd + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            Cursor = Cursors.Default;
            return dt;
        }

        private void SaveCarryOverAmt(string sSaveYn)
        {
            try
            {
                string sYmd = DateEditFrom.EditValue?.ToString().Substring(0, 4);
                string sAcntCd = BtnEditAcntCd.EditValue?.ToString();

                if (string.IsNullOrEmpty(sYmd))
                {
                    XtraMessageBox.Show("회계년도를 설정하세요.");
                    return;
                }
                if (string.IsNullOrEmpty(sAcntCd))
                {
                    XtraMessageBox.Show("계정코드를 입력하세요.");
                    return;
                }
                
                if(GridViewRetr.RowCount == 0)
                {
                    XtraMessageBox.Show("거래처별 이월내역을 추가하세요.");
                    return;
                }

                DataTable dt = (DataTable)GridRetr.DataSource;

                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.ACRDR  ");
                strSql.AppendLine("   FROM ACMSTF A  ");
                strSql.AppendLine("  WHERE A.ACCOD = '" + sAcntCd + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                string sAcrDr = cmd.ExecuteScalar()?.ToString();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sCvCod = dt.Rows[i]["CVCOD"]?.ToString();
                    string sCarryOverAmt = dt.Rows[i]["AMT"]?.ToString();
                    if (string.IsNullOrEmpty(sCarryOverAmt))
                        sCarryOverAmt = "0";
                    //string sCVGB = dt.Rows[i]["CVGB"]?.ToString();
                    string sId = FmMainToolBar2.UserID;

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO ACJANF ");
                    //strSql.AppendLine("           ( ACYEAR ");
                    //strSql.AppendLine("           , ACCOD ");
                    //strSql.AppendLine("           , CVCOD ");
                    //strSql.AppendLine("           , ACDRJN ");
                    //strSql.AppendLine("           , ACCRJN ");
                    //strSql.AppendLine("           , CUSER ");
                    //strSql.AppendLine("           , CDATE ) ");
                    //strSql.AppendLine("     VALUES( '" + sYmd + "' ");
                    //strSql.AppendLine("           , '" + sAcntCd + "' ");
                    //strSql.AppendLine("           , " + sCvCod + " ");
                    //if (sAcrDr.Equals("1")) //차변
                    //{
                    //    strSql.AppendLine("           , " + sCarryOverAmt + " ");
                    //    strSql.AppendLine("           , 0 ");
                    //}
                    //else
                    //{
                    //    strSql.AppendLine("           , 0 ");
                    //    strSql.AppendLine("           , " + sCarryOverAmt + " ");
                    //}
                    //strSql.AppendLine("           , " + sId + " ");
                    //strSql.AppendLine("           , NOW() ) ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //if (sAcrDr.Equals("1")) //차변
                    //{
                    //    strSql.AppendLine("             ACDRJN = " + sCarryOverAmt + " ");
                    //    strSql.AppendLine("           , ACCRJN = 0 ");
                    //}
                    //else
                    //{
                    //    strSql.AppendLine("             ACDRJN = 0 ");
                    //    strSql.AppendLine("           , ACCRJN = " + sCarryOverAmt + " ");
                    //}
                    //strSql.AppendLine("           , MUSER = " + sId + " ");
                    //strSql.AppendLine("           , MDATE = NOW() ");
                    #endregion

                    #region 거래처구분 추가버전
                    //strSql.AppendLine("IF EXISTS(SELECT* FROM ACJANF WHERE ACYEAR = '"+ sYmd + "' AND ACCOD = '"+ sAcntCd + "'AND CVGB = '"+ sCVGB + "' AND CVCOD = '"+ sCvCod + "') ");
                    //strSql.AppendLine("   BEGIN                                                                      ");
                    //strSql.AppendLine("         UPDATE ACJANF                                                        ");
                    //strSql.AppendLine("            SET MUSER = 0                                                ");
                    //strSql.AppendLine("              , CVGB = '"+ sCVGB + "'                   ");
                    //strSql.AppendLine("              , MDATE = CONVERT(VARCHAR(19), GETDATE(), 20)                   ");
                    //if (sAcrDr.Equals("1")) //차변
                    //{
                    //    strSql.AppendLine("           , ACDRJN = " + sCarryOverAmt + " ");
                    //    strSql.AppendLine("           , ACCRJN = 0 ");
                    //}
                    //else
                    //{
                    //    strSql.AppendLine("           , ACDRJN = 0 ");
                    //    strSql.AppendLine("           , ACCRJN = " + sCarryOverAmt + " ");
                    //}
                    //strSql.AppendLine("          WHERE ACYEAR = '" + sYmd + "' AND ACCOD = '" + sAcntCd + "' AND CVGB = '" + sCVGB + "' AND CVCOD = '" + sCvCod + "'                      ");
                    //strSql.AppendLine("     END                                                                      ");
                    //strSql.AppendLine("ELSE                                                                          ");
                    //strSql.AppendLine("   BEGIN                                                                      ");
                    //strSql.AppendLine("          INSERT INTO ACJANF                                                  ");
                    //strSql.AppendLine("               (ACYEAR                                                        ");
                    //strSql.AppendLine("               , ACCOD                                                        ");
                    //strSql.AppendLine("               , CVGB                                                        ");
                    //strSql.AppendLine("               , CVCOD                                                        ");
                    //strSql.AppendLine("               , ACDRJN                                                       ");
                    //strSql.AppendLine("               , ACCRJN                                                       ");
                    //strSql.AppendLine("               , CUSER                                                        ");
                    //strSql.AppendLine("               , CDATE )                                                      ");
                    //strSql.AppendLine("         VALUES( '" + sYmd + "' ");
                    //strSql.AppendLine("               , '" + sAcntCd + "' ");
                    //strSql.AppendLine("               , '" + sCVGB + "' ");
                    //strSql.AppendLine("               , " + sCvCod + " ");
                    //if (sAcrDr.Equals("1")) //차변
                    //{
                    //    strSql.AppendLine("           , " + sCarryOverAmt + " ");
                    //    strSql.AppendLine("           , 0 ");
                    //}
                    //else
                    //{
                    //    strSql.AppendLine("           , 0 ");
                    //    strSql.AppendLine("           , " + sCarryOverAmt + " ");
                    //}
                    //strSql.AppendLine("               , " + sId + " ");
                    //strSql.AppendLine("               , CONVERT(VARCHAR(19), GETDATE(), 20))                         ");
                    //strSql.AppendLine("     END                                                                      ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM ACJANF WHERE ACYEAR = '" + sYmd + "' AND ACCOD = '" + sAcntCd + "' AND CVCOD = '" + sCvCod + "') ");
                    strSql.AppendLine("   BEGIN                                                                      ");
                    strSql.AppendLine("         UPDATE ACJANF                                                        ");
                    strSql.AppendLine("            SET MUSER = 0                                                ");
                    strSql.AppendLine("              , MDATE = CONVERT(VARCHAR(19), GETDATE(), 20)                   ");
                    if (sAcrDr.Equals("1")) //차변
                    {
                        strSql.AppendLine("           , ACDRJN = " + sCarryOverAmt + " ");
                        strSql.AppendLine("           , ACCRJN = 0 ");
                    }
                    else
                    {
                        strSql.AppendLine("           , ACDRJN = 0 ");
                        strSql.AppendLine("           , ACCRJN = " + sCarryOverAmt + " ");
                    }
                    strSql.AppendLine("          WHERE ACYEAR = '" + sYmd + "' AND ACCOD = '" + sAcntCd + "' AND CVCOD = '" + sCvCod + "'                      ");
                    strSql.AppendLine("     END                                                                      ");
                    strSql.AppendLine("ELSE                                                                          ");
                    strSql.AppendLine("   BEGIN                                                                      ");
                    strSql.AppendLine("          INSERT INTO ACJANF                                                  ");
                    strSql.AppendLine("               (ACYEAR                                                        ");
                    strSql.AppendLine("               , ACCOD                                                        ");
                    strSql.AppendLine("               , CVCOD                                                        ");
                    strSql.AppendLine("               , ACDRJN                                                       ");
                    strSql.AppendLine("               , ACCRJN                                                       ");
                    strSql.AppendLine("               , CUSER                                                        ");
                    strSql.AppendLine("               , CDATE )                                                      ");
                    strSql.AppendLine("         VALUES( '" + sYmd + "' ");
                    strSql.AppendLine("               , '" + sAcntCd + "' ");
                    strSql.AppendLine("               , " + sCvCod + " ");
                    if (sAcrDr.Equals("1")) //차변
                    {
                        strSql.AppendLine("           , " + sCarryOverAmt + " ");
                        strSql.AppendLine("           , 0 ");
                    }
                    else
                    {
                        strSql.AppendLine("           , 0 ");
                        strSql.AppendLine("           , " + sCarryOverAmt + " ");
                    }
                    strSql.AppendLine("               , " + sId + " ");
                    strSql.AppendLine("               , CONVERT(VARCHAR(19), GETDATE(), 20))                         ");
                    strSql.AppendLine("     END                                                                      ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                AC07001F01 pFrm = (AC07001F01) this.Owner;

                if(pFrm != null)
                {
                    pFrm.selectRetrData(sYmd, sAcntCd);
                }

                if (sSaveYn.Equals("SAVE"))
                    DialogResult = DialogResult.OK;
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnCntsAdd_Click(object sender, EventArgs e)
        {
            SaveCarryOverAmt("");
            SetControls("ADD");
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            SaveCarryOverAmt("SAVE");
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void AC07001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F2)
            {
                BtnCntsAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }

        #region[GridView's Design]

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridView's Design]

        #region[ButtonEdit PopupClick Event]

        public DataRow DrPopupInfo;
        private void BtnEditAcntCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            if (btnEdit.ReadOnly == true)
                return;

            AC01001F03 frm = new AC01001F03();

            frm.AccCd = btnEdit.EditValue?.ToString().Trim();
            frm.PAC07001F02 = this;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (DrPopupInfo.ItemArray.Length > 0)
                {
                    BtnEditAcntCd.EditValue = DrPopupInfo["ACCOD"];
                    TxtAcntNm.EditValue = DrPopupInfo["ACDSP"];
                    //BtnEditAcntCd.ReadOnly = true;
                    GridViewRetr.Focus();
                }
                GridViewRetr.Focus();
            }
        }

        private void BtnEditAcntCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                if (btnEdit.ReadOnly == true)
                    return;

                string sVal = btnEdit.EditValue?.ToString().Trim();
                DataTable dt = GetAccInfo(sVal);
                if(dt.Rows.Count == 1)
                {
                    btnEdit.EditValue = dt.Rows[0]["ACCOD"];
                    TxtAcntNm.EditValue = dt.Rows[0]["ACDSP"];
                    GridViewRetr.Focus();
                }
                else
                {
                    BtnEditAcntCd_ButtonClick(sender, null);
                    GridViewRetr.Focus();
                }

            }
        }

        private DataTable GetAccInfo(string sVal)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT ACCOD ");
            strSql.AppendLine("      , ACNAM ");
            strSql.AppendLine("      , ACDSP ");
            strSql.AppendLine("   FROM ACMSTF ");
            strSql.AppendLine("  WHERE ACCOD = '" + sVal + "' ");
            strSql.AppendLine("     OR ACNAM LIKE '%" + sVal + "' ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #region 거래처구분 추가 버전
        //private DataTable GetBankInfo(string sBankNam)
        //{
        //    StringBuilder strSql = new StringBuilder();

        //    strSql.Clear();
        //    strSql.AppendLine("SELECT A1.ACNT_CD             ");
        //    strSql.AppendLine("     , A1.BANK_ACNT_NO        ");
        //    strSql.AppendLine("     , A1.BANK_CD             ");
        //    strSql.AppendLine("     , B1.COM_NM AS BANK_NM   ");
        //    strSql.AppendLine("     , A1.PSSEQ               ");
        //    strSql.AppendLine("     , A1.GEJAGB              ");
        //    strSql.AppendLine("     , B2.COM_NM AS GGBNM    ");
        //    strSql.AppendLine("     , A1.ACC_CD              ");
        //    strSql.AppendLine("     , A1.RPLC_ACC_CD         ");
        //    strSql.AppendLine("     , A1.BRANCH_NM           ");
        //    strSql.AppendLine("     , A1.FINANCE_GOODS_NM    ");
        //    strSql.AppendLine("     , A1.CNTR_YMD            ");
        //    strSql.AppendLine("     , A1.CNTR_AMT            ");
        //    strSql.AppendLine("     , A1.INRST_RATE          ");
        //    strSql.AppendLine("     , A1.FRST_PAYIN_YMD      ");
        //    strSql.AppendLine("     , A1.PAYIN_AMT           ");
        //    strSql.AppendLine("     , A1.PAYIN_DD            ");
        //    strSql.AppendLine("     , A1.PAYIN_METHOD        ");
        //    strSql.AppendLine("     , A1.EXPIRE_YMD          ");
        //    strSql.AppendLine("     , A1.EXPIRE_INRST_AMT    ");
        //    strSql.AppendLine("     , A1.TRMN_YN             ");
        //    strSql.AppendLine("     , A1.TRMN_YMD            ");
        //    strSql.AppendLine("     , A1.LIQ_CPTL_YN         ");
        //    strSql.AppendLine("     , A1.RPLC_PMNT_YN        ");
        //    strSql.AppendLine("     , A1.RECEIVE_ACNT_YN     ");
        //    strSql.AppendLine("  FROM ACC_ACNT_CD A1         ");
        //    strSql.AppendLine("  LEFT JOIN COM_BASE_CD B1    ");
        //    strSql.AppendLine("    ON A1.BANK_CD = B1.COM_CD ");
        //    strSql.AppendLine("   AND B1.CD_GB = 'BANK_CD'   ");
        //    strSql.AppendLine("  LEFT JOIN COM_BASE_CD B2    ");
        //    strSql.AppendLine("    ON A1.GEJAGB = B2.COM_CD  ");
        //    strSql.AppendLine("   AND B2.CD_GB = 'GEJAGB'    ");
        //    strSql.AppendLine("  WHERE A1.BANK_ACNT_NO LIKE '%" + sBankNam + "%' ");
        //    strSql.AppendLine("     OR B1.COM_NM LIKE '%" + sBankNam + "%' ");

        //    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        //    return dt;
        //}
        #endregion

        public DataRow DrDealerInfo;
        private void RepoBtnDealerCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;

            string sDealerCd = GridViewRetr.GetFocusedRowCellValue(GridColCvNam)?.ToString();

            AC02001F03 frm = new AC02001F03();

            frm.DealerCd = sDealerCd;
            frm.P_AC07001F02 = this;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (DrDealerInfo.ItemArray.Length > 0)
                {
                    GridViewRetr.SetFocusedRowCellValue(GridColCvCod, DrDealerInfo["DEALER_CD"]);
                    GridViewRetr.SetFocusedRowCellValue(GridColCvNam, DrDealerInfo["DEALER_NM"]);
                }
            }

            #region 거래처구분 추가 버전
            //string sCvgb = GridViewRetr.GetFocusedRowCellValue("CVGB")?.ToString();

            //if (string.IsNullOrEmpty(sCvgb))
            //{
            //    XtraMessageBox.Show("거래처구분을 먼저 선택해주세요.");
            //    GridViewRetr.FocusedColumn = GridColCvGb;
            //    return;
            //}

            //if (sCvgb.Equals("거래처"))
            //{
            //    string sDealerCd = GridViewRetr.GetFocusedRowCellValue(GridColCvNam)?.ToString();

            //    AC02001F03 frm = new AC02001F03();

            //    frm.DealerCd = sDealerCd;
            //    frm.P_AC07001F02 = this;
            //    if (frm.ShowDialog() == DialogResult.OK)
            //    {
            //        if (DrDealerInfo.ItemArray.Length > 0)
            //        {
            //            GridViewRetr.SetFocusedRowCellValue(GridColCvCod, DrDealerInfo["DEALER_CD"]);
            //            GridViewRetr.SetFocusedRowCellValue(GridColCvNam, DrDealerInfo["DEALER_NM"]);
            //        }
            //    }
            //}
            //else if (sCvgb.Equals("계좌"))
            //{
            //    string sBankCd = btnEdit.EditValue?.ToString();

            //    if (!string.IsNullOrEmpty(sBankCd))
            //    {
            //        DataTable dt = GetBankInfo(sBankCd);
            //        if (dt.Rows.Count == 1)
            //        {
            //            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColCvCod, dt.Rows[0]["ACNT_CD"]);

            //            string sCvNam = dt.Rows[0]["BANK_NM"]?.ToString();
            //            if (!string.IsNullOrEmpty(dt.Rows[0]["GGBNM"]?.ToString()))
            //            {
            //                sCvNam += "(" + dt.Rows[0]["GGBNM"]?.ToString() + ")";
            //            }
            //            sCvNam += "(" + dt.Rows[0]["BANK_ACNT_NO"]?.ToString() + ")";
            //            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColCvNam, sCvNam);
            //        }
            //        else
            //        {
            //            PopUpSchAcnt frm = new PopUpSchAcnt();
            //            frm.Owner = this;
            //            frm._FINDWORD = sBankCd;
            //            frm.DataRowSendEvent += new PopUpSchAcnt.SendDataHandler(GetBankInfo);
            //            if (frm.ShowDialog() == DialogResult.OK)
            //            {
            //                GridViewRetr.FocusedColumn = GridColRepNm;
            //            }
            //        }
            //    }
            //    else
            //    {
            //        PopUpSchAcnt frm = new PopUpSchAcnt();
            //        frm.Owner = this;
            //        frm._FINDWORD = sBankCd;
            //        frm.DataRowSendEvent += new PopUpSchAcnt.SendDataHandler(GetBankInfo);
            //        if (frm.ShowDialog() == DialogResult.OK)
            //        {
            //            GridViewRetr.FocusedColumn = GridColRepNm;
            //        }
            //    }
            //}
            #endregion
        }

        #region 거래처구분 추가버전
        //private void GetBankInfo(DataRow row)
        //{
        //    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColCvCod, row["ACNT_CD"]);

        //    string sCvNam = row["BANK_NM"]?.ToString();
        //    if (!string.IsNullOrEmpty(row["GGBNM"]?.ToString()))
        //    {
        //        sCvNam += "(" + row["GGBNM"]?.ToString() + ")";
        //    }
        //    sCvNam += "(" + row["BANK_ACNT_NO"]?.ToString() + ")";
        //    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColCvNam, sCvNam);
        //}
        #endregion

        private void RepoBtnDealerCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;

                #region 거래처구분 추가버전
                //string sCvgb = GridViewRetr.GetFocusedRowCellValue("CVGB")?.ToString();

                //if (string.IsNullOrEmpty(sCvgb))
                //{
                //    XtraMessageBox.Show("거래처구분을 먼저 선택해주세요.");
                //    GridViewRetr.FocusedColumn = GridColCvGb;
                //    return;
                //}

                //if (sCvgb.Equals("거래처"))
                //{
                //    string sVal = btnEdit.EditValue?.ToString().Trim();

                //    DataTable dt = GetDealerInfo(sVal);
                //    if (dt.Rows.Count == 1)
                //    {
                //        GridViewRetr.SetFocusedRowCellValue(GridColCvCod, dt.Rows[0]["DEALER_CD"]);
                //        GridViewRetr.SetFocusedRowCellValue(GridColCvNam, dt.Rows[0]["DEALER_NM"]);
                //    }
                //    else
                //    {
                //        RepoBtnDealerCd_ButtonClick(null, null);
                //    }
                //}
                //else if (sCvgb.Equals("계좌"))
                //{
                //    string sBankCd = btnEdit.EditValue?.ToString();

                //    if (!string.IsNullOrEmpty(sBankCd))
                //    {
                //        DataTable dt = GetBankInfo(sBankCd);
                //        if (dt.Rows.Count == 1)
                //        {
                //            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColCvCod, dt.Rows[0]["ACNT_CD"]);

                //            string sCvNam = dt.Rows[0]["BANK_NM"]?.ToString();
                //            if (!string.IsNullOrEmpty(dt.Rows[0]["GGBNM"]?.ToString()))
                //            {
                //                sCvNam += "(" + dt.Rows[0]["GGBNM"]?.ToString() + ")";
                //            }
                //            sCvNam += "(" + dt.Rows[0]["BANK_ACNT_NO"]?.ToString() + ")";
                //            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColCvNam, sCvNam);
                //        }
                //        else
                //        {
                //            PopUpSchAcnt frm = new PopUpSchAcnt();
                //            frm.Owner = this;
                //            frm._FINDWORD = sBankCd;
                //            frm.DataRowSendEvent += new PopUpSchAcnt.SendDataHandler(GetBankInfo);
                //            if (frm.ShowDialog() == DialogResult.OK)
                //            {

                //            }
                //        }
                //    }
                //    else
                //    {
                //        GridViewRetr.SetFocusedRowCellValue("CVCOD", null);
                //        GridViewRetr.SetFocusedRowCellValue("CVNAM", null);
                //    }
                //}
                #endregion

                string sVal = btnEdit.EditValue?.ToString().Trim();

                DataTable dt = GetDealerInfo(sVal);
                if (dt.Rows.Count == 1)
                {
                    GridViewRetr.SetFocusedRowCellValue(GridColCvCod, dt.Rows[0]["DEALER_CD"]);
                    GridViewRetr.SetFocusedRowCellValue(GridColCvNam, dt.Rows[0]["DEALER_NM"]);
                }
                else
                {
                    RepoBtnDealerCd_ButtonClick(null, null);
                }
            }
        }

        private DataTable GetDealerInfo(string sVal)
        {
            StringBuilder strSql = new StringBuilder();

            /*
             * 수정일자 : 2021-02-07 (현업요청)
             * 수정자 : 고혜성
             * 수정내용 : 거래처초성검색 추가
             */
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT DEALER_CD ");
            strSql.AppendLine("      , DEALER_NM ");
            strSql.AppendLine("   FROM ACC_DEALER_CD ");
            strSql.AppendLine("  WHERE 1=1 ");
            if(double.TryParse(sVal, out double result))
                strSql.AppendLine("OR DEALER_CD = " + sVal);
            strSql.AppendLine("     OR DEALER_NM LIKE '%" + sVal + "%' ");
            strSql.AppendLine("     OR INITIAL_NM LIKE '%" + sVal + "%' ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #endregion[ButtonEdit PopupClick Event]

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (VerificatePreviousRowValue(GridViewRetr.GetDataRow(GridViewRetr.FocusedRowHandle)))
                return;

            GridViewRetr.AddNewRow();
            GridViewRetr.UpdateCurrentRow();

            //GridViewRetr.SetFocusedRowCellValue(GridColCvGb, "거래처");
        }

        //Row 추가 시 이전 Row Value Check
        private bool VerificatePreviousRowValue(DataRow row)
        {
            if (ComnGridFunc.VerificateCheckLastRowFocusing(GridViewRetr))
                return true;

            bool bResult = false;

            string[] sArr = new string[1];

            sArr[0] = row["CVCOD"]?.ToString();

            for (int i = 0; i < sArr.Length; i++)
            {
                if (string.IsNullOrEmpty(sArr[i]))
                    bResult = true;
            }

            return bResult;
        }

        private bool CancelCurrentRowValueCheck(DataRow row)
        {
            if (GridViewRetr.RowCount < 2)
                return true;

            /*
                GridView의 마지막 로우에만 해당 로직 적용
             */
            if (GridViewRetr.FocusedRowHandle != (GridViewRetr.RowCount - 1))
                return true;

            /*
                DB에 존재하는 데이터는 ACTRAN 테이블의 PK값이 존재하여
                전표번호가 존재할 시 ROW 취소는 할 수 없다.
             */
            string sAcCod = row["ACCOD"]?.ToString();

            if (string.IsNullOrEmpty(sAcCod))
                return false;
            else
                return true;
        }

        private DataTable SetInitDataTable(string[] sColumnNames)
        {
            DataTable dt = new DataTable();

            dt.TableName = "Table";
            for (int i = 0; i < sColumnNames.Length; i++)
            {
                dt.Columns.Add(sColumnNames[i]);
            }
            dt.Rows.Add(dt.NewRow());
            return dt;
        }

        private string[] SetDataList()
        {
            DataTable dt = GetDealerAccountInfo("9999", "9999");

            string[] sArrColumns = new string[dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sArrColumns[i] = dt.Columns[i].ColumnName;
            }

            return sArrColumns;
        }

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewRetr.RowCount < 1)
                return;

            if (e.KeyCode == Keys.Down)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (CancelCurrentRowValueCheck(GridViewRetr.GetDataRow(GridViewRetr.FocusedRowHandle)))
                    return;
                GridViewRetr.DeleteSelectedRows();
                GridViewRetr.UpdateCurrentRow();
            }
        }

        private void RepoTxtNumOnly_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txtEdit = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue("AMT", txtEdit.EditValue);
        }

        private void RepoBtnDealerCd_EditValueChanged(object sender, EventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridColCvNam, btnEdit.EditValue);
        }

        private void RepoTxtNumOnly_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                string sCvCod = GridViewRetr.GetFocusedRowCellValue(GridColCvCod)?.ToString();
                if (string.IsNullOrEmpty(sCvCod))
                {
                    XtraMessageBox.Show("거래처를 입력하세요.");
                    GridViewRetr.FocusedColumn = GridColCvNam;
                    return;
                }

                GridViewRetr.AddNewRow();
                GridViewRetr.FocusedRowHandle = GridViewRetr.RowCount - 1;
                GridViewRetr.FocusedColumn = GridColCvNam;

            }
        }

        private void RepoCboCvgb_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                ComboBoxEdit cbo = (ComboBoxEdit)sender;

                if (cbo.SelectedIndex == 0)
                    cbo.SelectedIndex = 1;
                else
                    cbo.SelectedIndex = 0;
            }
        }
    }
}