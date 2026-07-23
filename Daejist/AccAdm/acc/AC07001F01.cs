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
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
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
*            2. 레이아웃 전체 저장 설정
*            
*            
* 수정일자 : 2023-01-04
* 수정자   : 정은영
* 수정내용 : (현업요청)
*            1. 전체 삭제 후 저장하는걸로 변경
*/
namespace AccAdm
{
    public partial class AC07001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC07001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC07001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Now.ToString("yyyy") + "-01-01";

            DataTable dtACrDr = GetLookUpData("1", "Y", "Y"); //차대구분
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupACrDr, dtACrDr, GridAcc, GridColChaDaeGb, "CD", "NM", "");

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewAcc, GridViewCv };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            string path = ComnEtcFunc.GetLayoutPath();
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmd = DateEditFrom.EditValue?.ToString().Substring(0, 4);
            string sAcCod = BtnEditAcCod.EditValue?.ToString().Trim();

            selectRetrData(sYmd, sAcCod);
        }

        public void selectRetrData(string sYmd, string sAcCod)
        {
            GridAcc.DataSource = null;
            GridCv.DataSource = null;
            GridAcc.DataSource = GetAccountInfo(sYmd, sAcCod);
        }
        
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            AC07001F02 frm = new AC07001F02();
            frm.AddModifyGb = "ADD";
            frm.Owner = this;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                //BtnRetr_Click(null, null);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sAcntYear = GridViewAcc.GetFocusedRowCellValue(GridColAcntYear)?.ToString();
            string sAcCod = GridViewAcc.GetFocusedRowCellValue(GridColAcntCd)?.ToString();
            string sCvCod = GridViewCv.GetFocusedRowCellValue(GridCol2CvCod)?.ToString();

            string sMsg = string.Format("전표년도 : {0}" +
                "\r\n계정코드 : {1}" +
                "\r\n계정명 : {2}" +
                "\r\n거래처명 : {3}" +
                "\r\n이월금액 : {4}" +
                "해당 이월자료에 대해서 삭제를 진행하시겠습니까?"
                , GridViewAcc.GetFocusedRowCellValue(GridColAcntYear)
                , GridViewAcc.GetFocusedRowCellValue(GridColAcntCd)
                , GridViewAcc.GetFocusedRowCellValue(GridColAcntNm)
                , GridViewCv.GetFocusedRowCellValue(GridCol2CvNam)
                , GridViewCv.GetFocusedRowCellValue(GridCol2CarryOverAmt) );

            if (XtraMessageBox.Show(sMsg, "이월잔액 건 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
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
                strSql.AppendLine(" DELETE FROM ACJANF ");
                strSql.AppendLine("       WHERE ACYEAR = '" + sAcntYear + "' ");
                strSql.AppendLine("         AND ACCOD = '" + sAcCod + "' ");
                strSql.AppendLine("         AND CVCOD = '" + sCvCod + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                string sAcrDr = cmd.ExecuteScalar()?.ToString();


                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제가 완료되었습니다.");

                GridAcc.DataSource = null;
                BtnRetr.PerformClick();
                GridViewAcc.LocateByDisplayText(0, GridColAcntCd, sAcCod);
            }
            catch(Exception ex)
            {
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

            ComnEtcFunc.ExportExcelFile("거래처별 이월_", GridCv);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC07001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if(e.KeyCode == Keys.F4)
            {
                BtnDelete.PerformClick();
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

        #region[Execute By Query]

        private DataTable GetAccountInfo(string sAccountYmd, string sAcCod)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" SELECT A.ACYEAR ");
            //strSql.AppendLine(" 	 , A.ACCOD ");
            //strSql.AppendLine(" 	 , B.ACDSP ");
            //strSql.AppendLine(" 	 , B.ACRDR ");
            //strSql.AppendLine(" 	 , SUM(CASE WHEN B.ACRDR = '1' THEN IFNULL(A.ACDRJN, 0) ");
            //strSql.AppendLine(" 	            WHEN B.ACRDR = '2' THEN IFNULL(A.ACCRJN, 0) ");
            //strSql.AppendLine(" 	             END ) AS CARRY_OVER_AMT  ");
            //strSql.AppendLine("   FROM ACJANF A ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            //strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            //strSql.AppendLine("  WHERE A.ACYEAR = '" + sAccountYmd + "' ");
            //strSql.AppendLine("    AND (('" + sAcCod + "' = '' AND 1 = 1) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + sAcCod + "' <> '' AND A.ACCOD = '" + sAcCod + "'))");
            //strSql.AppendLine("  GROUP BY A.ACCOD ");
            #endregion

            strSql.AppendLine(" SELECT A.ACYEAR ");
            strSql.AppendLine(" 	 , A.ACCOD ");
            strSql.AppendLine(" 	 , B.ACDSP ");
            strSql.AppendLine(" 	 , B.ACRDR ");
            strSql.AppendLine(" 	 , SUM(CASE WHEN B.ACRDR = '1' THEN ISNULL(A.ACDRJN, 0) ");
            strSql.AppendLine(" 	            WHEN B.ACRDR = '2' THEN ISNULL(A.ACCRJN, 0) ");
            strSql.AppendLine(" 	             END ) AS CARRY_OVER_AMT  ");
            strSql.AppendLine("   FROM ACJANF A ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            strSql.AppendLine("  WHERE A.ACYEAR = '" + sAccountYmd + "' ");
            strSql.AppendLine("    AND (('" + sAcCod + "' = '' AND 1 = 1) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + sAcCod + "' <> '' AND A.ACCOD = '" + sAcCod + "'))");
            strSql.AppendLine("  GROUP BY A.ACYEAR, A.ACCOD, B.ACDSP, B.ACRDR ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetDealerAccountInfo(string sAccountYmd, string sAcntCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.ACYEAR ");
            strSql.AppendLine(" 	 , A.ACCOD ");
            strSql.AppendLine(" 	 , C.DEALER_CD AS CVCOD ");
            strSql.AppendLine(" 	 , C.DEALER_NM AS CVNAM ");
            strSql.AppendLine(" 	 , C.REP_NM ");
            strSql.AppendLine(" 	 , CASE WHEN B.ACRDR = '1' THEN A.ACDRJN ");
            strSql.AppendLine(" 	        WHEN B.ACRDR = '2' THEN A.ACCRJN ");
            strSql.AppendLine(" 	        END AS CARRY_OVER_AMT  ");
            strSql.AppendLine(" 	 , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER ");
            strSql.AppendLine(" 	 , A.CDATE  ");
            strSql.AppendLine(" 	 , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER ");
            strSql.AppendLine(" 	 , A.MDATE ");
            strSql.AppendLine("   FROM ACJANF A ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C ");
            strSql.AppendLine("     ON A.CVCOD = C.DEALER_CD ");
            strSql.AppendLine("  WHERE A.ACYEAR = '" + sAccountYmd + "' ");
            strSql.AppendLine("    AND A.ACCOD = '" + sAcntCd + "' ");

            #region 거래처구분 추가 버전
            //strSql.AppendLine("WITH TEMP1 AS(                                   ");
            //strSql.AppendLine("    SELECT A.ACYEAR                              ");
            //strSql.AppendLine("         , A.ACCOD                               ");
            //strSql.AppendLine("         , A.CVCOD AS CVCOD                  ");
            //strSql.AppendLine("         , C.DEALER_NM AS CVNAM                  ");
            //strSql.AppendLine("         , C.REP_NM                              ");
            //strSql.AppendLine("         , CASE WHEN B.ACRDR = '1' THEN A.ACDRJN ");
            //strSql.AppendLine("                WHEN B.ACRDR = '2' THEN A.ACCRJN ");
            //strSql.AppendLine("                END AS CARRY_OVER_AMT            ");
            //strSql.AppendLine("         , A.CUSER                               ");
            //strSql.AppendLine("         , A.CDATE                               ");
            //strSql.AppendLine("         , A.MUSER                               ");
            //strSql.AppendLine("         , A.MDATE                               ");
            //strSql.AppendLine("      FROM ACJANF A                              ");
            //strSql.AppendLine("      LEFT OUTER JOIN ACMSTF B                   ");
            //strSql.AppendLine("        ON A.ACCOD = B.ACCOD                     ");
            //strSql.AppendLine("      LEFT OUTER JOIN ACC_DEALER_CD C            ");
            //strSql.AppendLine("        ON A.CVCOD = C.DEALER_CD                 ");
            //strSql.AppendLine("     WHERE A.CVGB = '거래처'                     ");
            //strSql.AppendLine("     UNION ALL                                   ");
            //strSql.AppendLine("    SELECT A.ACYEAR                              ");
            //strSql.AppendLine("         , A.ACCOD                               ");
            //strSql.AppendLine("         , A.CVCOD AS CVCOD                    ");
            //strSql.AppendLine("         , CASE WHEN C3.COM_NM IS NULL THEN CONCAT(C2.COM_NM, '(', C.BANK_ACNT_NO, ')')");
            //strSql.AppendLine("                ELSE C2.COM_NM + '(' + C3.COM_NM + ')(' + C.BANK_ACNT_NO + ')' END AS CVNAM");
            //strSql.AppendLine("         , '' AS REP_NM                          ");
            //strSql.AppendLine("         , CASE WHEN B.ACRDR = '1' THEN A.ACDRJN ");
            //strSql.AppendLine("                WHEN B.ACRDR = '2' THEN A.ACCRJN ");
            //strSql.AppendLine("                END AS CARRY_OVER_AMT            ");
            //strSql.AppendLine("         , A.CUSER                               ");
            //strSql.AppendLine("         , A.CDATE                               ");
            //strSql.AppendLine("         , A.MUSER                               ");
            //strSql.AppendLine("         , A.MDATE                               ");
            //strSql.AppendLine("      FROM ACJANF A                              ");
            //strSql.AppendLine("      LEFT OUTER JOIN ACMSTF B                   ");
            //strSql.AppendLine("        ON A.ACCOD = B.ACCOD                     ");
            //strSql.AppendLine("      LEFT OUTER JOIN ACC_ACNT_CD C              ");
            //strSql.AppendLine("        ON A.CVCOD = C.ACNT_CD                   ");
            //strSql.AppendLine("      LEFT JOIN COM_BASE_CD C2                   ");
            //strSql.AppendLine("        ON C.BANK_CD = C2.COM_CD                 ");
            //strSql.AppendLine("       AND C2.CD_GB = 'BANK_CD'                  ");
            //strSql.AppendLine("      LEFT JOIN COM_BASE_CD C3                   ");
            //strSql.AppendLine("        ON C.GEJAGB = C3.COM_CD                  ");
            //strSql.AppendLine("       AND C3.CD_GB = 'GEJAGB'                   ");
            //strSql.AppendLine("     WHERE A.CVGB = '계좌'                       ");
            //strSql.AppendLine(")                                                ");
            //strSql.AppendLine(" SELECT *                                          ");
            //strSql.AppendLine("   FROM TEMP1                                     ");
            //strSql.AppendLine("  WHERE ACYEAR = '" + sAccountYmd + "' ");
            //strSql.AppendLine("    AND ACCOD = '" + sAcntCd + "' ");
            //strSql.AppendLine(" ORDER BY REPLACE(CVNAM,'(주)','')               ");
            #endregion

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #endregion[Execute By Query]
        
        #region[GetLookupData]

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("N"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1")) //차대구분
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC01001_02 '");
            }
            if (sGb.Equals("2")) 
            {
                strSql.AppendLine(" SELECT ACCOD AS CD  ");
                strSql.AppendLine("      , ACNAM AS NM  ");
                strSql.AppendLine("      , ACCOD AS SEQ ");
                strSql.AppendLine("   FROM ACMSTF A  ");
            }

            if (sParam.Equals("Y"))
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

        #endregion[GetLookupData]

        #region[GridView's Design]

        private void GridViewAcc_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewAcc_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridView's Design]

        private void GridViewAcc_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (GridViewAcc.RowCount <= 0)
                return;

            string sYmd = GridViewAcc.GetFocusedRowCellValue("ACYEAR")?.ToString();
            string sAcCod = GridViewAcc.GetFocusedRowCellValue("ACCOD")?.ToString();

            GridCv.DataSource = GetDealerAccountInfo(sYmd, sAcCod);
        }

        private void GridViewCv_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                AC07001F02 frm = new AC07001F02();
                frm.AddModifyGb = "MODIFY";
                frm.DrAccInfo = GridViewAcc.GetFocusedDataRow();
                frm.Owner = this;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    //BtnRetr_Click(null, null);
                }
            }
        }

        private void GridViewAcc_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                AC07001F02 frm = new AC07001F02();
                frm.AddModifyGb = "MODIFY";
                frm.DrAccInfo = GridViewAcc.GetFocusedDataRow();
                frm.Owner = this;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    //BtnRetr_Click(null, null);
                }
            }
        }

        private void AC07001F01_TextChanged(object sender, EventArgs e)
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

        public DataRow DrPopupInfo;
        private void BtnEditAcCod_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString();
            AC01001F03 frm = new AC01001F03();
            frm.P_AC07001F01 = this;
            frm.AccCd = sVal;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                btnEdit.EditValue = DrPopupInfo["ACCOD"];
                TxtAcNam.EditValue = DrPopupInfo["ACNAM"];
            }
        }

        private void BtnEditAcCod_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sVal = btnEdit.EditValue?.ToString();

                if (string.IsNullOrEmpty(sVal))
                {
                    TxtAcNam.EditValue = string.Empty;
                    return;
                }

                DataTable dt = GetAccInfo(sVal);
                if(dt.Rows.Count == 1)
                {
                    btnEdit.EditValue = dt.Rows[0]["ACCOD"];
                    TxtAcNam.EditValue = dt.Rows[0]["ACNAM"];
                }
                else
                {
                    BtnEditAcCod_ButtonClick(sender, null);
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
            strSql.AppendLine("   FROM ACMSTF ");
            strSql.AppendLine("  WHERE ACCOD = '" + sVal + "' ");
            strSql.AppendLine("     OR ACNAM LIKE '%" + sVal + "%' ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewCv_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                GridView view = sender as GridView;
                GridHitInfo hitInfo = view.CalcHitInfo(e.Location);
                if (hitInfo.InRow || hitInfo.InColumnPanel)
                {
                    DevExpress.Utils.DXMouseEventArgs args = DevExpress.Utils.DXMouseEventArgs.GetMouseArgs(e);
                    if (args != null) args.Handled = true;
                    view.FocusedRowHandle = hitInfo.RowHandle;
                    view.FocusedColumn = hitInfo.Column;
                    //show context menu here  
                }
            }
        }

        public string _ACYEAR = string.Empty;
        private void BtnYIW_Click(object sender, EventArgs e)
        {
            AC07001F03 frm = new AC07001F03();

            frm.Owner = this;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                SetYearMaGam();
            }
        }

        private void SetYearMaGam()
        {
            string sCashCd = ComnEtcFunc.CashCode;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");

            #region mariaDB
            //strSql.AppendLine("SELECT Z1.ACCOD");
            //strSql.AppendLine("     , Z1.CVCOD                                                                                        ");
            //strSql.AppendLine("     , CASE WHEN Z2.ACRDR = '1' THEN Z1.JAMT ELSE 0 END AS ACDRJN                                      ");
            //strSql.AppendLine("     , CASE WHEN Z2.ACRDR != '1' THEN Z1.JAMT ELSE 0 END AS ACCRJN                                     ");
            //strSql.AppendLine("  FROM(                                                                                                ");
            //strSql.AppendLine("         SELECT A1.ACCOD, A1.CVCOD, SUM(A1.JAMT) AS JAMT                                               ");
            //strSql.AppendLine("                    FROM(                                                                              ");
            //strSql.AppendLine("                           SELECT A1.ACCOD, A1.CVCOD, IFNULL(A1.ACDRJN, 0) + IFNULL(A1.ACCRJN, 0) JAMT ");
            //strSql.AppendLine("                             FROM ACJANF A1                                                            ");
            //strSql.AppendLine("                             LEFT JOIN ACC_DEALER_CD B1                                                ");
            //strSql.AppendLine("                               ON A1.CVCOD = B1.DEALER_CD                                              ");
            //strSql.AppendLine("                            WHERE A1.ACYEAR = '" + _ACYEAR + "'                                         ");
            //strSql.AppendLine("                              AND A1.ACCOD <> '" + sCashCd + "'                                            ");
            //strSql.AppendLine("                            UNION ALL                                                                  ");
            //strSql.AppendLine("                           SELECT A1.ACCOD, A1.CVCOD, SUM(CASE WHEN B1.ACRDR = '1' THEN IFNULL(A1.ACAMT, 0) - IFNULL(A1.ADAMT, 0) ELSE IFNULL(A1.ADAMT, 0) - IFNULL(ACAMT, 0) END) JAMT");
            //strSql.AppendLine("                             FROM ACTRAN A1                                                     ");
            //strSql.AppendLine("                             LEFT JOIN ACMSTF B1                                                ");
            //strSql.AppendLine("                               ON A1.ACCOD = B1.ACCOD                                           ");
            //strSql.AppendLine("                             LEFT OUTER JOIN ACC_DEALER_CD C1                                   ");
            //strSql.AppendLine("                               ON A1.CVCOD = C1.DEALER_CD                                       ");
            //strSql.AppendLine("                            WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231'");
            //strSql.AppendLine("                              AND A1.APVYN = 'Y'                                                ");
            //strSql.AppendLine("                              AND A1.ATGUB not IN('1', '2')                                     ");
            //strSql.AppendLine("                              and A1.ACCOD <> '" + sCashCd + "'                                     ");
            //strSql.AppendLine("                            GROUP BY A1.ACCOD, A1.CVCOD                                         ");
            //strSql.AppendLine("                            UNION ALL                                                           ");
            //strSql.AppendLine("                            SELECT A1.ACCOD, 6018 AS CVCOD, IFNULL(A1.ACDRJN, 0) + IFNULL(A1.ACCRJN, 0) + SUM(X1.JAMT) JAMT");
            //strSql.AppendLine("                             FROM ACJANF A1                                       ");
            //strSql.AppendLine("                             left join(                                           ");
            //strSql.AppendLine("                                     SELECT Y1.ACCOD, SUM(Y1.BAL_AMT) AS JAMT     ");
            //strSql.AppendLine("                                         FROM(                                    ");
            //strSql.AppendLine("                                             SELECT '" + sCashCd + "' AS ACCOD        ");
            //strSql.AppendLine("                                                , SUM(IFNULL(ADAMT, 0)) AS ACAMT  ");
            //strSql.AppendLine("                                                , SUM(IFNULL(ACAMT, 0)) AS ADAMT  ");
            //strSql.AppendLine("                                                , SUM(IFNULL(ADAMT, 0)) - SUM(IFNULL(ACAMT, 0)) AS BAL_AMT");
            //strSql.AppendLine("                                             FROM ACTRAN A1                                               ");
            //strSql.AppendLine("                                            WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231'");
            //strSql.AppendLine("                                              AND A1.ATGUB IN('1', '2')     ");
            //strSql.AppendLine("                                              AND A1.ACCOD <> '" + sCashCd + "' ");
            //strSql.AppendLine("                                            UNION ALL                       ");
            //strSql.AppendLine("                                           SELECT '" + sCashCd + "' AS ACCOD    ");
            //strSql.AppendLine("                                                , SUM(IFNULL(ACAMT, 0)) AS ACAMT ");
            //strSql.AppendLine("                                                , SUM(IFNULL(ADAMT, 0)) AS ADAMT ");
            //strSql.AppendLine("                                                , SUM(IFNULL(ACAMT, 0)) - SUM(IFNULL(ADAMT, 0)) AS BAL_AMT");
            //strSql.AppendLine("                                             FROM ACTRAN A1                                               ");
            //strSql.AppendLine("                                            WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231'");
            //strSql.AppendLine("                                              AND A1.ATGUB = '3'          ");
            //strSql.AppendLine("                                              AND A1.ACCOD = '" + sCashCd + "'");
            //strSql.AppendLine("                                         )Y1                              ");
            //strSql.AppendLine("                                         GROUP BY Y1.ACCOD                ");
            //strSql.AppendLine("                             )X1                                          ");
            //strSql.AppendLine("                               on A1.ACCOD = X1.ACCOD                     ");
            //strSql.AppendLine("                            WHERE A1.ACYEAR = '" + _ACYEAR + "'            ");
            //strSql.AppendLine("                              and A1.ACCOD = '" + sCashCd + "'                ");
            //strSql.AppendLine("                         ) A1                                             ");
            //strSql.AppendLine("                   GROUP BY A1.ACCOD, A1.CVCOD                            ");
            //strSql.AppendLine("       ) Z1                                                               ");
            //strSql.AppendLine("    LEFT JOIN ACMSTF Z2                                                   ");
            //strSql.AppendLine("      ON Z1.ACCOD = Z2.ACCOD                                              ");
            //strSql.AppendLine("   WHERE Z1.JAMT != 0                                                     ");
            #endregion

            #region 거래처구분 추가버전
            //strSql.AppendLine("SELECT Z1.ACCOD                                                    ");
            //strSql.AppendLine("     , Z1.CVCOD                                                    ");
            //strSql.AppendLine("     , Z1.CVGB                                                    ");
            //strSql.AppendLine("     , CASE WHEN Z2.ACRDR = '1' THEN Z1.JAMT ELSE 0 END AS ACDRJN  ");
            //strSql.AppendLine("     , CASE WHEN Z2.ACRDR != '1' THEN Z1.JAMT ELSE 0 END AS ACCRJN ");
            //strSql.AppendLine("  FROM(                                                            ");
            //strSql.AppendLine("         SELECT A1.ACCOD, A1.CVCOD, A1.CVGB, SUM(A1.JAMT) AS JAMT  ");
            //strSql.AppendLine("                    FROM(                                          ");
            //strSql.AppendLine("                           SELECT A1.ACCOD, A1.CVCOD, A1.CVGB, ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT");
            //strSql.AppendLine("                             FROM ACJANF A1          ");
            //strSql.AppendLine("                            WHERE A1.ACYEAR = '" + _ACYEAR + "'                                                   ");
            //strSql.AppendLine("                            UNION ALL                ");
            //strSql.AppendLine("                           SELECT A1.ACCOD, A1.CVCOD, A1.CVGB, SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(A1.ACAMT, 0) - ISNULL(A1.ADAMT, 0)");
            //strSql.AppendLine("                                                                  ELSE ISNULL(A1.ADAMT, 0) - ISNULL(ACAMT, 0) END) JAMT                   ");
            //strSql.AppendLine("                             FROM ACTRAN A1                   ");
            //strSql.AppendLine("                             LEFT JOIN ACMSTF B1              ");
            //strSql.AppendLine("                               ON A1.ACCOD = B1.ACCOD         ");
            //strSql.AppendLine("                             LEFT OUTER JOIN ACC_DEALER_CD C1 ");
            //strSql.AppendLine("                               ON A1.CVCOD = C1.DEALER_CD     ");
            //strSql.AppendLine("                            WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231' ");
            //strSql.AppendLine("                              AND A1.APVYN = 'Y'                ");
            //strSql.AppendLine("                            GROUP BY A1.ACCOD, A1.CVCOD, A1.CVGB");
            //strSql.AppendLine("                         ) A1                                   ");
            //strSql.AppendLine("                   GROUP BY A1.ACCOD, A1.CVCOD, A1.CVGB         ");
            //strSql.AppendLine("       ) Z1                                                     ");
            //strSql.AppendLine("    LEFT JOIN ACMSTF Z2                                         ");
            //strSql.AppendLine("      ON Z1.ACCOD = Z2.ACCOD                                    ");
            //strSql.AppendLine("   WHERE Z1.JAMT != 0                                           ");
            #endregion

            strSql.AppendLine("SELECT Z1.ACCOD                                                    ");
            strSql.AppendLine("     , Z1.CVCOD                                                    ");
            strSql.AppendLine("     , CASE WHEN Z2.ACRDR = '1' THEN Z1.JAMT ELSE 0 END AS ACDRJN  ");
            strSql.AppendLine("     , CASE WHEN Z2.ACRDR != '1' THEN Z1.JAMT ELSE 0 END AS ACCRJN ");
            strSql.AppendLine("  FROM(                                                            ");
            strSql.AppendLine("         SELECT A1.ACCOD, A1.CVCOD, SUM(A1.JAMT) AS JAMT           ");
            strSql.AppendLine("                    FROM(                                          ");
            strSql.AppendLine("                           SELECT A1.ACCOD, A1.CVCOD, ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT");
            strSql.AppendLine("                             FROM ACJANF A1               ");
            strSql.AppendLine("                             LEFT JOIN ACC_DEALER_CD B1   ");
            strSql.AppendLine("                               ON A1.CVCOD = B1.DEALER_CD ");
            strSql.AppendLine("                            WHERE A1.ACYEAR = '" + _ACYEAR + "'      ");
            strSql.AppendLine("                              and A1.ACCOD <> '"+ sCashCd + "'      ");
            strSql.AppendLine("                            UNION ALL                     ");
            strSql.AppendLine("                           SELECT A1.ACCOD, A1.CVCOD, SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(A1.ACAMT, 0) - ISNULL(A1.ADAMT, 0) ELSE ISNULL(A1.ADAMT, 0) - ISNULL(ACAMT, 0) END) JAMT");
            strSql.AppendLine("                             FROM ACTRAN A1                                  ");
            strSql.AppendLine("                             LEFT JOIN ACMSTF B1                             ");
            strSql.AppendLine("                               ON A1.ACCOD = B1.ACCOD                        ");
            strSql.AppendLine("                             LEFT OUTER JOIN ACC_DEALER_CD C1                ");
            strSql.AppendLine("                               ON A1.CVCOD = C1.DEALER_CD                    ");
            strSql.AppendLine("                            WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231' ");
            strSql.AppendLine("                              AND A1.APVYN = 'Y'                             ");
            //strSql.AppendLine("                              AND A1.ATGUB not IN('1', '2')                  ");
            strSql.AppendLine("                              and A1.ACCOD <> '" + sCashCd + "'                         ");
            strSql.AppendLine("                            GROUP BY A1.ACCOD, A1.CVCOD                      ");
            strSql.AppendLine("                            UNION ALL                                        ");
            strSql.AppendLine("                            SELECT X1.ACCOD, 6018 AS CVCOD, SUM(X1.JAMT) JAMT");
            strSql.AppendLine("                             FROM(                                           ");
            strSql.AppendLine("                                    SELECT A1.ACCOD, ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) AS JAMT");
            strSql.AppendLine("                                     FROM ACJANF A1         ");
            strSql.AppendLine("                                    WHERE A1.ACYEAR = '" + _ACYEAR + "'");
            strSql.AppendLine("                                      AND A1.ACCOD = '" + sCashCd + "' ");
            strSql.AppendLine("                                    UNION ALL               ");
            strSql.AppendLine("                                   SELECT A1.ACCOD, SUM(BAL_AMT) AS JAMT");
            strSql.AppendLine("                                     FROM(                       ");
            strSql.AppendLine("                                         SELECT '" + sCashCd + "' AS ACCOD  ");
            strSql.AppendLine("                                             , SUM(ISNULL(ADAMT, 0)) AS ACAMT  ");
            strSql.AppendLine("                                             , SUM(ISNULL(ACAMT, 0)) AS ADAMT  ");
            strSql.AppendLine("                                             , SUM(ISNULL(ADAMT, 0)) - SUM(ISNULL(ACAMT, 0)) AS BAL_AMT");
            strSql.AppendLine("                                          FROM ACTRAN A1                                 ");
            strSql.AppendLine("                                         WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231'");
            strSql.AppendLine("                                           AND A1.ATGUB IN('1', '2')                     ");
            strSql.AppendLine("                                           AND A1.ACCOD <> '" + sCashCd + "'                        ");
            strSql.AppendLine("                                         UNION ALL                                       ");
            strSql.AppendLine("                                        SELECT '" + sCashCd + "' AS ACCOD                           ");
            strSql.AppendLine("                                             , SUM(ISNULL(ACAMT, 0)) AS ACAMT            ");
            strSql.AppendLine("                                             , SUM(ISNULL(ADAMT, 0)) AS ADAMT            ");
            strSql.AppendLine("                                             , SUM(ISNULL(ACAMT, 0)) - SUM(ISNULL(ADAMT, 0)) AS BAL_AMT");
            strSql.AppendLine("                                          FROM ACTRAN A1                                 ");
            strSql.AppendLine("                                         WHERE A1.TDATE BETWEEN '" + _ACYEAR + "0101' AND '" + _ACYEAR + "1231'");
            strSql.AppendLine("                                           AND A1.ATGUB = '3'                            ");
            strSql.AppendLine("                                           AND A1.ACCOD = '" + sCashCd + "'                         ");
            strSql.AppendLine("                                       )A1           ");
            strSql.AppendLine("                                   GROUP BY A1.ACCOD ");
            strSql.AppendLine("                                ) X1                 ");
            strSql.AppendLine("                            GROUP BY X1.ACCOD        ");
            strSql.AppendLine("                         ) A1                        ");
            strSql.AppendLine("                   GROUP BY A1.ACCOD, A1.CVCOD       ");
            strSql.AppendLine("       ) Z1                                          ");
            strSql.AppendLine("    LEFT JOIN ACMSTF Z2                              ");
            strSql.AppendLine("      ON Z1.ACCOD = Z2.ACCOD                         ");
            strSql.AppendLine("   WHERE Z1.JAMT != 0                                ");
                                                                                                                        
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count>0)
            {
                try
                {
                    double.TryParse(_ACYEAR, out double dACYEAR);
                    double sACYEAR = dACYEAR + 1;

                    Cursor = Cursors.WaitCursor;

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    //#0001
                    strSql.Clear();
                    strSql.AppendLine(" DELETE FROM ACJANF     ");
                    strSql.AppendLine("  WHERE ACYEAR = '" + sACYEAR + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sACCOD = dt.Rows[i]["ACCOD"]?.ToString();
                        string sCVCOD = dt.Rows[i]["CVCOD"]?.ToString();
                        if (string.IsNullOrEmpty(sCVCOD))
                            sCVCOD = "0";
                        //string sCVGB = dt.Rows[i]["CVGB"]?.ToString();
                        string sACDRJN = dt.Rows[i]["ACDRJN"]?.ToString();
                        string sACCRJN = dt.Rows[i]["ACCRJN"]?.ToString();
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
                        //strSql.AppendLine("     VALUES( '" + sACYEAR + "' ");
                        //strSql.AppendLine("           , '" + sACCOD + "' ");
                        //strSql.AppendLine("           , " + sCVCOD + " ");
                        //strSql.AppendLine("           , " + sACDRJN + " ");
                        //strSql.AppendLine("           , " + sACCRJN + " ");
                        //strSql.AppendLine("           , " + sId + " ");
                        //strSql.AppendLine("           , NOW() ) ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             ACDRJN = " + sACDRJN + " ");
                        //strSql.AppendLine("           , ACCRJN = " + sACCRJN + " ");
                        //strSql.AppendLine("           , MUSER = " + sId + " ");
                        //strSql.AppendLine("           , MDATE = NOW() ");
                        #endregion mariaDB

                        #region 거래처구분 추가버전
                        //strSql.AppendLine("IF EXISTS(SELECT* FROM ACJANF WHERE ACYEAR = '" + sACYEAR + "' AND ACCOD = '" + sACCOD + "' AND CVCOD = " + sCVCOD + " AND CVGB = '" + sCVGB + "' )");
                        //strSql.AppendLine("   BEGIN                                                                                                            ");
                        //strSql.AppendLine("         UPDATE ACJANF                                                                                              ");
                        //strSql.AppendLine("            SET ACDRJN = " + sACDRJN + "                                                                                      ");
                        //strSql.AppendLine("              , ACCRJN = " + sACCRJN + "                                                                                         ");
                        //strSql.AppendLine("              , MUSER = " + sId + "                                                                                           ");
                        //strSql.AppendLine("              , MDATE = CONVERT(VARCHAR(19), GETDATE(), 20)                                                         ");
                        //strSql.AppendLine("          WHERE ACYEAR = '" + sACYEAR + "' AND ACCOD = '" + sACCOD + "' AND CVCOD = " + sCVCOD + " AND CVGB = '" + sCVGB + "'");
                        //strSql.AppendLine("     END                                                                                                            ");
                        //strSql.AppendLine("ELSE                                                                                                                ");
                        //strSql.AppendLine("    BEGIN                                                                                                           ");
                        //strSql.AppendLine("           INSERT INTO ACJANF                                                                                       ");
                        //strSql.AppendLine("                (ACYEAR                                                                                             ");
                        //strSql.AppendLine("                , ACCOD                                                                                             ");
                        //strSql.AppendLine("                , CVCOD                                                                                             ");
                        //strSql.AppendLine("                , CVGB");
                        //strSql.AppendLine("                , ACDRJN                                                                                            ");
                        //strSql.AppendLine("                , ACCRJN                                                                                            ");
                        //strSql.AppendLine("                , CUSER                                                                                             ");
                        //strSql.AppendLine("                , CDATE )                                                                                           ");
                        //strSql.AppendLine("          VALUES('" + sACYEAR + "'                                                                                             ");
                        //strSql.AppendLine("                , '" + sACCOD + "'                                                                                            ");
                        //strSql.AppendLine("                , " + sCVCOD + "                                                                                          ");
                        //strSql.AppendLine("                , '" + sCVGB + "'");
                        //strSql.AppendLine("                , " + sACDRJN + "                                                                                             ");
                        //strSql.AppendLine("                , " + sACCRJN + "                                                                                                 ");
                        //strSql.AppendLine("                , " + sId + "                                                                                                 ");
                        //strSql.AppendLine("                , CONVERT(VARCHAR(19), GETDATE(), 20))                                                              ");
                        //strSql.AppendLine("      END                                                                                                           ");
                        #endregion

                        strSql.AppendLine("IF EXISTS(SELECT* FROM ACJANF WHERE ACYEAR = '" + sACYEAR + "' AND ACCOD = '" + sACCOD + "' AND CVCOD = " + sCVCOD + " )");
                        strSql.AppendLine("   BEGIN                                                                                                            ");
                        strSql.AppendLine("         UPDATE ACJANF                                                                                              ");
                        strSql.AppendLine("            SET ACDRJN = " + sACDRJN + "                                                                                      ");
                        strSql.AppendLine("              , ACCRJN = " + sACCRJN + "                                                                                         ");
                        strSql.AppendLine("              , MUSER = " + sId + "                                                                                           ");
                        strSql.AppendLine("              , MDATE = CONVERT(VARCHAR(19), GETDATE(), 20)                                                         ");
                        strSql.AppendLine("          WHERE ACYEAR = '" + sACYEAR + "' AND ACCOD = '" + sACCOD + "' AND CVCOD = " + sCVCOD + "");
                        strSql.AppendLine("     END                                                                                                            ");
                        strSql.AppendLine("ELSE                                                                                                                ");
                        strSql.AppendLine("    BEGIN                                                                                                           ");
                        strSql.AppendLine("           INSERT INTO ACJANF                                                                                       ");
                        strSql.AppendLine("                (ACYEAR                                                                                             ");
                        strSql.AppendLine("                , ACCOD                                                                                             ");
                        strSql.AppendLine("                , CVCOD                                                                                             ");
                        strSql.AppendLine("                , ACDRJN                                                                                            ");
                        strSql.AppendLine("                , ACCRJN                                                                                            ");
                        strSql.AppendLine("                , CUSER                                                                                             ");
                        strSql.AppendLine("                , CDATE )                                                                                           ");
                        strSql.AppendLine("          VALUES('" + sACYEAR + "'                                                                                             ");
                        strSql.AppendLine("                , '" + sACCOD + "'                                                                                            ");
                        strSql.AppendLine("                , " + sCVCOD + "                                                                                          ");
                        strSql.AppendLine("                , " + sACDRJN + "                                                                                             ");
                        strSql.AppendLine("                , " + sACCRJN + "                                                                                                 ");
                        strSql.AppendLine("                , " + sId + "                                                                                                 ");
                        strSql.AppendLine("                , CONVERT(VARCHAR(19), GETDATE(), 20))                                                              ");
                        strSql.AppendLine("      END                                                                                                           ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show(_ACYEAR + "년도의 잔액을 " + sACYEAR + "년도로 이월 완료되었습니다.");

                    DateEditFrom.EditValue = DateTime.Parse(sACYEAR.ToString() + "-01-01");

                    BtnRetr.PerformClick();
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show(ex.Message);
                }
                finally
                {
                    _ACYEAR = string.Empty;
                    Cursor = Cursors.Default;
                }
            }
            else
            {
                XtraMessageBox.Show("전표가 없습니다. 확인해주세요");
            }
        }

        private void TxtAcNam_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateEditFrom.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevYear(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateEditFrom.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateEditFrom.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextYear(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateEditFrom.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }
    }
}