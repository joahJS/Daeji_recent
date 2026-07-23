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
using System.Diagnostics;
using System.Net;
using System.IO;
using Popbill.Fax;
using System.Data.SqlClient;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 수정일자 : 2021-02-07
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            거래처초성검색 추가 (쿼리)
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*            2. 레이아웃 전체 저장 설정
*/
namespace AccAdm
{
    public partial class AC01001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC01001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void SP01001F01_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DataTable dtAcrDr = GetLookUpData("1", "Y", "Y");
            DataTable dtAgubun = GetLookUpData("2", "Y", "Y");
            DataTable dtFindWord = GetLookUpData("3", "N", "N");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAcrDr, dtAcrDr, GridAcc, GridColAcrDr, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAgubn, dtAgubun, GridAcc, GridColAgubn, "CD", "NM", "");

            ComGrid.SetLookUpEdit(LkupFindWord, dtFindWord, "CD", "CD", "Y");
            LkupFindWord.Properties.PopulateColumns();
            LkupFindWord.Properties.Columns[1].Visible = false;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            GridAccSbj.DataSource = GetAccountSubjectInfo();

            CboFindSbj_SelectedIndexChanged(null, null);

            arrGrdView = new GridView[] { GridViewAcc, GridViewAccSbj };
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
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sUseYn = CboUseYn.EditValue?.ToString();
                string sFindSbj = CboFindSbj.EditValue?.ToString();
                string sFindWord = LkupFindWord.EditValue == null ? string.Empty : LkupFindWord.EditValue.ToString();
                
                switch (sUseYn)
                {
                    case "전체":
                        sUseYn = "ALL";
                        break;
                    case "사용중":
                        sUseYn = "Y";
                        break;
                    case "사용안함":
                        sUseYn = "N";
                        break;
                }

                switch (sFindSbj)
                {
                    case "계정코드":
                        sFindSbj = "ACCOD";
                        break;
                    case "계정명":
                        sFindSbj = "ACNAM";
                        break;
                    case "성격":
                        sFindSbj = "AGUBN";
                        break;
                    case "관계계정코드":
                        sFindSbj = "ASMCD";
                        break;
                    case "관계계정명":
                        sFindSbj = "";
                        break;
                }

                string[] sRange = GridViewAccSbj.GetFocusedRowCellValue("RANGE")?.ToString().Split('-');
                string sRange_From = sRange[0];
                string sRange_To = sRange[1];

                GridAcc.DataSource = GetAccountInfo(sRange_From, sRange_To, sUseYn, sFindSbj, sFindWord, "ALL");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        
        private void GridViewAccSbj_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
        //    if (GridViewAccSbj.RowCount < 1)
        //        return;

        //    LkupFindWord.EditValue = string.Empty;
        //    BtnRetr_Click(null, null);
        }

        private void GridViewAcc_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (GridViewAcc.RowCount < 1)
                return;

            string sAccCd = GridViewAcc.GetFocusedRowCellValue("ACCOD")?.ToString();
            GridRmk.DataSource =  GetAccountRemarkInfo(sAccCd);
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }
        
        #region[DataTable By Query]

        private DataTable GetAccountInfo(string sFrom, string sTo, string sUseYn, string sFindSbj, string sFindWord, string sGb)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.ACCOD   ");
            strSql.AppendLine("      , A.ACNAM   ");
            strSql.AppendLine("      , A.ACDSP   ");
            strSql.AppendLine("      , A.AGUBN   ");
            strSql.AppendLine("      , C.COM_NM AS AGUBN   ");
            strSql.AppendLine("      , A.ACRDR   ");
            strSql.AppendLine("      , A.ASMCD   ");
            strSql.AppendLine("      , B.ACNAM AS ASMNM ");
            strSql.AppendLine("      , A.CHGYN   ");
            strSql.AppendLine("      , A.CKCOD   ");
            strSql.AppendLine("      , A.ABGGN   ");
            strSql.AppendLine("      , A.AJNGN   ");
            strSql.AppendLine("      , A.USEYN   ");
            strSql.AppendLine("      , A.HDEC1   ");
            strSql.AppendLine("      , A.HDEC2   ");
            strSql.AppendLine("      , A.HDEC3   ");
            strSql.AppendLine("      , A.HDEC4   ");
            strSql.AppendLine("      , A.HDEC5   ");
            strSql.AppendLine("      , A.HDEC6   ");
            strSql.AppendLine("      , A.HDEC7   ");
            strSql.AppendLine("      , A.HDEC8   ");
            strSql.AppendLine("      , A.HDEC9   ");
            strSql.AppendLine("      , A.RK   ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER");
            strSql.AppendLine("      , A.CDATE                                                                                             ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER");
            strSql.AppendLine("      , A.MDATE                                                                                             ");
            strSql.AppendLine("   FROM ACMSTF A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ASMCD = B.ACCOD ");
            strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C");
            strSql.AppendLine("     ON A.AGUBN = C.COM_CD      ");
            strSql.AppendLine("     AND C.CD_GB = 'AC01001_03' ");
            strSql.AppendLine("  WHERE 1 = 1");
            if (sGb.Equals("RANGE"))
            {
                if (sFindWord.Replace(" ", "").Length == 0)
                {
                    strSql.AppendLine("    AND A.ACCOD >= '" + sFrom + "' ");
                    strSql.AppendLine("    AND A.ACCOD <= '" + sTo + "' ");
                }
            }
            if (!sUseYn.Equals("ALL"))
                strSql.AppendLine("    AND A.USEYN = '" + sUseYn + "' ");
            if (!sFindWord.Equals(""))
            {
                if (sFindSbj.Equals("ASMNM"))
                    strSql.AppendLine("    AND B." + sFindSbj + " = '" + sFindWord + "' ");
                else if (sFindSbj.Equals("AGUBN"))
                    strSql.AppendLine("    AND C.COM_NM = '" + sFindWord + "' ");
                else if (sFindSbj.Equals(""))
                {
                    strSql.AppendLine("    AND B.ACNAM = '" + sFindWord + "' ");
                }
                else
                    strSql.AppendLine("    AND A." + sFindSbj + " = '" + sFindWord + "' ");
            }
            strSql.AppendLine("ORDER BY A.ACCOD");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private DataTable GetAccountRemarkInfo(string sAccCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC1, '') AS RESULT ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '"+ sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC2, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC3, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC4, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC5, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC6, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC7, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC8, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC9, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("        ) AS X1 ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetAccountSubjectInfo()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("   FROM(           ");
            strSql.AppendLine("         SELECT A.COM_NM AS ACCNM  ");
            strSql.AppendLine("         	 , CONCAT(A.COM_SUB_CD1,'-', A.COM_SUB_CD2) AS 'RANGE'  ");
            strSql.AppendLine("           FROM COM_BASE_CD A  ");
            strSql.AppendLine("          WHERE A.CD_GB = 'AC01001_01' ");
            strSql.AppendLine("       ) X1  ");
            strSql.AppendLine("  ORDER BY X1.RANGE  ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }

        #endregion[DataTable By Query]
        
        #region[KEY_DOWN_EVENT]

        //찾을단어 단축키
        private void LkupFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        //폼 단축키
        private void AC01001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDel_Click(null, null);
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
        #endregion[KEY_DOWN_EVENT]
        
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
                strSql.AppendLine("  WHERE A.CD_GB = 'AC01001_02'");
            }
            else if (sGb.Equals("2")) //계정성격
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC01001_03'");
            }
            else if (sGb.Equals("3") || sGb.Equals("6")) //계정코드, 관계계정코드
            {
                strSql.AppendLine(" SELECT DISTINCT ACCOD AS CD");
                strSql.AppendLine("                 , '' AS NM");
                strSql.AppendLine(" FROM ACMSTF                ");
                strSql.AppendLine(" WHERE ACCOD != ''");
                strSql.AppendLine("GROUP BY ACCOD");
            }
            else if (sGb.Equals("4") || sGb.Equals("7"))//계정명, 관계계정명
            {
                strSql.AppendLine(" SELECT DISTINCT ACNAM AS CD");
                strSql.AppendLine("                 , '' AS NM");
                strSql.AppendLine(" FROM ACMSTF                ");
                strSql.AppendLine(" WHERE ACNAM != ''");
                strSql.AppendLine("GROUP BY ACNAM");
            }
            else if (sGb.Equals("5"))//계정성격
            {
                strSql.AppendLine(" SELECT COM_NM AS CD ");
                strSql.AppendLine("                 , '' AS NM");
                strSql.AppendLine(" FROM COM_BASE_CD");
                strSql.AppendLine(" WHERE CD_GB = 'AC01001_03'");
                strSql.AppendLine("GROUP BY COM_NM");
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
        

        //계정과목 현황 더블클릭 이벤트
        private void GridViewAcc_RowClick(object sender, RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                string sAccCd = GridViewAcc.GetFocusedRowCellValue("ACCOD")?.ToString();
                if (string.IsNullOrEmpty(sAccCd))
                {
                    XtraMessageBox.Show("계정코드 정보가 존재하지 않습니다.");
                    return;
                }

                AC01001F02 frm = new AC01001F02();

                frm.Owner = this;
                frm.AddModifyGb = "MOD";
                frm.AccCd = sAccCd;

                if(frm.ShowDialog() == DialogResult.OK)
                {
                }
            }
        }

        public string _ACCOD = string.Empty;
        //조회버튼 클릭 이벤트
        public void GetGridRetrData()
        {
            BtnRetr.PerformClick();
            GridViewAcc.FocusedRowHandle = GridViewAcc.LocateByDisplayText(0, GridColAcCod, _ACCOD);
        }
       
        //추가버튼 이벤트
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            AC01001F02 frm = new AC01001F02();

            frm.Owner = this;
            frm.AddModifyGb = "ADD";

            if (frm.ShowDialog() == DialogResult.OK)
            {
            }
        }

        //체크박스 사용여부 변경이벤트
        private void CboUseYn_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sAccCd = GridViewAcc.GetFocusedRowCellValue("ACCOD")?.ToString();
            if (string.IsNullOrEmpty(sAccCd))
            {
                XtraMessageBox.Show("삭제하려는 행을 정확히 선택하세요.");
                return;
            }

            if (XtraMessageBox.Show("선택하신 정보가 삭제됩니다.\r\n정말로 진행하시겠습니까?", "삭제여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DeleteAccountInfo(sAccCd);
            }
        }

        private void DeleteAccountInfo(string sAccCd)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM ACMSTF ");
                strSql.AppendLine("  WHERE ACCOD = '" + sAccCd + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제가 완료되었습니다.");

                int idx = GridViewAcc.FocusedRowHandle;
                BtnRetr.PerformClick();

                GridViewAcc.FocusedRowHandle = idx - 1;
            }
            catch (Exception ex)
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

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = "계정코드 리스트 ";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridAcc.ExportToXls(FileName);
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

        #region[GridView Row's Design]


        private void GridViewAcc_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewAcc_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewAccSbj_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRmk_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridView Row's Stripe Pattern]

        private void CboFindSbj_SelectedIndexChanged(object sender, EventArgs e)
        {
            int cboFindIdx = CboFindSbj.SelectedIndex;
            cboFindIdx = cboFindIdx + 3;

            DataTable dtFindWord = GetLookUpData(cboFindIdx.ToString(), "N", "N");

            ComGrid.SetLookUpEdit(LkupFindWord, dtFindWord, "CD", "CD", "Y");
            LkupFindWord.Properties.PopulateColumns();
            LkupFindWord.Properties.Columns[1].Visible = false;
        }

        private void AC01001F01_TextChanged(object sender, EventArgs e)
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

        private void GridViewAccSbj_RowClick(object sender, RowClickEventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sUseYn = CboUseYn.EditValue?.ToString();
                string sFindSbj = CboFindSbj.EditValue?.ToString();
                string sFindWord = LkupFindWord.EditValue == null ? string.Empty : LkupFindWord.EditValue.ToString();

                switch (sUseYn)
                {
                    case "전체":
                        sUseYn = "ALL";
                        break;
                    case "사용중":
                        sUseYn = "Y";
                        break;
                    case "사용안함":
                        sUseYn = "N";
                        break;
                }

                switch (sFindSbj)
                {
                    case "계정코드":
                        sFindSbj = "ACCOD";
                        break;
                    case "계정명":
                        sFindSbj = "ACNAM";
                        break;
                    case "성격":
                        sFindSbj = "AGUBN";
                        break;
                    case "관계계정코드":
                        sFindSbj = "ASMCD";
                        break;
                    case "관계계정명":
                        sFindSbj = "";
                        break;
                }

                string[] sRange = GridViewAccSbj.GetFocusedRowCellValue("RANGE")?.ToString().Split('-');
                string sRange_From = sRange[0];
                string sRange_To = sRange[1];

                GridAcc.DataSource = GetAccountInfo(sRange_From, sRange_To, sUseYn, sFindSbj, sFindWord, "RANGE");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void AC01001F01_Shown(object sender, EventArgs e)
        {
            LkupFindWord.Focus();
        }
    }
}