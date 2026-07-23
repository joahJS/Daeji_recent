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

namespace AccAdm
{
    public partial class AC01001F03 : DevExpress.XtraEditors.XtraForm
    {
        public AC01001F03()
        {
            InitializeComponent();
        }

        public string AccCd { get; set; } 
        public DataRow SelectedRow { get; set; }

        public AC01001F02 PAC01001F02;
        public AC02001F02 PAC02001F02;
        public AC07001F02 PAC07001F02;
        public AC09001F01 P_AC09001F01;
        public AC07001F01 P_AC07001F01;
        public AC08001F01 P_AC08001F01;
        public AC14001F02 P_AC14001F02;

        private void AC01001F03_Load(object sender, EventArgs e)
        {
            DataTable dtAcrDr = GetLookUpData("1", "Y", "Y");
            DataTable dtAgubun = GetLookUpData("2", "Y", "Y");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAcrDr, dtAcrDr, GridAcc, GridColAcrDr, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAguBn, dtAgubun, GridAcc, GridColAguBn, "CD", "NM", "");

            CboFindWord_SelectedIndexChanged(null, null);
            FmMainToolBar2._FontSetting.SetGridView(GridViewAcc);
            if (!string.IsNullOrEmpty(AccCd))
            {
                TxtFindWord.EditValue = AccCd;
                BtnRetr_Click(null, null);
            }
        }

        private void AC01001F03_Shown(object sender, EventArgs e)
        {
            if(GridViewAcc.RowCount > 0)
            {
                GridViewAcc.Focus();
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sFindWord = TxtFindWord.EditValue?.ToString();
            DataTable dt = GetAccountInfo(sFindWord);
            GridAcc.DataSource = dt;
            if(dt.Rows.Count > 0)
            {
                GridViewAcc.Focus();
            }
        }

        private void BtnSelet_Click(object sender, EventArgs e)
        {
            if (GridViewAcc.RowCount < 1)
            {
                XtraMessageBox.Show("적용하려는 행을 정확히 선택하세요.");
                return;
            }

            SubmitRowDate();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }


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
            else if (sGb.Equals("2")) //계정성격
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC01001_03 '");
            }
            else if (sGb.Equals("3")) //계정코드
            {
                strSql.AppendLine(" SELECT ACCOD AS CD");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , ACCOD AS SEQ ");
                strSql.AppendLine("   FROM ACMSTF                ");
                //strSql.AppendLine("  WHERE ACCOD != ''");
                strSql.AppendLine("  GROUP BY ACCOD");
            }
            else if (sGb.Equals("4"))//계정명
            {
                strSql.AppendLine(" SELECT ACNAM AS CD");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , ACCOD AS SEQ ");
                strSql.AppendLine("   FROM ACMSTF                ");
                //strSql.AppendLine("  WHERE ACNAM != ''");
                strSql.AppendLine("  GROUP BY ACNAM");
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

        #region[DataTable By Query]

        private DataTable GetAccountInfo(string sFindWord)
        {
            int idx = CboFindWord.SelectedIndex;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.ACCOD   ");
            strSql.AppendLine("      , A.ACNAM   ");
            strSql.AppendLine("      , C.ACNAM AS ACTOP ");
            strSql.AppendLine("      , A.ACDSP   ");
            strSql.AppendLine("      , A.AGUBN   ");
            strSql.AppendLine("      , A.ACRDR   ");
            strSql.AppendLine("      , A.ASMCD   ");
            strSql.AppendLine("      , B.ACNAM AS ASMNM ");
            strSql.AppendLine("   FROM ACMSTF A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ASMCD = B.ACCOD ");
            strSql.AppendLine("   LEFT OUTER JOIN ACTOPF C ");
            strSql.AppendLine("     ON A.ACCOD BETWEEN C.AFROM AND C.ATO ");
            strSql.AppendLine("    AND C.SEQNO <> '0' AND A.USEYN = 'Y' ");
            strSql.AppendLine("  WHERE 1 = 1");
            if (!string.IsNullOrEmpty(sFindWord))
            {
                if (idx == 0)
                    strSql.AppendLine("    AND A.ACCOD LIKE '%" + sFindWord + "%'");
                else if (idx == 1)
                    strSql.AppendLine("    AND A.ACNAM LIKE '%" + sFindWord + "%' ");
            }
            strSql.AppendLine("    AND A.USEYN = 'Y' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion[DataTable By Query]

        #region[KeyDown Event]

        //계정코드 엔터이벤트
        private void TxtAccCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr_Click(null, null);
        }

        //계정명 엔터이벤트
        private void TxtAccNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr_Click(null, null);
        }

        //폼 전체 단축키 이벤트
        private void AC01001F03_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSelet_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        #endregion[KeyDown Event]

        private void GridViewAcc_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                SubmitRowDate();
            }
        }

        private void SubmitRowDate()
        {
            DataRow row = GridViewAcc.GetFocusedDataRow();

            if (PAC01001F02 != null)
                PAC01001F02.DrAccInfo = row;
            else if (PAC02001F02 != null)
                PAC02001F02.DrAccInfo = row;
            else if (PAC07001F02 != null)
                PAC07001F02.DrPopupInfo = row;
            else if (P_AC09001F01 != null)
                P_AC09001F01.DrPopupInfo = row;
            else if (P_AC07001F01 != null)
                P_AC07001F01.DrPopupInfo = row;
            else if (P_AC08001F01 != null)
                P_AC08001F01.DrPopupInfo = row;
            else if (P_AC14001F02 != null)
                P_AC14001F02.DrAccInfo = row;

            DialogResult = DialogResult.OK;
        }

        #region[GridView Row's Design]

        private void GridViewAcc_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewAcc_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        #endregion[GridView Row's Design]

        private void CboFindWord_SelectedIndexChanged(object sender, EventArgs e)
        {
            //int idx = CboFindWord.SelectedIndex;
            //DataTable dt = new DataTable();

            //if (idx == 0)
            //{
            //    dt = GetLookUpData("3", "Y", "Y");
            //    ComGrid.SetLookUpEdit(LkupFindWord, dt, "CD", "CD", "");
            //    LkupFindWord.Properties.PopulateColumns();
            //    LkupFindWord.Properties.Columns[1].Visible = false;
            //}
            //else if(idx == 1)
            //{
            //    dt = GetLookUpData("4", "Y", "Y");
            //    ComGrid.SetLookUpEdit(LkupFindWord, dt, "CD", "CD", "");
            //    LkupFindWord.Properties.PopulateColumns();
            //    LkupFindWord.Properties.Columns[1].Visible = false;
            //}
        }

        private void GridViewAcc_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSelet_Click(null, null);
        }
    }
}