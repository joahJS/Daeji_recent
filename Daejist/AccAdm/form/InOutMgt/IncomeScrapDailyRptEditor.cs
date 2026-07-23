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

namespace AccAdm
{
    public partial class IncomeScrapDailyRptEditor : DevExpress.XtraEditors.XtraForm
    {
        public IncomeScrapDailyRptEditor()
        {
            InitializeComponent();
        }

        public DataRow[] drArrResult { get; set; }

        private void IncomeScrapDailyRptEditor_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            DataTable dtDealer = GetLookUpData("1", "Y", "Y");
            DataTable dtGrade = GetLookUpData("2", "Y", "Y");
            DataTable dtChrg = GetLookUpData("3", "Y", "Y");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupDealer, dtDealer, GridRetr, GridColDealer, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGrade, dtGrade, GridRetr, GridColGrade, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupChrg, dtChrg, GridRetr, GridColChrgId, "CD", "NM", "");

            BtnRetr.PerformClick();
        }

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

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT CONVERT(VARCHAR,DEALER_CD)  AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.J_SERIAL AS CD");
                strSql.AppendLine("      , A.GUBUN1 AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY GUBUN1) AS SEQ");
                strSql.AppendLine("   FROM JAJAE A");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.EMP_ID AS CD ");
                strSql.AppendLine("      , A.EMP_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_NM) AS SEQ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A ");
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);
            string sDealerNm = TxtDealerNm.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT C.J_BNUM ");
            strSql.AppendLine("      , C.J_SERIAL ");
            strSql.AppendLine("      , A.J_DATE ");
            strSql.AppendLine("      , D.DEALER_CD ");
            strSql.AppendLine("      , C.GUMSUBIGO ");
            strSql.AppendLine("      , E.EMP_NM ");
            strSql.AppendLine("      , A.DANJUNG ");
            strSql.AppendLine("      , A.HALIN ");
            strSql.AppendLine("      , A.MIDANGA ");
            strSql.AppendLine("      , A.DANGA ");
            strSql.AppendLine("      , A.CKONGKEP ");
            strSql.AppendLine("      , A.IKONGKEP ");
            strSql.AppendLine("      , E.EMP_ID ");
            strSql.AppendLine("      , (A.MIDANGA - A.DANGA) AS DIFFERENCE");
            strSql.AppendLine("      , (A.IKONGKEP + A.CKONGKEP) AS TOTALAMT");
            strSql.AppendLine("   FROM INLIST A ");
            strSql.AppendLine("      , JAJAE B ");
            strSql.AppendLine("      , MESURING C ");
            strSql.AppendLine("      , ACC_DEALER_CD D");
            strSql.AppendLine("      , HR_EMP_BASIS E ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL ");
            strSql.AppendLine("    AND A.J_ID = C.IPCHULGO_MAIPID ");
            strSql.AppendLine("    AND A.J_ID1 = D.DEALER_CD ");
            strSql.AppendLine("    AND D.CHRG_ID = E.EMP_ID ");
            strSql.AppendLine("    AND A.KERATYPE = '매입' ");
            strSql.AppendLine("    AND A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.J_LOTNO = '1' ");
            strSql.AppendLine("    AND B.DAEGUBUN <> '슈레더' ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브' ");
            if(!string.IsNullOrEmpty(sDealerNm)) strSql.AppendLine("    AND D.DEALER_NM LIKE '%" + sDealerNm + "%'");
            strSql.AppendLine("  ORDER BY A.J_DATE, E.EMP_NM, D.DEALER_NM ASC");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

        }

        private void DateEditFrom_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void DateEditTo_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }
        
        private void BtnConfirm_Click(object sender, EventArgs e)
        {
            int[] iArrSelectedRow = GridViewRetr.GetSelectedRows();
            drArrResult = new DataRow[iArrSelectedRow.Length];

            for(int i = 0; i < iArrSelectedRow.Length; i++)
            {
                drArrResult[i] = GridViewRetr.GetDataRow(iArrSelectedRow[i]);
            }
            
            this.DialogResult = DialogResult.OK;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            drArrResult = null;
            this.DialogResult = DialogResult.Cancel;
        }

        private void TxtDealerNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void IncomeScrapDailyRptEditor_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                //BtnCrete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnConfirm_Click(null, null);
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
                Cursor = Cursors.WaitCursor;
                //BtnReport_Click(null, null);
                Cursor = Cursors.Default;
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}