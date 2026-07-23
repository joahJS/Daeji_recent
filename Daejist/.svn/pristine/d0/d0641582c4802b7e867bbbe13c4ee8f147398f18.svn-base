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
    public partial class PD01001F03 : DevExpress.XtraEditors.XtraForm
    {
        public PD01001F03()
        {
            InitializeComponent();
        }

        public ProdMgtReport P_ProdMgtReport;
        public string MAKENO;
        public string MAKENO_LN;

        private void PD01001F03_Load(object sender, EventArgs e)
        {
            DateEditYmd.EditValue = DateTime.Now;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT B.EMP_NM ");
            strSql.AppendLine("   FROM MAKE_M A  ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B  ");
            strSql.AppendLine("     ON A.MUSER_ID = B.EMP_ID ");
            strSql.AppendLine("  WHERE MAKENO = " + MAKENO + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmd = DateEditYmd.EditValue?.ToString().Substring(0, 10);
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("J_DATE", DateEditYmd.EditValue?.ToString().Substring(0, 10));

            GridRetr.DataSource = GetInfo(dicParams);
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if(GridViewRetr.GetSelectedRows().Length == 0)
            {
                XtraMessageBox.Show("적용하려는 데이터에 체크를 진행하세요.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("선택하신 {0}건에 대하여 검수내역 적용을 진행하시겠습니까?", GridViewRetr.GetSelectedRows().Length), "검수내역 저장여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            if(P_ProdMgtReport != null)
            {
                P_ProdMgtReport.RST_GRID_VIEW = GridViewRetr;
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private DataTable GetInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            //strSql.AppendLine(" SELECT * ");
            //strSql.AppendLine("   FROM (");
            strSql.AppendLine("          SELECT A.J_DATE ");
            strSql.AppendLine("               , A.J_BNUM AS CARNO ");
            strSql.AppendLine("               , A.MAIPCHERID AS DEALER_CD ");
            strSql.AppendLine("               , B.IDT_NO  ");
            strSql.AppendLine("               , B.DEALER_NM AS DEALER_NM ");
            strSql.AppendLine("               , SUBSTRING(A.SECONDTIME,12,5) AS FIRST_TIME ");
            strSql.AppendLine("               , SUBSTRING(A.FIRSTTIME,12,5) AS SECOND_TIME ");
            strSql.AppendLine("               , A.J_SERIAL AS GRADE_CD ");
            strSql.AppendLine("               , A.GUBUN1 AS GRADE_NM ");
            strSql.AppendLine("               , A.WEIGHT AS DJ_WEIGHT ");
            strSql.AppendLine("               , A.IWEIGHT AS IN_WEIGHT ");
            strSql.AppendLine("               , A.ICHAGAM AS LOSS ");
            strSql.AppendLine("               , A.GUMSUBIGO AS ISPT_NOTE ");
            strSql.AppendLine("               , A.GUMSU_SERIAL AS ISPT_CD ");
            strSql.AppendLine("               , C.EMP_NM AS ISPT_NM ");
            strSql.AppendLine("            FROM MESURING A");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD B");
            strSql.AppendLine("              ON A.MAIPCHERID = B.DEALER_CD");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS C ");
            strSql.AppendLine("              ON A.GUMSU_SERIAL = C.EMP_ID");
            strSql.AppendLine("           WHERE A.KERATYPE = '입고'");
            strSql.AppendLine("             AND A.J_DATE = '" + dicParams["J_DATE"] + "' ");
            strSql.AppendLine("           ORDER BY J_DATE DESC");
            //strSql.AppendLine("        ) X1");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (GridViewRetr.IsRowSelected(e.RowHandle))
                GridViewRetr.UnselectRow(e.RowHandle);
            else
                GridViewRetr.SelectRow(e.RowHandle);
        }

        private void PD01001F03_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
            {
                BtnClose_Click(null, null);
            }
            else if(e.KeyCode == Keys.F3)
            {
                BtnApply_Click(null, null);
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void DateEditYmd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}