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
    public partial class IN07001F03 : DevExpress.XtraEditors.XtraForm
    {
        public IN07001F03()
        {
            InitializeComponent();
        }

        public string sCVCOD1 = string.Empty;
        public string sRK1    = string.Empty;
        public string sCOST1 = string.Empty;
        public string sCVCOD2 = string.Empty;
        public string sRK2 = string.Empty;
        public string sCOST2 = string.Empty;

        private void IN07001F03_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            DataTable dtCompany = GetLookupData("1", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupCvcod1, dtCompany, "CD", "NM", "");
            ComLib.ComGrid.SetLookUpEdit(LkupCvcod2, dtCompany, "CD", "NM", "");
        }

        private DataTable GetLookupData(string sGb, string sNullYn)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD ");
                strSql.AppendLine("      , '' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT CAST(DEALER_CD AS CHAR) AS CD     ");
                strSql.AppendLine(" 	 , DEALER_NM AS NM     ");
                strSql.AppendLine("   FROM ACC_DEALER_CD");
                strSql.AppendLine("  WHERE EOB_YN = 'N'         ");
            }

            strSql.AppendLine("  ORDER BY CD ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnAppl_Click(object sender, EventArgs e)
        {
            sCVCOD1 = LkupCvcod1.EditValue?.ToString();
            sRK1 = TxtRk1.EditValue?.ToString();
            sCOST1 = TxtCost1.EditValue?.ToString();
            sCVCOD2 = LkupCvcod2.EditValue?.ToString();
            sRK2 = TxtRk2.EditValue?.ToString();
            sCOST2 = TxtCost2.EditValue?.ToString();

            DialogResult = DialogResult.OK;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void IN07001F03_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                BtnAppl.PerformClick();
        }
    }
}