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
    public partial class AC13001F02 : DevExpress.XtraEditors.XtraForm
    {
        public AC13001F01 AC13001F01;

        public AC13001F02()
        {
            InitializeComponent();
        }

        private void AC13001F02_Load(object sender, EventArgs e)
        {
            DataTable dtAcCod = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupAcCod, dtAcCod, "CD", "NM", "");
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            string sAcCod = LkupAcCod.EditValue?.ToString();
            string sAcNam = LkupAcCod.Text;
            if (string.IsNullOrEmpty(sAcCod))
            {
                XtraMessageBox.Show("계정코드를 선택하세요.");
                return;
            }

            if (XtraMessageBox.Show("계정코드 : " + sAcCod + "\r\n계정명 : " + sAcNam + "\r\n현재 선택한 계정코드로 선택한 비용들에 일괄처리를 진행하겠습니까?", "일괄처리여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (AC13001F01 != null)
                {
                    AC13001F01.ACCOD = sAcCod;
                    AC13001F01.ACNAM = sAcNam;
                }

                DialogResult = DialogResult.OK;
            }
        }

        private void BtnCancle_Click(object sender, EventArgs e)
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
            
            if (sGb.Equals("1")) //계정명
            {
                strSql.AppendLine(" SELECT A.ACCOD AS CD ");
                strSql.AppendLine("      , A.ACNAM AS NM");
                strSql.AppendLine("      , 1 AS SEQ");
                strSql.AppendLine("   FROM ACMSTF A ");
                strSql.AppendLine("  WHERE A.USEYN = 'Y' ");
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

        private void AC13001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
            {
            }
            else if(e.KeyCode == Keys.F3)
            {
                BtnApply_Click(null, null);
            }
        }
    }
}