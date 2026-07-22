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
    public partial class IN08001F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN08001F01()
        {
            InitializeComponent();
        }

        public MesMgtProgramDev PARENT_FORM;
        public string JUNPYO_ID = string.Empty;
        private void IN08001F01_Load(object sender, EventArgs e)
        {
            DataTable dt = GetMesureInfo(JUNPYO_ID);
            if(dt.Rows.Count > 0)
            {
                DateEditJDate.EditValue = dt.Rows[0]["J_DATE"];
                TxtKeratype.EditValue = dt.Rows[0]["KERATYPE"];
                TxtJBnum.EditValue = dt.Rows[0]["J_BNUM"];
                TxtGubun1.EditValue = dt.Rows[0]["GUBUN1"];
                TxtLoss.EditValue = dt.Rows[0]["LOSS"];
                TxtAcptWeight.EditValue = dt.Rows[0]["ACPT_WEIGHT"];
                TxtDealer.EditValue = dt.Rows[0]["DEALER"];

            }
            TxtRemark.SelectAll();
            TxtRemark.Focus();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            string sRemark = TxtRemark.EditValue?.ToString().Trim();
            if (string.IsNullOrEmpty(sRemark))
            {
                XtraMessageBox.Show("수정사유를 입력하세요.");
                TxtRemark.EditValue = sRemark;
                TxtRemark.SelectAll();
                TxtRemark.Focus();
                return;
            }

            if (PARENT_FORM != null)
            {
                PARENT_FORM.REMARK_RESULT = TxtRemark.EditValue?.ToString().Trim();
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnBypass_Click(object sender, EventArgs e)
        {
            if (PARENT_FORM != null)
            {
                DialogResult = DialogResult.OK;
            }
        }

        private void IN08001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                BtnApply_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnBypass_Click(null, null);
            }
        }

        private void TxtRemark_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (XtraMessageBox.Show("적용하시겠습니까?", "YesOrNo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        BtnApply_Click(null, null);
                    }
                }
            }
        }

        private DataTable GetMesureInfo(string sJunpyoId)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.J_DATE ");
            strSql.AppendLine("      , A.KERATYPE ");
            strSql.AppendLine("      , A.J_BNUM ");
            strSql.AppendLine("      , A.GUBUN1 ");
            strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END AS LOSS ");
            strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END AS ACPT_WEIGHT ");
            strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER ");
            strSql.AppendLine("   FROM MESURING A   ");
            strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");
            
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

    }
}