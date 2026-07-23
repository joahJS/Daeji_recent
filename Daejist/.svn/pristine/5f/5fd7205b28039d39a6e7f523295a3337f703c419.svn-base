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

namespace AccAdm
{
    public partial class IN06001F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN06001F01()
        {
            InitializeComponent();
        }

        public DataRow ROW_INFO = null;
        public AccMeasureCloseDev PARENT_FORM;
        public double RESULT_WEIGHT = 0;
        private void IN06001F01_Load(object sender, EventArgs e)
        {
            if(ROW_INFO != null)
            {
                string sKeraType = ROW_INFO["GB"]?.ToString();
                if (sKeraType.Equals("입고"))
                {
                    TxtDealer.EditValue = ROW_INFO["DEALER"];
                    TxtUnitPrc.EditValue = ROW_INFO["IDANGA"];
                    TxtWeight.EditValue = RESULT_WEIGHT;
                }
                else if (sKeraType.Equals("출고"))
                {
                    TxtDealer.EditValue = ROW_INFO["DEALER"];
                    TxtUnitPrc.EditValue = ROW_INFO["ODANGA"];
                    TxtWeight.EditValue = RESULT_WEIGHT;
                }
            }
            TxtRemark.Focus();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if(PARENT_FORM != null)
            {
                string sTxtRemark = TxtRemark.EditValue?.ToString();
                if (string.IsNullOrEmpty(sTxtRemark))
                {
                    XtraMessageBox.Show("수정사유를 입력해주세요.");
                    TxtRemark.Focus();
                    return;
                }

                PARENT_FORM.REMARK_RESULT = sTxtRemark;
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnByPass_Click(object sender, EventArgs e)
        {
            if (PARENT_FORM != null)
            {
                DialogResult = DialogResult.Cancel;
            }
        }

        private void IN06001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F3)
            {
                BtnApply_Click(null, null);
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnByPass_Click(null, null);
            }
        }

        private void TxtRemark_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                if (XtraMessageBox.Show("적용하시겠습니까?", "YesOrNo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    BtnApply_Click(null, null);
                }
            }
        }
    }
}