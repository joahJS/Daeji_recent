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
    public partial class RptApplSystemP2 : DevExpress.XtraEditors.XtraForm
    {
        public RptApplSystemP2()
        {
            InitializeComponent();
        }

        private void RptApplSystemP2_Shown(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            TxtReson.Focus();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            string sReson = TxtReson.EditValue?.ToString();

            if (string.IsNullOrEmpty(sReson))
            {
                XtraMessageBox.Show("반려사유를 입력해주세요.");
                return;
            }

            RptApplSystem pFrm = (RptApplSystem) this.Owner;
            pFrm._RESON = sReson;

            DialogResult = DialogResult.OK;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}