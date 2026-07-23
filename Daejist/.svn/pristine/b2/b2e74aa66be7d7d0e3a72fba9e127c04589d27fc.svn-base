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
    public partial class IN001F02_POP02 : DevExpress.XtraEditors.XtraForm
    {
        public IN001F02_POP02()
        {
            InitializeComponent();
        }

        private void IN001F02_POP02_Shown(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            DateYY.EditValue = DateTime.Today;
            DateYY.Focus();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            IN001F02 pFrm = (IN001F02)this.Owner;

            pFrm._BASYY = DateYY.EditValue?.ToString()?.Substring(0, 4);

            DialogResult = DialogResult.OK;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void DateYY_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnApply.Focus();
            }
        }

        private void btnsam_Click(object sender, EventArgs e)
        {
            IN001F02_POP03 frm = new IN001F02_POP03();
            frm.Owner = this;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                
            }
        }
    }
}