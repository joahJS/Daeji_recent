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
    public partial class IN07001F02 : DevExpress.XtraEditors.XtraForm
    {
        public string sUnitPrc { get; set; }
        public string sCarryCost { get; set; }

        public IN07001F02()
        {
            InitializeComponent();
        }

        private void IN07001F02_Load(object sender, EventArgs e)
        {
            TxtUnitPrc.Focus();
        }

        private void TxtUnitPrc_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            sUnitPrc = TxtUnitPrc.EditValue?.ToString();
            sCarryCost = TxtCarryCost.EditValue?.ToString();

            if (string.IsNullOrEmpty(sUnitPrc))
            {
                XtraMessageBox.Show("적용단가를 입력하세요.");
                return;
            }
            DialogResult = DialogResult.OK;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void TxtUnitPrc_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                TxtCarryCost.Focus();
            }

            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                e.SuppressKeyPress = true;
        }

        private void UpAndDownKey_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                e.SuppressKeyPress = true;
        }
    }
}