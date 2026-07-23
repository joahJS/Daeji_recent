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
    public partial class PD02001F02 : DevExpress.XtraEditors.XtraForm
    {
        public PD02001F02()
        {
            InitializeComponent();
        }

        public PD02001F01 P_PD02001F01;
        private void PD02001F02_Load(object sender, EventArgs e)
        {
            DateEditYmd.EditValue = DateTime.Today;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            string sYmd = DateEditYmd.EditValue?.ToString().Substring(0, 10);
            if (string.IsNullOrEmpty(sYmd))
            {
                XtraMessageBox.Show("기준일자를 설정하세요.");
                return;
            }

            if (P_PD02001F01 != null)
            {
                P_PD02001F01.RST_DT = DateEditYmd.EditValue?.ToString().Substring(0, 10);
                DialogResult = DialogResult.OK;
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void DateEditYmd_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
                simpleButton1.PerformClick();
        }

        private void PD02001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                simpleButton1.PerformClick();
            else if (e.KeyCode == Keys.Escape)
                simpleButton2.PerformClick();
        }
    }
}