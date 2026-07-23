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
    public partial class AC07001F03 : DevExpress.XtraEditors.XtraForm
    {
        public AC07001F03()
        {
            InitializeComponent();
        }

        private void AC07001F03_Load(object sender, EventArgs e)
        {
            DateACYEAR.EditValue = DateTime.Today.AddYears(-1);
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            string sDate = DateACYEAR.EditValue?.ToString().Substring(0,4);
            double.TryParse(sDate, out double dDate);
            AC07001F01 frm = (AC07001F01)this.Owner;

            frm._ACYEAR = sDate;

            if(XtraMessageBox.Show(sDate + "년도의 잔액을 " + (dDate+1) + "년도로 이월합니다.", "이월 확인" ,MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

    }
}