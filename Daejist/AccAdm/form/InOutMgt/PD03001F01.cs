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
    public partial class PD03001F01 : DevExpress.XtraEditors.XtraForm
    {
        public PD03001F01()
        {
            InitializeComponent();
        }

        public MesMgtProgramDev _ParentForm;

        private void PD03001F01_Load(object sender, EventArgs e)
        {
            TxtIP.EditValue = ComnEtcFunc.MESURING_IP;
        }

        private void BtnConn_Click(object sender, EventArgs e)
        {
            if(TxtIP.Text.Length > 15)
            {
                XtraMessageBox.Show("IP를 정확히 입력하세요.\r\n15자 초과(기호 '.' 포함)");
                return;
            }

            _ParentForm._IP = TxtIP.EditValue.ToString().Trim();
            DialogResult = DialogResult.OK;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void TxtIP_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnConn.PerformClick();
        }
    }
}