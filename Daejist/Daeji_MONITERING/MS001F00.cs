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
using DevExpress.XtraEditors.ViewInfo;
using DevExpress.Utils.Drawing;

namespace Daeji_MONITERING
{
    public partial class MS001F00 : DevExpress.XtraEditors.XtraForm
    {
        public MS001F00()
        {
            InitializeComponent();
        }

        public string _MSG = string.Empty;

        private void MS001F00_Load(object sender, EventArgs e)
        {
            LoadMsg();
        }

        public void LoadMsg()
        {
            CSafeSetMsg(memoEdit1, _MSG);
        }

        //메세지 값변경
        delegate void CrossThreadSafetySetTextMsg(MemoEdit ctl, string str);
        private void CSafeSetMsg(MemoEdit ctl, string str)
        {
            if (ctl.InvokeRequired)
            {
                ctl.Invoke(new CrossThreadSafetySetTextMsg(CSafeSetMsg), ctl, str);
            }
            else
            {
                ctl.Text = str;
            }
        }

        //백그라운드 폼닫기
        delegate void CrossThreadSafetySetForm(Form frm);
        private void CSafeSetForm(Form frm)
        {
            if (frm.InvokeRequired)
            {
                frm.Invoke(new CrossThreadSafetySetForm(CSafeSetForm), frm);
            }
            else
            {
                frm.Close();
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Form mFrm = Application.OpenForms["DJMONITERING"];

            if(mFrm != null)
            {
                DJMONITERING moni = (DJMONITERING)mFrm;
                moni._ALMDISPALY = DJMONITERING.ALMDISPALY.OFF;
            }

            Form fchk = Application.OpenForms["BACKGRUOND"];

            if (fchk != null)
                fchk.Close();

            this.Close();
        }
    }
}