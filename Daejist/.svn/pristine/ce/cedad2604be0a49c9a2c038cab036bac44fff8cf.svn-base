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

namespace Daeji_MONITERING
{
    public partial class BACKGRUOND : DevExpress.XtraEditors.XtraForm
    {
        public BACKGRUOND()
        {
            InitializeComponent();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void BtnSetting_Click(object sender, EventArgs e)
        {
            Form fchk = Application.OpenForms["SETTING"];

            if (fchk == null)
            {
                SETTING frm = new SETTING();

                frm.Owner = this;
                frm.Show();
            }
        }
    }
}