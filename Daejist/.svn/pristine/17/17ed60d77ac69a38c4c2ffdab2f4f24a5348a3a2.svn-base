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
    public partial class IN11001F02 : DevExpress.XtraEditors.XtraForm
    {
        public delegate void SendDataHandler(DataRow row);
        public event SendDataHandler DataRowSendEvent;

        public IN11001F02()
        {
            InitializeComponent();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("DANGA");

            DataRow row = dt.NewRow();

            string sTxtDanGa = TxtDanGa.EditValue?.ToString();
            row["DANGA"] = sTxtDanGa;

            DataRowSendEvent(row);
            DialogResult = DialogResult.OK;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void IN11001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                BtnApply.PerformClick();
            else if (e.KeyCode == Keys.Escape) { }
        }

        private void TxtDanGa_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnApply.Focus();
        }
    }
}