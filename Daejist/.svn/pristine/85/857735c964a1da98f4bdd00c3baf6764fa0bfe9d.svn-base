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
    public partial class IN010F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN010F01()
        {
            InitializeComponent();
        }

        public AccFieldPSDailyRpt _ParentForm;
        private void IN010F01_Load(object sender, EventArgs e)
        {
            _ParentForm._Weight_Gubun = AccFieldPSDailyRpt.WEIGHT_GUBUN.None;
        }

        private void BtnPurc_Click(object sender, EventArgs e)
        {
            _ParentForm._Weight_Gubun = AccFieldPSDailyRpt.WEIGHT_GUBUN.Purc;
            DialogResult = DialogResult.OK;
        }

        private void BtnSale_Click(object sender, EventArgs e)
        {
            _ParentForm._Weight_Gubun = AccFieldPSDailyRpt.WEIGHT_GUBUN.Sale;
            DialogResult = DialogResult.OK;
        }

        private void BtnCancle_Click(object sender, EventArgs e)
        {
            _ParentForm._Weight_Gubun = AccFieldPSDailyRpt.WEIGHT_GUBUN.None;
            DialogResult = DialogResult.Cancel;
        }
    }
}