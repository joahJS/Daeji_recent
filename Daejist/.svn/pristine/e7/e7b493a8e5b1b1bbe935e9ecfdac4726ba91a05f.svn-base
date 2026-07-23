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
    public partial class MesImageRetr : DevExpress.XtraEditors.XtraForm
    {
        public MesImageRetr()
        {
            InitializeComponent();
        }

        public Dictionary<string, Image> _MESURE_IMAGE;

        private void MesImageRetr_Load(object sender, EventArgs e)
        {
            Pic1_1.Image = _MESURE_IMAGE["1_1"];
            Pic1_2.Image = _MESURE_IMAGE["1_2"];
            Pic1_3.Image = _MESURE_IMAGE["1_3"];
            Pic2_1.Image = _MESURE_IMAGE["2_1"];
            Pic2_2.Image = _MESURE_IMAGE["2_2"];
            Pic2_3.Image = _MESURE_IMAGE["2_3"];
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void Pic1_1_DoubleClick(object sender, EventArgs e)
        {
            BtnClose.PerformClick();
        }
    }
}