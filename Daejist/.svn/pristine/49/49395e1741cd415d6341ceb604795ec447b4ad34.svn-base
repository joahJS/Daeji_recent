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
    public partial class IN05001F03 : DevExpress.XtraEditors.XtraForm
    {
        public IN05001F03()
        {
            InitializeComponent();
        }


        public Dictionary<string, Image> _RESULT;
        private void IN05001F03_Load(object sender, EventArgs e)
        {
            RetrieveImage();
        }

        private void RetrieveImage()
        {
            if (_RESULT == null)
                return;

            int i = 1;
            foreach(KeyValuePair<string, Image> param in _RESULT)
            {
                if (i == 1)
                    Pic1.Image = param.Value;
                else if (i == 2)
                    Pic2.Image = param.Value;
                else if (i == 3)
                    Pic3.Image = param.Value;

                if (i == 3)
                    break;

                i++;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}