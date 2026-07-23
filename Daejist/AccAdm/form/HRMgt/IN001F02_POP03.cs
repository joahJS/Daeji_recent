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
    public partial class IN001F02_POP03 : DevExpress.XtraEditors.XtraForm
    {
        public IN001F02_POP03()
        {
            InitializeComponent();
        }

        private void IN001F02_POP03_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }
    }
}