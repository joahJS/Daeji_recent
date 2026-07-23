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

/* 
 * 작성일자 : 
 * 작성자   : 
 * 비고     : 
 *            
 *            
 * ---------------------------------HISTORY---------------------------------
 * 
 * 수정일자 : 2021-05-03
 * 수정자   : 정은영
 * Reference Key : #0001
 * 수정내용 : (현업요청)
 *            1. 이미지 더블클릭시 창 닫기
 */
namespace AccAdm
{
    public partial class PD04001F03 : DevExpress.XtraEditors.XtraForm
    {
        public PD04001F03()
        {
            InitializeComponent();
        }

        public Image _IMAGE = null;
        private void PD04001F03_Load(object sender, EventArgs e)
        {
            pictureEdit1.Image = _IMAGE;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        /*
        * #0001
        */
        private void pictureEdit1_DoubleClick(object sender, EventArgs e)
        {
            BtnClose.PerformClick();
        }
    }
}