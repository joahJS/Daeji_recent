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
using ComLib;

namespace AccAdm
{
    public partial class SL001F01_POP02 : DevExpress.XtraEditors.XtraForm
    {
        public SL001F01_POP02()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_SL001F01";

        private void SL001F01_POP02_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            SetEditEmpty();
        }

        private void SetEditEmpty()
        {
            //DtBasdt.EditValue = DateTime.Now;
            string today1 = DateTime.Today.ToString();
            string today2 = DateTime.Now.AddMonths(-1).ToString("yyyy-MM");
            DtBasdt.EditValue = today2;
            DtGKdt.EditValue = today1.Substring(0, 10);
        }

        #region [버튼클릭이벤트]
        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sBasdt = DtBasdt.EditValue?.ToString();
            string sPaydt = DtGKdt.EditValue?.ToString();

            if (string.IsNullOrEmpty(sBasdt))
            {
                XtraMessageBox.Show("급여년월을 입력해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sPaydt))
            {
                XtraMessageBox.Show("지급일자를 입력해주세요.");
                return;
            }

            DataTable result = SaveInsaChangeData();

            if (result != null)
            {
                int sResult = int.Parse(result.Rows[0]["RESULT"]?.ToString());
                string sMsg = result.Rows[0]["MSG"]?.ToString();

                XtraMessageBox.Show(sMsg);

                if (sResult > 0)
                {
                    SL001F01 frm = (SL001F01)this.Owner;
                    //frm._SLINO = sSlino;
                    DialogResult = DialogResult.OK;
                }
                else
                {
                    DtBasdt.Focus();
                }
            }
        }

        private DataTable SaveInsaChangeData()
        {
            string sBasdt = DtBasdt.EditValue?.ToString().Substring(0, 7);
            string sDtGKdt = DtGKdt.EditValue?.ToString().Substring(0, 10);
            string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "SAVE2");
            dicParams.Add("BASYM", sBasdt);
            dicParams.Add("PAYDT", sDtGKdt);
            dicParams.Add("USER", sUser);

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SL001F01_POP02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }

        #endregion


    }
}