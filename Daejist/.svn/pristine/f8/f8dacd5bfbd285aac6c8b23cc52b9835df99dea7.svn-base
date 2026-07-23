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
    public partial class IN001F02_POP01 : DevExpress.XtraEditors.XtraForm
    {
        public IN001F02_POP01()
        {
            InitializeComponent();

        }
        private string PROCEDURE_ID = "DP_IN001F02_POP01";

        private void IN001F02_POP01_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {

           
            string sCbfam = Cbfam.EditValue.ToString();
            string sTxtCost = TxtCost.EditValue?.ToString();


            if (string.IsNullOrEmpty(sTxtCost))
            {
                XtraMessageBox.Show("월 급여액을 입력해주세요.");
                TxtCost.Focus();
                return;
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            try
            {
                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("COST", sTxtCost);
                //dicParams.Add("FAM", sCbfam);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if(dt.Rows.Count==0)
                {
                    XtraMessageBox.Show("해당 정보가 없습니다. 월 급여액을 다시 입력해주세요.");
                    TxtCost.Focus();
                    return;
                }
                
                if (sCbfam == "1")
                {
                    sCbfam = "INWN01";
                }
                else if (sCbfam == "2")
                {
                    sCbfam = "INWN02";
                }
                else if (sCbfam == "3")
                {
                    sCbfam = "INWN03";
                }
                else if (sCbfam == "4")
                {
                    sCbfam = "INWN04";
                }
                else if (sCbfam == "5")
                {
                    sCbfam = "INWN05";
                }
                else if (sCbfam == "6")
                {
                    sCbfam = "INWN06";
                }
                else if (sCbfam == "7")
                {
                    sCbfam = "INWN07";
                }
                else if (sCbfam == "8")
                {
                    sCbfam = "INWN08";
                }
                else if (sCbfam == "9")
                {
                    sCbfam = "INWN09";
                }
                else if (sCbfam == "10")
                {
                    sCbfam = "INWN10";
                }
                else if (sCbfam == "11")
                {
                    sCbfam = "INWN11";
                }
                Txttotal.EditValue = dt.Rows[0][sCbfam].ToString();


            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 오류");
            }

        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void IN001F02_POP01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
        }
    }
}