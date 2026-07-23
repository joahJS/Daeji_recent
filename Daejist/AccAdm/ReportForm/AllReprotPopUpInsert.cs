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
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class AllReprotPopUpInsert : DevExpress.XtraEditors.XtraForm
    {
        private string PROCEDURE_ID = "DP_AllReport";

        public AllReprotPopUpInsert()
        {
            InitializeComponent();
        }

       
        private void AllReprotPopUpInsert_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            DateTime today = DateTime.Now.Date;
            Dtyyyydt.EditValue = today.AddYears(0);
        }


        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                
                string sFrom = Dtyyyydt.EditValue?.ToString().Substring(0, 4);
                //string s1 = Dtyyyydt.EditValue?.ToString().Substring(0, 10);
                DateTime sd1 = (DateTime)Dtyyyydt.EditValue;
                int sd2 = sd1.AddYears(-1).Year;
                string sFromx = sd2.ToString();
                //string srFrom = Dtyyyydt.EditValue?.ToString().Substring(0, 4);



                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("SDATE", sFromx);

                DataTable dtC = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dtC.Rows.Count==0)
                {
                    XtraMessageBox.Show("'" + sFromx + "'년도 마감자료가 없습니다.");
                    return;
                }

                
                if (XtraMessageBox.Show("'" + sFrom + "'년도 마감을 하시겠습니까?", "마감 진행 여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                string sxyear = string.Concat(sFrom, "-01-01");
                string sxFrom = string.Concat(sFromx, "-01-01");
                string sxTo = string.Concat(sFromx, "-12-31");
                string x1;
                string x2;
                string x3;
                string x4;
                string x5;
                string x6;
                string x7;
                string x8;
                string x9;
                
                
                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("FDATE", sxFrom);
                dicParams.Add("TDATE", sxTo);
                
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                
                if (dt==null)
                {
                    x1 = "0";
                    x2 = "0";
                    x3 = "0";
                    x4 = "0";
                    x5 = "0";
                    x6 = "0";
                    x7 = "0";
                    x8 = "0";
                    x9 = "0";
                }
                else
                {
                    x1 = dt.Rows[0]["A1"]?.ToString();
                    x2 = dt.Rows[0]["B1"]?.ToString();
                    x3 = dt.Rows[0]["C1"]?.ToString();
                    x4 = dt.Rows[0]["A2"]?.ToString();
                    x5 = dt.Rows[0]["B2"]?.ToString();
                    x6 = dt.Rows[0]["C2"]?.ToString();
                    x7 = dt.Rows[0]["A3"]?.ToString();
                    x8 = dt.Rows[0]["B3"]?.ToString();
                    x9 = dt.Rows[0]["C3"]?.ToString();
                }
                
                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");
                dicParams.Add("GDATE", sxyear);
                dicParams.Add("SURDW", x1); 
                dicParams.Add("SURDD", x2); 
                dicParams.Add("SURDC", x3); 
                dicParams.Add("GOCULW", x4);
                dicParams.Add("GOCULD", x5);
                dicParams.Add("GOCULC", x6);
                dicParams.Add("TOTALW", x7);
                dicParams.Add("TOTALD", x8);
                dicParams.Add("TOTALC", x9);
                
                DataTable dtZ = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                //UserMgt fm = this.Owner as UserMgt;
                //fm._USRID = sUserID;
                DialogResult = DialogResult.OK;

            }
            catch (Exception ex)
            {

            }
        }
        private void AllReprotPopUpInsert_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }
        
    }
}