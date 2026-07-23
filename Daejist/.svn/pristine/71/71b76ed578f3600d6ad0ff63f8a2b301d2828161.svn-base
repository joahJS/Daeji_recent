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
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class GE003F01 : DevExpress.XtraEditors.XtraForm
    {
        public GE003F01()
        {
            InitializeComponent();
        }

        public string _EMPID = string.Empty;
        public string AddModifyGb { get; set; }
        public DataRow _Row;

        public delegate void SendDataHandler(Dictionary<string, string> dicParams);
        public event SendDataHandler DataRowSendEvent;

        private void GE003F01_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            if (AddModifyGb.Equals("ADD"))
            {
                DateBasdt.EditValue = DateTime.Now.ToString("yyyy-MM-dd");
                DateBasdt.ReadOnly = false;
            }
            else if (AddModifyGb.Equals("MOD"))
            {
                RadiYNGUB.EditValue = _Row["YNGUB"];
                DateBasdt.EditValue = _Row["BASDT"];
                TxtYncnt.EditValue = _Row["YNCNT"];
                MemoRk.EditValue = _Row["RK"];

                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                DateBasdt.ReadOnly = true;
            }
            else if (AddModifyGb.Equals("ALL"))
            {
                DateBasdt.EditValue = DateTime.Now.ToString("yyyy-MM-dd");
                DateBasdt.ReadOnly = false;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }

        private void GE003F01_Shown(object sender, EventArgs e)
        {
            DateBasdt.Focus();
        }

        private void SaveYnch_m()
        {
            try
            {
                string sBASDT = DateBasdt.EditValue?.ToString().Substring(0,10);
                string sYNCNT = TxtYncnt.EditValue?.ToString();
                string sYNGUB = RadiYNGUB.EditValue?.ToString();
                string sRK = MemoRk.EditValue?.ToString();
                string sUSER = FmMainToolBar2.drUser["USRCD"]?.ToString();

                double dYNCNT = 0;

                if (string.IsNullOrEmpty(sYNCNT))
                {
                    XtraMessageBox.Show("연차 발생수량을 입력해주세요.");
                    TxtYncnt.Focus();
                    return;
                }

                double.TryParse(sYNCNT, out dYNCNT);

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" IF EXISTS(SELECT * FROM YNCH_M WHERE EMPID = '"+_EMPID+"' AND BASDT = '"+ sBASDT + "')");
                strSql.AppendLine("     BEGIN                                                     ");
                strSql.AppendLine("             UPDATE YNCH_M                                     ");
                strSql.AppendLine("                SET YNCNT = "+ dYNCNT + "                                  ");
                strSql.AppendLine("                  , YNGUB = '" + sYNGUB + "'");
	            strSql.AppendLine("                  , RK = '"+ sRK + "'                                    ");
	            strSql.AppendLine("                  , MUSER = "+ sUSER + "                                  ");
	            strSql.AppendLine("                  , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20)");
                strSql.AppendLine("             WHERE EMPID = '"+ _EMPID + "' AND BASDT = '"+ sBASDT + "'                   ");
                strSql.AppendLine("         END                                                   ");
                strSql.AppendLine(" ELSE                                                          ");
                strSql.AppendLine("     BEGIN                                                     ");
                strSql.AppendLine("             INSERT INTO YNCH_M( EMPID                         ");
                strSql.AppendLine("                               , BASDT                         ");
                strSql.AppendLine("                               , YNGUB");
                strSql.AppendLine("                               , YNCNT                         ");
                strSql.AppendLine("                               , RK                            ");
                strSql.AppendLine("                               , CUSER )                       ");
                strSql.AppendLine("                         VALUES( '"+ _EMPID + "'                        ");
                strSql.AppendLine("                               , '"+ sBASDT + "'                       ");
                strSql.AppendLine("                               , '" + sYNGUB + "'");
                strSql.AppendLine("                               , "+ dYNCNT + "                         ");
                strSql.AppendLine("                               , '"+ sRK + "'                            ");
                strSql.AppendLine("                               , "+ sUSER + ")                        ");
                strSql.AppendLine("         END                                                   ");
                                                                                                  
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                GE003F00 pFrm = this.Owner as GE003F00;
                pFrm._EMPID = _EMPID;
                pFrm._BASDT = sBASDT;
                XtraMessageBox.Show("저장이 완료되었습니다.");
                pFrm.SetGridViewRetr1Data();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnSaveMt_Click(object sender, EventArgs e)
        {
            SaveYnch_m();
            resetEdit();
            DateBasdt.Focus();
        }

        private void resetEdit()
        {
            RadiYNGUB.EditValue = "0";
            DateBasdt.EditValue = null;
            TxtYncnt.EditValue = 0;
            MemoRk.EditValue = null;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (AddModifyGb.Equals("ALL"))
            {
                string sBASDT = DateBasdt.EditValue?.ToString().Substring(0, 10);
                string sYNCNT = TxtYncnt.EditValue?.ToString();
                string sYNGUB = RadiYNGUB.EditValue?.ToString();
                string sRK = MemoRk.EditValue?.ToString();

                if (string.IsNullOrEmpty(sYNCNT))
                {
                    XtraMessageBox.Show("연차 발생수량을 입력해주세요.");
                    TxtYncnt.Focus();
                    return;
                }


                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("BASDT",sBASDT );
                dicParams.Add("YNCNT",sYNCNT );
                dicParams.Add("YNGUB",sYNGUB );
                dicParams.Add("RK",sRK);

                DataRowSendEvent(dicParams);
                DialogResult = DialogResult.OK;
            }
            else
            {
                SaveYnch_m();
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void GE003F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                BtnSaveMt.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }
    }
}