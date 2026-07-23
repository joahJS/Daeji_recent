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
    public partial class SL001F00_POP01 : DevExpress.XtraEditors.XtraForm
    {
        public SL001F00_POP01()
        {
            InitializeComponent();
        }

        public DataRow _ModRow;
        string _sSlino = string.Empty;
        public DataRow _InsaInfo;
        public string Dtstdt2;

        public enum ModeGubun { Add, Modi }
        public ModeGubun _ModeGubun = ModeGubun.Add;
        public enum ChgGubun { Chgbf, Chgaf }
        public ChgGubun _ChgGubun = ChgGubun.Chgbf;

        private string PROCEDURE_ID = "DP_SL001F00";

        private void SL001F00_POP01_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.gp_SetColorFocused(layoutControl3);
            string today1 = DateTime.Today.ToString();
            DateBASYM.EditValue = Dtstdt2;

            ComnEtcFunc.SetBoundLookUp(LkupEmpid, "HR_EMP_BASIS", "", "");
            ComnEtcFunc.SetBoundLookUp(LkupDept, "ACC_DEPT_CD", "DEPT_CD", "DEPT_NM");
            ComnEtcFunc.SetBoundLookUp(LkupJikgub, "COM_BASE_CD", "GRADE_CD", "");

            if (_ModeGubun == ModeGubun.Add)
            {
            }
            else if (_ModeGubun == ModeGubun.Modi)
            {
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                if (_ModRow != null)
                {
                    LkupEmpid.EditValue = _ModRow["EMPID"];
                    DateBASYM.EditValue = _ModRow["BASYM"];
                    LkupDept.EditValue = _ModRow["DEPT_CD"];
                    LkupJikgub.EditValue = _ModRow["GRADE_CD"];

                    //기본
                    TxtGibon.EditValue = _ModRow["GIBON"];
                    TxtSigub.EditValue = _ModRow["SIGUB"];
                    TxtYjsigub.EditValue = _ModRow["YJSIGUB"];

                    //변동수당
                    tx_CSUDN1.EditValue = _ModRow["CSUDN1"];
                    tx_CSUDN6.EditValue = _ModRow["CSUDN6"];
                    tx_CSUDN7.EditValue = _ModRow["CSUDN7"];

                    //변동공제
                    tx_CGONJ1.EditValue = _ModRow["CGONJ1"];
                    tx_CGONJ2.EditValue = _ModRow["CGONJ2"];
                    tx_CGONJ3.EditValue = _ModRow["CGONJ3"];
                    tx_CGONJ4.EditValue = _ModRow["CGONJ4"];
                    tx_CGONJ5.EditValue = _ModRow["CGONJ5"];
                    tx_CGONJ6.EditValue = _ModRow["CGONJ6"];

                    MemRk.EditValue = _ModRow["RK"];

                    _ModRow = null;

                }
                DateBASYM.ReadOnly = true;
                LkupEmpid.ReadOnly = true;
            }
        }

        private void SetEditEmpty()
        {
            string today1 = DateTime.Today.ToString();
            DateBASYM.EditValue = today1.Substring(0, 7);

            LkupEmpid.EditValue = null;
            LkupDept.EditValue = null;
            LkupJikgub.EditValue = null;

            //기본
            TxtGibon.EditValue = 0;
            TxtSigub.EditValue = 0;
            TxtYjsigub.EditValue = 0;

            tx_CSUDN1.EditValue = 0;
            tx_CSUDN6.EditValue = 0;
            tx_CSUDN7.EditValue = 0;

            tx_CGONJ1.EditValue = 0;
            tx_CGONJ2.EditValue = 0;
            tx_CGONJ3.EditValue = 0;
            tx_CGONJ4.EditValue = 0;
            tx_CGONJ5.EditValue = 0;
            tx_CGONJ6.EditValue = 0;

            MemRk.EditValue = null;
        }

        #region [버튼클릭이벤트]
        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sEmpid = LkupEmpid.EditValue?.ToString();
            string sBasdt = DateBASYM.EditValue?.ToString();
            if (string.IsNullOrEmpty(sEmpid))
            {
                XtraMessageBox.Show("사원을 선택해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sBasdt))
            {
                XtraMessageBox.Show("일자를 입력해주세요.");
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
                    SL001F00 frm = (SL001F00)this.Owner;
                    DialogResult = DialogResult.OK;
                }
                else
                {
                    DateBASYM.Focus();
                }
            }
        }

        private DataTable SaveInsaChangeData()
        {
            string sEmpid = LkupEmpid.EditValue?.ToString();
            string sBasym = DateBASYM.EditValue?.ToString().Substring(0,7);

            string sCSUDN1 = tx_CSUDN1.EditValue?.ToString();
            string sCSUDN6 = tx_CSUDN6.EditValue?.ToString();
            string sCSUDN7 = tx_CSUDN7.EditValue?.ToString();

            string sCGONJ1 = tx_CGONJ1.EditValue?.ToString();
            string sCGONJ2 = tx_CGONJ2.EditValue?.ToString();
            string sCGONJ3 = tx_CGONJ3.EditValue?.ToString();
            string sCGONJ4 = tx_CGONJ4.EditValue?.ToString();
            string sCGONJ5 = tx_CGONJ5.EditValue?.ToString();
            string sCGONJ6 = tx_CGONJ6.EditValue?.ToString();

            double dCSUDN1 = 0;
            double dCSUDN6 = 0;
            double dCSUDN7 = 0;
            double dCGONJ1 = 0;
            double dCGONJ2 = 0;
            double dCGONJ3 = 0;
            double dCGONJ4 = 0;
            double dCGONJ5 = 0;
            double dCGONJ6 = 0;

            double.TryParse(sCSUDN1, out dCSUDN1);
            double.TryParse(sCSUDN6, out dCSUDN6);
            double.TryParse(sCSUDN7, out dCSUDN7);
            double.TryParse(sCGONJ1, out dCGONJ1);
            double.TryParse(sCGONJ2, out dCGONJ2);
            double.TryParse(sCGONJ3, out dCGONJ3);
            double.TryParse(sCGONJ4, out dCGONJ4);
            double.TryParse(sCGONJ5, out dCGONJ5);
            double.TryParse(sCGONJ6, out dCGONJ6);

            string sRk = MemRk.EditValue?.ToString();
            string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();

            Dictionary<string, object> dicParams = new Dictionary<string, object>();

            if(_ModeGubun != ModeGubun.Modi)
            {
                dicParams.Clear();
                dicParams.Add("CMD", "LIST2-1");
                dicParams.Add("BASYM", sBasym);
                dicParams.Add("EMPID", sEmpid);

                DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if(dtChk != null)
                {
                    if(int.TryParse(dtChk.Rows[0]["CNT"]?.ToString(), out int iResult))
                    {
                        if(iResult > 0)
                        {
                            if (XtraMessageBox.Show("이미 " + sBasym + " 일자의 데이터가 존재합니다. 수정하시겠습니까?", "수정확인", MessageBoxButtons.YesNo) == DialogResult.No)
                            {
                                //DateBas.Focus();
                                return null;
                            }
                        }
                    }
                }
            }

            dicParams.Clear();
            dicParams.Add("CMD", "SAVE1");
            dicParams.Add("BASYM", sBasym);
            dicParams.Add("EMPID", sEmpid);

            dicParams.Add("CSUDN1", dCSUDN1);
            dicParams.Add("CSUDN6", dCSUDN6);
            dicParams.Add("CSUDN7", dCSUDN7);

            dicParams.Add("CGONJ1", dCGONJ1);
            dicParams.Add("CGONJ2", dCGONJ2);
            dicParams.Add("CGONJ3", dCGONJ3);
            dicParams.Add("CGONJ4", dCGONJ4);
            dicParams.Add("CGONJ5", dCGONJ5);
            dicParams.Add("CGONJ6", dCGONJ6);
            dicParams.Add("RK", sRk);

            if (_ModeGubun == ModeGubun.Add)
            {
                dicParams.Add("USER", sUser);
            }
            else if (_ModeGubun == ModeGubun.Modi)
            {
                //dicParams.Add("SLINO", _sSlino);
                dicParams.Add("USER", sUser);
            }

            return DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
        }

        private void BtnMtSv_Click(object sender, EventArgs e)
        {
            string sEmpid = LkupEmpid.EditValue?.ToString();
            string sBasdt = DateBASYM.EditValue?.ToString();
            if (string.IsNullOrEmpty(sEmpid))
            {
                XtraMessageBox.Show("사원을 선택해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sBasdt))
            {
                XtraMessageBox.Show("일자를 입력해주세요.");
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
                    DateBASYM.Focus();
                    SetEditEmpty();

                    SL001F00 frm = (SL001F00)this.Owner;
                    frm.BtnRetr_Click(null, null);
                }
                else
                {
                    DateBASYM.Focus();
                }
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SL001F00_POP01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnMtSv.PerformClick();
        }

        private void LkupEmpid_EditValueChanged(object sender, EventArgs e)
        {
            string sEmpid = LkupEmpid.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sEmpid))
            {
                if (_ModeGubun == ModeGubun.Add)
                {
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Clear();
                    dicParams.Add("CMD", "LIST3");
                    dicParams.Add("EMPID", sEmpid);

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                    if (dt != null && dt.Rows.Count != 0)
                    {
                        string sDept = dt.Rows[0]["DEPT_CD"]?.ToString();
                        string sJw = dt.Rows[0]["GRADE_CD"]?.ToString();
                        string sGIBON  = dt.Rows[0]["GIBON"]?.ToString();
                        string sSIGUB1 = dt.Rows[0]["SIGUB1"]?.ToString();
                        string sSIGUB2 = dt.Rows[0]["SIGUB2"]?.ToString();

                        LkupDept.EditValue = sDept;
                        LkupJikgub.EditValue = sJw;
                        TxtGibon.EditValue = sGIBON;
                        TxtSigub.EditValue = sSIGUB1;
                        TxtYjsigub.EditValue = sSIGUB2;
                    }

                    SetWkTime();
                }
            }
        }

        private void DateBASYM_EditValueChanged(object sender, EventArgs e)
        {
            SetWkTime();
        }

        private void SetWkTime()
        {
            string sBasym = DateBASYM.EditValue?.ToString().Substring(0, 7);
            string sEmpid = LkupEmpid.EditValue?.ToString();

            if(!string.IsNullOrEmpty(sBasym) && !string.IsNullOrEmpty(sEmpid))
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("EMPID", sEmpid);
                dicParams.Add("BASYM", sBasym);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null && dt.Rows.Count != 0)
                {
                    string sJANUP = dt.Rows[0]["JANUP"]?.ToString();
                    string sYAGUN = dt.Rows[0]["YAGUN"]?.ToString();
                    string sTKGUN = dt.Rows[0]["TKGUN"]?.ToString();
                    string sDANG = dt.Rows[0]["DANG"]?.ToString();
                    string sSUK = dt.Rows[0]["SUK"]?.ToString();
                    string sCHGAM = dt.Rows[0]["CHGAM"]?.ToString();

                    TxtJANUP.EditValue = sJANUP;
                    TxtYAGUN.EditValue = sYAGUN;
                    TxtTKGUN.EditValue = sTKGUN;
                    TxtDANG.EditValue = sDANG;
                    TxtSUK.EditValue = sSUK;
                    TxtCHGAM.EditValue = sCHGAM;

                    double dJANUP = 0;
                    double dYAGUN = 0;
                    double dTKGUN = 0;
                    double dDANG = 0;
                    double dSUK = 0;
                    double dCHGAM = 0;

                    double.TryParse(sJANUP, out dJANUP);
                    double.TryParse(sYAGUN, out dYAGUN);
                    double.TryParse(sTKGUN, out dTKGUN);
                    double.TryParse(sDANG, out dDANG);
                    double.TryParse(sSUK, out dSUK);
                    double.TryParse(sCHGAM, out dCHGAM);

                    double dTotJan = (dJANUP + dYAGUN + dTKGUN + dDANG + dSUK) - dCHGAM;

                    TxtTotJan.EditValue = dTotJan;
                }
            }
        }
        #endregion

        private void LkupJikgub_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                TxtGibon.Focus();
        }

        private void TxtTotJan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                tx_CGONJ1.Focus();
        }

        private void tx_CSUDN7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                MemRk.Focus();
        }

        private void LkupEmpid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                LkupDept.Focus();
        }

        private void MemRk_KeyDown(object sender, KeyEventArgs e)
        {
            if (_ModeGubun == ModeGubun.Add)
            {
                if (e.KeyCode == Keys.Enter)
                    BtnMtSv.Focus();
            }
            else if (_ModeGubun == ModeGubun.Modi)
            {
                if (e.KeyCode == Keys.Enter)
                    BtnSave.Focus();
            }
            
        }

        private void TxtTotJan_EditValueChanged(object sender, EventArgs e)
        {
            if(_ModeGubun == ModeGubun.Add)
            {
                string sYJSIGUB = TxtYjsigub.EditValue?.ToString();
                string sCHGAM = TxtCHGAM.EditValue?.ToString();

                string sJANUP = TxtJANUP.EditValue?.ToString();
                string sYAGUN = TxtYAGUN.EditValue?.ToString();
                string sTKGUN = TxtTKGUN.EditValue?.ToString();
                string sDANG = TxtDANG.EditValue?.ToString();
                string sSUK = TxtSUK.EditValue?.ToString();

                double dYJSIGUB = 0;
                double dCHGAM = 0;

                double dJANUP = 0;
                double dYAGUN = 0;
                double dTKGUN = 0;
                double dDANG = 0;
                double dSUK = 0;

                double.TryParse(sYJSIGUB, out dYJSIGUB);
                double.TryParse(sCHGAM, out dCHGAM);

                double.TryParse(sJANUP, out dJANUP);
                double.TryParse(sYAGUN, out dYAGUN);
                double.TryParse(sTKGUN, out dTKGUN);
                double.TryParse(sDANG, out dDANG);
                double.TryParse(sSUK, out dSUK);

                tx_CSUDN1.EditValue = Math.Round(((dJANUP + dYAGUN + dTKGUN + dDANG - dCHGAM) * dYJSIGUB) + (dSUK * 50000), 0, MidpointRounding.AwayFromZero);
            }
        }
    }
}