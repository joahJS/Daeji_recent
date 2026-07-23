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
    public partial class SL001F01_POP01 : DevExpress.XtraEditors.XtraForm
    {
        public SL001F01_POP01()
        {
            InitializeComponent();
        }

        public DataRow _ModRow;
        string _sSlino = string.Empty;
        public DataRow _InsaInfo;
        public string Dtstdt2;

        public enum ModeGubun { Add, Modi }
        public ModeGubun _ModeGubun = ModeGubun.Add;

        private string PROCEDURE_ID = "DP_SL001F01";

        private void SL001F01_POP01_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            ComnEtcFunc.SetBoundLookUp(LkupEmpid, "HR_EMP_BASIS", "", "");
            ComnEtcFunc.SetBoundLookUp(LkupDept, "ACC_DEPT_CD", "DEPT_CD", "DEPT_NM");
            ComnEtcFunc.SetBoundLookUp(LkupJikgub, "COM_BASE_CD", "GRADE_CD", "");
            ComnEtcFunc.SetBoundLookUp(LkupPBKCOD, "COM_BASE_CD", "BANK_CD", "");

            string today1 = DateTime.Today.ToString();
            DatePAYDT.EditValue = today1.Substring(0, 10);
            DateBASYM.EditValue = Dtstdt2;
            SetPlus();

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
                    DatePAYDT.EditValue = _ModRow["PAYDT"];
                    LkupDept.EditValue = _ModRow["PDEPT"];
                    LkupJikgub.EditValue = _ModRow["PJKWK"];

                    //기존수당
                    TxtPGJSD1.EditValue = _ModRow["PGJSD1"];
                    TxtPGJSD2.EditValue = _ModRow["PGJSD2"];
                    //TxtPGJSD3.EditValue = _ModRow["PGJSD3"];
                    //TxtPGJSD4.EditValue = _ModRow["PGJSD4"];
                    //TxtPGJSD5.EditValue = _ModRow["PGJSD5"];
                    //TxtPGJSD6.EditValue = _ModRow["PGJSD6"];
                    //기존공제
                    TxtPCHSD1.EditValue = _ModRow["PCHSD1"];
                    //TxtPCHSD2.EditValue = _ModRow["PCHSD2"];
                    //TxtPCHSD3.EditValue = _ModRow["PCHSD3"];
                    //TxtPCHSD4.EditValue = _ModRow["PCHSD4"];
                    //TxtPCHSD5.EditValue = _ModRow["PCHSD5"];
                    TxtPCHSD6.EditValue = _ModRow["PCHSD6"];
                    TxtPCHSD7.EditValue = _ModRow["PCHSD7"];
                    //TxtPCHSD8.EditValue = _ModRow["PCHSD8"];

                    //변동수당
                    TxtPGJGJ1.EditValue = _ModRow["PGJGJ1"];
                    TxtPGJGJ2.EditValue = _ModRow["PGJGJ2"];
                    TxtPGJGJ3.EditValue = _ModRow["PGJGJ3"];
                    TxtPGJGJ4.EditValue = _ModRow["PGJGJ4"];
                    //TxtPGJGJ5.EditValue = _ModRow["PGJGJ5"];

                    //변동공제
                    TxtPCHGJ1.EditValue = _ModRow["PCHGJ1"];
                    TxtPCHGJ2.EditValue = _ModRow["PCHGJ2"];
                    TxtPCHGJ3.EditValue = _ModRow["PCHGJ3"];
                    TxtPCHGJ4.EditValue = _ModRow["PCHGJ4"];
                    TxtPCHGJ5.EditValue = _ModRow["PCHGJ5"];

                    TxtPGROSS.EditValue = _ModRow["PGROSS"];
                    TxtPTOTGO.EditValue = _ModRow["PTOTGO"];
                    TxtPCHAIN.EditValue = _ModRow["PCHAIN"];

                    MemRk.EditValue = _ModRow["RK"];

                    _ModRow = null;
                }
            }
        }

        private void SetEditEmpty()
        {
            string today1 = DateTime.Today.ToString();

            if(_ModeGubun == ModeGubun.Add)
            {
                DateBASYM.EditValue = today1.Substring(0, 7);

                LkupEmpid.EditValue = null;
                DatePAYDT.EditValue = today1;
                LkupDept.EditValue = null;
                LkupJikgub.EditValue = null;

                //기존수당
                TxtPGJSD1.EditValue = 0;
                TxtPGJSD2.EditValue = 0;
                //TxtPGJSD3.EditValue = _ModRow["PGJSD3"];
                //TxtPGJSD4.EditValue = _ModRow["PGJSD4"];
                //TxtPGJSD5.EditValue = _ModRow["PGJSD5"];
                //TxtPGJSD6.EditValue = _ModRow["PGJSD6"];
                //기존공제
                TxtPCHSD1.EditValue = 0;
                //TxtPCHSD2.EditValue = _ModRow["PCHSD2"];
                //TxtPCHSD3.EditValue = _ModRow["PCHSD3"];
                //TxtPCHSD4.EditValue = _ModRow["PCHSD4"];
                //TxtPCHSD5.EditValue = _ModRow["PCHSD5"];
                TxtPCHSD6.EditValue = 0;
                TxtPCHSD7.EditValue = 0;
                //TxtPCHSD8.EditValue = _ModRow["PCHSD8"];

                //변동수당
                TxtPGJGJ1.EditValue = 0;
                TxtPGJGJ2.EditValue = 0;
                TxtPGJGJ3.EditValue = 0;
                TxtPGJGJ4.EditValue = 0;
                //TxtPGJGJ5.EditValue = _ModRow["PGJGJ5"];

                //변동공제
                TxtPCHGJ1.EditValue = 0;
                TxtPCHGJ2.EditValue = 0;
                TxtPCHGJ3.EditValue = 0;
                TxtPCHGJ4.EditValue = 0;
                TxtPCHGJ5.EditValue = 0;

                TxtPGROSS.EditValue = 0;
                TxtPTOTGO.EditValue = 0;
                TxtPCHAIN.EditValue = 0;

                MemRk.EditValue = null;
            }
            else if (_ModeGubun == ModeGubun.Modi)
            {
                if (_ModRow != null)
                {
                    LkupEmpid.EditValue = _ModRow["EMPID"];
                    DateBASYM.EditValue = _ModRow["BASYM"];
                    DatePAYDT.EditValue = _ModRow["PAYDT"];
                    LkupDept.EditValue = _ModRow["PDEPT"];
                    LkupJikgub.EditValue = _ModRow["PJKWK"];

                    //기존수당
                    TxtPGJSD1.EditValue = _ModRow["PGJSD1"];
                    TxtPGJSD2.EditValue = _ModRow["PGJSD2"];
                    //TxtPGJSD3.EditValue = _ModRow["PGJSD3"];
                    //TxtPGJSD4.EditValue = _ModRow["PGJSD4"];
                    //TxtPGJSD5.EditValue = _ModRow["PGJSD5"];
                    //TxtPGJSD6.EditValue = _ModRow["PGJSD6"];
                    //기존공제
                    TxtPCHSD1.EditValue = _ModRow["PCHSD1"];
                    //TxtPCHSD2.EditValue = _ModRow["PCHSD2"];
                    //TxtPCHSD3.EditValue = _ModRow["PCHSD3"];
                    //TxtPCHSD4.EditValue = _ModRow["PCHSD4"];
                    //TxtPCHSD5.EditValue = _ModRow["PCHSD5"];
                    TxtPCHSD6.EditValue = _ModRow["PCHSD6"];
                    TxtPCHSD7.EditValue = _ModRow["PCHSD7"];
                    //TxtPCHSD8.EditValue = _ModRow["PCHSD8"];

                    //변동수당
                    TxtPGJGJ1.EditValue = _ModRow["PGJGJ1"];
                    TxtPGJGJ2.EditValue = _ModRow["PGJGJ2"];
                    TxtPGJGJ3.EditValue = _ModRow["PGJGJ3"];
                    TxtPGJGJ4.EditValue = _ModRow["PGJGJ4"];
                    //TxtPGJGJ5.EditValue = _ModRow["PGJGJ5"];

                    //변동공제
                    TxtPCHGJ1.EditValue = _ModRow["PCHGJ1"];
                    TxtPCHGJ2.EditValue = _ModRow["PCHGJ2"];
                    TxtPCHGJ3.EditValue = _ModRow["PCHGJ3"];
                    TxtPCHGJ4.EditValue = _ModRow["PCHGJ4"];
                    TxtPCHGJ5.EditValue = _ModRow["PCHGJ5"];

                    TxtPGROSS.EditValue = _ModRow["PGROSS"];
                    TxtPTOTGO.EditValue = _ModRow["PTOTGO"];
                    TxtPCHAIN.EditValue = _ModRow["PCHAIN"];

                    MemRk.EditValue = _ModRow["RK"];

                    _ModRow = null;
                }
            }
        }

        #region [버튼클릭이벤트]
        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sEmpid = LkupEmpid.EditValue?.ToString();
            string sBasdt = DateBASYM.EditValue?.ToString();
            string sPAYDT = DatePAYDT.EditValue?.ToString();

            if (string.IsNullOrEmpty(sEmpid))
            {
                XtraMessageBox.Show("사원을 입력해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sBasdt))
            {
                XtraMessageBox.Show("급여년월을 입력해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sPAYDT))
            {
                XtraMessageBox.Show("지급일자를 입력해주세요.");
                return;
            }

            DataTable result = SaveInsaChangeData();

            if (result != null)
            {
                int sResult = int.Parse(result.Rows[0]["RESULT"]?.ToString());
                string sMsg = result.Rows[0]["MSG"]?.ToString();
                //string sSlino = string.Empty;
                //if (result.Rows[0]["SLINO"] != null)
                //    sSlino = result.Rows[0]["SLINO"]?.ToString();

                XtraMessageBox.Show(sMsg);

                if (sResult > 0)
                {
                    SL001F01 frm = (SL001F01)this.Owner;
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
            string sEMPID = LkupEmpid.EditValue?.ToString();
            string sBASYM = DateBASYM.EditValue?.ToString();
            string sPAYDT = DatePAYDT.EditValue?.ToString();

            string sPDEPT = LkupDept.EditValue?.ToString();
            string sPJKWK = LkupJikgub.EditValue?.ToString();

            string sPWKTM    = TxtPWKTM.EditValue?.ToString();
            string sPYJTM    = TxtPYJTM.EditValue?.ToString();
            string sPGIBON   = TxtPGIBON.EditValue?.ToString();
            string sPSIGUB1  = TxtPSIGUB1.EditValue?.ToString();
            string sPSIGUB2  = TxtPSIGUB2.EditValue?.ToString();
            string sPGJSD1   = TxtPGJSD1.EditValue?.ToString();
            string sPGJSD2   = TxtPGJSD2.EditValue?.ToString();
            //string sPGJSD3   = TxtPGJSD3.EditValue?.ToString();
            //string sPGJSD4   = TxtPGJSD4.EditValue?.ToString();
            //string sPGJSD5   = TxtPGJSD5.EditValue?.ToString();
            //string sPGJSD6   = TxtPGJSD6.EditValue?.ToString();
            //string sPGJSD7   = TxtPGJSD7.EditValue?.ToString();
            string sPCHSD1   = TxtPCHSD1.EditValue?.ToString();
            //string sPCHSD2   = TxtPCHSD2.EditValue?.ToString();
            //string sPCHSD3   = TxtPCHSD3.EditValue?.ToString();
            //string sPCHSD4   = TxtPCHSD4.EditValue?.ToString();
            //string sPCHSD5   = TxtPCHSD5.EditValue?.ToString();
            string sPCHSD6   = TxtPCHSD6.EditValue?.ToString();
            string sPCHSD7   = TxtPCHSD7.EditValue?.ToString();
            //string sPCHSD8   = TxtPCHSD8.EditValue?.ToString();
            string sPGROSS   = TxtPGROSS.EditValue?.ToString();
            string sPGJGJ1   = TxtPGJGJ1.EditValue?.ToString();
            string sPGJGJ2   = TxtPGJGJ2.EditValue?.ToString();
            string sPGJGJ3   = TxtPGJGJ3.EditValue?.ToString();
            string sPGJGJ4   = TxtPGJGJ4.EditValue?.ToString();
            //string sPGJGJ5   = TxtPGJGJ5.EditValue?.ToString();
            //string sPGJGJ6   = TxtPGJGJ6.EditValue?.ToString();
            //string sPGJGJ7   = TxtPGJGJ7.EditValue?.ToString();
            string sPCHGJ1   = TxtPCHGJ1.EditValue?.ToString();
            string sPCHGJ2   = TxtPCHGJ2.EditValue?.ToString();
            string sPCHGJ3   = TxtPCHGJ3.EditValue?.ToString();
            string sPCHGJ4   = TxtPCHGJ4.EditValue?.ToString();
            string sPCHGJ5   = TxtPCHGJ5.EditValue?.ToString();
            string sPCHGJ6   = TxtPCHGJ6.EditValue?.ToString();
            //string sPCHGJ7   = TxtPCHGJ7.EditValue?.ToString();
            string sPTOTGO   = TxtPTOTGO.EditValue?.ToString();
            string sPCHAIN   = TxtPCHAIN.EditValue?.ToString();
            string sRK       = MemRk.EditValue?.ToString();
            string sPBKCOD   = LkupPBKCOD.EditValue?.ToString();
            string sPBKNUM   = TxtPBKNUM.EditValue?.ToString();
            string sPBKOWN = TxtPBKOWN.EditValue?.ToString();
            string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST2-1");
            dicParams.Add("BASYM", sBASYM);
            dicParams.Add("EMPID", sEMPID);

            DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            if (dtChk != null && dtChk.Rows.Count > 0 && _ModeGubun != ModeGubun.Modi)
            {
                if (XtraMessageBox.Show("이미 " + sBASYM + " 일자의 데이터가 존재합니다. 수정하시겠습니까?", "수정확인", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    //DateBas.Focus();
                    return null;
                }
            }

            dicParams.Clear();
            dicParams.Add("CMD", "SAVE1");
            dicParams.Add("BASYM", sBASYM  );
            dicParams.Add("PAYDT", sPAYDT);
            dicParams.Add("EMPID", sEMPID  );
            dicParams.Add("PDEPT", sPDEPT  );
            dicParams.Add("PJKWK", sPJKWK  );
            dicParams.Add("PWKTM", sPWKTM  );
            dicParams.Add("PYJTM", sPYJTM  );
            dicParams.Add("PGIBON", sPGIBON );
            dicParams.Add("PSIGUB1", sPSIGUB1);
            dicParams.Add("PSIGUB2", sPSIGUB2);
            dicParams.Add("PGJSD1", sPGJSD1 );
            dicParams.Add("PGJSD2", sPGJSD2 );
            //dicParams.Add("PGJSD3", sPGJSD3 );
            //dicParams.Add("PGJSD4", sPGJSD4 );
            //dicParams.Add("PGJSD5", sPGJSD5 );
            //dicParams.Add("PGJSD6", sPGJSD6 );
            //dicParams.Add("PGJSD7", sPGJSD7 );
            dicParams.Add("PCHSD1", sPCHSD1 );
            //dicParams.Add("PCHSD2", sPCHSD2 );
            //dicParams.Add("PCHSD3", sPCHSD3 );
            //dicParams.Add("PCHSD4", sPCHSD4 );
            //dicParams.Add("PCHSD5", sPCHSD5 );
            dicParams.Add("PCHSD6", sPCHSD6 );
            dicParams.Add("PCHSD7", sPCHSD7 );
            //dicParams.Add("PCHSD8", sPCHSD8 );
            dicParams.Add("PGROSS", sPGROSS );
            dicParams.Add("PGJGJ1", sPGJGJ1 );
            dicParams.Add("PGJGJ2", sPGJGJ2 );
            dicParams.Add("PGJGJ3", sPGJGJ3 );
            dicParams.Add("PGJGJ4", sPGJGJ4 );
            //dicParams.Add("PGJGJ5", sPGJGJ5 );
            //dicParams.Add("PGJGJ6", sPGJGJ6 );
            //dicParams.Add("PGJGJ7", sPGJGJ7 );
            dicParams.Add("PCHGJ1", sPCHGJ1 );
            dicParams.Add("PCHGJ2", sPCHGJ2 );
            dicParams.Add("PCHGJ3", sPCHGJ3 );
            dicParams.Add("PCHGJ4", sPCHGJ4 );
            dicParams.Add("PCHGJ5", sPCHGJ5 );
            dicParams.Add("PCHGJ6", sPCHGJ6 );
            //dicParams.Add("PCHGJ7", sPCHGJ7 );
            dicParams.Add("PTOTGO", sPTOTGO );
            dicParams.Add("PCHAIN", sPCHAIN );
            dicParams.Add("RK", sRK);
            dicParams.Add("PBKCOD", sPBKCOD );
            dicParams.Add("PBKNUM", sPBKNUM );
            dicParams.Add("PBKOWN", sPBKOWN);


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
            string sPAYDT = DatePAYDT.EditValue?.ToString();

            if (string.IsNullOrEmpty(sEmpid))
            {
                XtraMessageBox.Show("사원을 입력해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sBasdt))
            {
                XtraMessageBox.Show("급여년월을 입력해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(sPAYDT))
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
                    DateBASYM.Focus();
                    SetEditEmpty();

                    SL001F01 frm = (SL001F01)this.Owner;
                    //frm._SLINO = result.Rows[0]["SLINO"]?.ToString();
                    frm.BtnRetr_Click(null, null);
                }
                else
                {
                    DateBASYM.Focus();
                }
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            SetEditEmpty();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SL001F01_POP01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnMtSv.PerformClick();
        }

        #endregion

        #region[계산]

        private void TxtPGJSD1_EditValueChanged(object sender, EventArgs e)
        {
            SetPlus();
        }

        private void SetPlus()
        {
            string sPGJSD1 = TxtPGJSD1.EditValue?.ToString();
            string sPGJSD2 = TxtPGJSD2.EditValue?.ToString();
            //string sPGJSD3 = TxtPGJSD3.EditValue?.ToString();
            //string sPGJSD4 = TxtPGJSD4.EditValue?.ToString();
            //string sPGJSD5 = TxtPGJSD5.EditValue?.ToString();
            //string sPGJSD6 = TxtPGJSD6.EditValue?.ToString();
            //string sPGJSD7 = TxtPGJSD7.EditValue?.ToString();
            string sPCHSD1 = TxtPCHSD1.EditValue?.ToString();
            //string sPCHSD2 = TxtPCHSD2.EditValue?.ToString();
            //string sPCHSD3 = TxtPCHSD3.EditValue?.ToString();
            //string sPCHSD4 = TxtPCHSD4.EditValue?.ToString();
            //string sPCHSD5 = TxtPCHSD5.EditValue?.ToString();
            string sPCHSD6 = TxtPCHSD6.EditValue?.ToString();
            string sPCHSD7 = TxtPCHSD7.EditValue?.ToString();
            //string sPCHSD8 = TxtPCHSD8.EditValue?.ToString();
            string sPGJGJ1 = TxtPGJGJ1.EditValue?.ToString();
            string sPGJGJ2 = TxtPGJGJ2.EditValue?.ToString();
            string sPGJGJ3 = TxtPGJGJ3.EditValue?.ToString();
            string sPGJGJ4 = TxtPGJGJ4.EditValue?.ToString();
            //string sPGJGJ5 = TxtPGJGJ5.EditValue?.ToString();
            //string sPGJGJ6 = TxtPGJGJ6.EditValue?.ToString();
            //string sPGJGJ7 = TxtPGJGJ7.EditValue?.ToString();
            string sPCHGJ1 = TxtPCHGJ1.EditValue?.ToString();
            string sPCHGJ2 = TxtPCHGJ2.EditValue?.ToString();
            string sPCHGJ3 = TxtPCHGJ3.EditValue?.ToString();
            string sPCHGJ4 = TxtPCHGJ4.EditValue?.ToString();
            string sPCHGJ5 = TxtPCHGJ5.EditValue?.ToString();
            string sPCHGJ6 = TxtPCHGJ6.EditValue?.ToString();
            //string sPCHGJ7 = TxtPCHGJ7.EditValue?.ToString();

            double dPGJSD1 = 0;
            double dPGJSD2 = 0;
            double dPGJSD3 = 0;
            double dPGJSD4 = 0;
            double dPGJSD5 = 0;
            double dPGJSD6 = 0;
            double dPGJSD7 = 0;
            double dPCHSD1 = 0;
            double dPCHSD2 = 0;
            double dPCHSD3 = 0;
            double dPCHSD4 = 0;
            double dPCHSD5 = 0;
            double dPCHSD6 = 0;
            double dPCHSD7 = 0;
            double dPCHSD8 = 0;
            double dPGJGJ1 = 0;
            double dPGJGJ2 = 0;
            double dPGJGJ3 = 0;
            double dPGJGJ4 = 0;
            double dPGJGJ5 = 0;
            double dPGJGJ6 = 0;
            double dPGJGJ7 = 0;
            double dPCHGJ1 = 0;
            double dPCHGJ2 = 0;
            double dPCHGJ3 = 0;
            double dPCHGJ4 = 0;
            double dPCHGJ5 = 0;
            double dPCHGJ6 = 0;
            double dPCHGJ7 = 0;

            double.TryParse(sPGJSD1, out dPGJSD1);
            double.TryParse(sPGJSD2, out dPGJSD2);
            //double.TryParse(sPGJSD3, out dPGJSD3);
            //double.TryParse(sPGJSD4, out dPGJSD4);
            //double.TryParse(sPGJSD5, out dPGJSD5);
            //double.TryParse(sPGJSD6, out dPGJSD6);
            //double.TryParse(sPGJSD7, out dPGJSD7);
            double.TryParse(sPCHSD1, out dPCHSD1);
            //double.TryParse(sPCHSD2, out dPCHSD2);
            //double.TryParse(sPCHSD3, out dPCHSD3);
            //double.TryParse(sPCHSD4, out dPCHSD4);
            //double.TryParse(sPCHSD5, out dPCHSD5);
            double.TryParse(sPCHSD6, out dPCHSD6);
            double.TryParse(sPCHSD7, out dPCHSD7);
            //double.TryParse(sPCHSD8, out dPCHSD8);
            double.TryParse(sPGJGJ1, out dPGJGJ1);
            double.TryParse(sPGJGJ2, out dPGJGJ2);
            double.TryParse(sPGJGJ3, out dPGJGJ3);
            double.TryParse(sPGJGJ4, out dPGJGJ4);
            //double.TryParse(sPGJGJ5, out dPGJGJ5);
            //double.TryParse(sPGJGJ6, out dPGJGJ6);
            //double.TryParse(sPGJGJ7, out dPGJGJ7);
            double.TryParse(sPCHGJ1, out dPCHGJ1);
            double.TryParse(sPCHGJ2, out dPCHGJ2);
            double.TryParse(sPCHGJ3, out dPCHGJ3);
            double.TryParse(sPCHGJ4, out dPCHGJ4);
            double.TryParse(sPCHGJ5, out dPCHGJ5);
            double.TryParse(sPCHGJ6, out dPCHGJ6);
            //double.TryParse(sPCHGJ7, out dPCHGJ7);

            TxtPGROSS.EditValue = dPGJSD1 + dPGJSD2 + dPGJSD3 + dPGJSD4 + dPGJSD5 + dPGJSD6 + dPGJSD7
                                + dPCHSD1 + dPCHSD2 + dPCHSD3 + dPCHSD4 + dPCHSD5 + dPCHSD6 + dPCHSD7 + dPCHSD8;
            TxtPTOTGO.EditValue = dPGJGJ1 + dPGJGJ2 + dPGJGJ3 + dPGJGJ4 + dPGJGJ5 + dPGJGJ6 + dPGJGJ7
                                + dPCHGJ1 + dPCHGJ2 + dPCHGJ3 + dPCHGJ4 + dPCHGJ5 + dPCHGJ6 + dPCHGJ7;
            TxtPCHAIN.EditValue = dPGJSD1 + dPGJSD2 + dPGJSD3 + dPGJSD4 + dPGJSD5 + dPGJSD6 + dPGJSD7
                                + dPCHSD1 + dPCHSD2 + dPCHSD3 + dPCHSD4 + dPCHSD5 + dPCHSD6 + dPCHSD7 + dPCHSD8
                                - dPGJGJ1 - dPGJGJ2 - dPGJGJ3 - dPGJGJ4 - dPGJGJ5 - dPGJGJ6 - dPGJGJ7
                                - dPCHGJ1 - dPCHGJ2 - dPCHGJ3 - dPCHGJ4 - dPCHGJ5 - dPCHGJ6 - dPCHGJ7;
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
                        string sBKOWN = dt.Rows[0]["ACNT_HDR"]?.ToString();
                        string sBKCOD = dt.Rows[0]["DEAL_BANK_CD"]?.ToString();
                        string sBKNUM = dt.Rows[0]["PMNT_ACNT_FST"]?.ToString();
                        string sBTIME1 = dt.Rows[0]["BASTIME1"]?.ToString();
                        string sBTIME2 = dt.Rows[0]["BASTIME2"]?.ToString();

                        LkupDept.EditValue = sDept;
                        LkupJikgub.EditValue = sJw;
                        LkupPBKCOD.EditValue = sBKCOD;
                        TxtPBKOWN.EditValue = sBKOWN;
                        TxtPBKNUM.EditValue = sBKNUM;
                        TxtPWKTM.EditValue = sBTIME1;
                        TxtPYJTM.EditValue = sBTIME2;
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

            if (!string.IsNullOrEmpty(sBasym) && !string.IsNullOrEmpty(sEmpid))
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
                    TxtChagam.EditValue = sCHGAM;

                    double dJANUP = 0;
                    double dYAGUN = 0;
                    double dTKGUN = 0;
                    double dDANG  = 0;
                    double dSUK   = 0;
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
    }

}