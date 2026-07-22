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
    public partial class GE001F01 : DevExpress.XtraEditors.XtraForm
    {
        public GE001F01()
        {
            InitializeComponent();
        }

        private static double _YAGAN = 18; //야간 기준시간. 18시 이후 출근자: 야간
        private static string _NTBAB = string.Empty;//잔업시작시간

        private static string _YAGAN1 = string.Empty;//야간시작
        private static string _YAGAN2 = string.Empty;//야간종료

        private static string _ILBAN1 = string.Empty;//일반출근
        private static string _ILBAN2 = string.Empty;//일반퇴근
        private static string _YUKIC1 = string.Empty;//여직원출근
        private static string _YUKIC2 = string.Empty;//여직원퇴근

        string sGENDER = string.Empty;

        private void GE001F01_Load(object sender, EventArgs e)
        {
            SetGETimes();

            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.gp_SetColorFocused(layoutControl3);
            ComnEtcFunc.SetDateToValue(DateBas);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            //근태구분
            dicParams.Clear();
            dicParams.Add("CMD", "COMCD");
            dicParams.Add("RCDTP", "WKGUB");
            DataTable dtWkgub = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);
            ComLib.ComGrid.SetLookUpEdit(LkupWkgb, dtWkgub, "CD", "NM", "");

            //dicParams.Clear();
            //dicParams.Add("CMD", "COMCD2");
            //dicParams.Add("RCDTP", "SNEGB");
            //dicParams.Add("SUBCOD1", "1");//출근구분
            //DataTable dtSRTGB = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);
            //ComLib.ComGrid.SetLookUpEdit(LkupStrgb, dtSRTGB, "CD", "NM", "");

            //dicParams.Clear();
            //dicParams.Add("CMD", "COMCD2");
            //dicParams.Add("RCDTP", "SNEGB");
            //dicParams.Add("SUBCOD1", "2");//퇴근구분
            //DataTable dtENDGB = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);
            //ComLib.ComGrid.SetLookUpEdit(LkupEndgb, dtENDGB, "CD", "NM", "");

            dicParams.Clear();
            dicParams.Add("CMD", "GETINSA");
            DataTable dtInsa = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);
            ComLib.ComGrid.SetLookUpEdit(lkupUsr, dtInsa, "CD", "NM", "");

        }

        private void SetGETimes()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "COMCD3");
                dicParams.Add("RCDTP", "GMUGB");//근태구분
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);

                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sComcd = dt.Rows[i]["COM_CD"]?.ToString();
                        string sComSub1 = dt.Rows[i]["COM_SUB_CD1"]?.ToString();
                        string sComSub2 = dt.Rows[i]["COM_SUB_CD2"]?.ToString();

                        if (!string.IsNullOrEmpty(sComcd))
                        {
                            switch (sComcd)
                            {
                                case "0":
                                    _ILBAN1 = sComSub1;//일반출근
                                    _ILBAN2 = sComSub2;//일반퇴근
                                    break;
                                case "1":
                                    _YUKIC1 = sComSub1;//여직원출근
                                    _YUKIC2 = sComSub2;//여직원퇴근
                                    break;
                                case "2":
                                    if (DateTime.TryParse(sComSub1, out DateTime dateTime))
                                    {
                                        _YAGAN = dateTime.Hour; //야간 기준시간. 18시 이후 출근자: 야간
                                    }
                                    break;
                                case "3":
                                    _YAGAN1 = sComSub1;//야간시작
                                    _YAGAN2 = sComSub2;//야간종료
                                    break;
                                case "4":
                                    _NTBAB = sComSub1;//잔업시작시간
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        #region 저장
        private void BtnSvMt_Click(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateBas.EditValue?.ToString().Substring(0, 10)))
            {
                XtraMessageBox.Show("마감된 일자의 근태 데이터는 추가할 수 없습니다.");
                return;
            }

            if (SaveGDAYF())
            {
                DateBas.EditValue = DateTime.Now;
                LkupWkgb.EditValue = null;
                lkupUsr.EditValue = null;
                TxtEmpNo.EditValue = null;
                TxtWkbas.EditValue = 0.0;
                LkupStrgb.EditValue = null;
                LkupEndgb.EditValue = null;
                TxtGWKTM1.EditValue = 0.0;
                TxtGWKTM2.EditValue = 0.0;
                TxtGWKTM3.EditValue = 0.0;
                TxtGWKTM4.EditValue = 0.0;
                TxtGWKTM5.EditValue = 0.0;
                TxtGWKTM6.EditValue = 0.0;
                TxtGWKTM7.EditValue = 0.0;
                TxtGWKTM8.EditValue = 0.0;
                TxtGWKTM8.EditValue = 0.0;
                TxtWKINTM.EditValue = null;
                TxtWKOTTM.EditValue = null;
                TxtGOINTM.EditValue = null;
                TxtGOOTTM.EditValue = null;
                MemoRk1.EditValue = null;

                DateBas.Focus();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateBas.EditValue?.ToString().Substring(0, 10)))
            {
                XtraMessageBox.Show("마감된 일자의 근태 데이터는 추가할 수 없습니다.");
                return;
            }

            if (SaveGDAYF())
                DialogResult = DialogResult.OK;
        }

        //마감확인
        private bool ChkMagamYN(string sDate)
        {
            bool result = false;
            try
            {

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT COUNT(*) AS CNT     ");
                strSql.AppendLine("   FROM GDAYF               ");
                strSql.AppendLine("  WHERE BASDT = '" + sDate + "'");
                strSql.AppendLine("    AND MCHKYN = 'Y'        ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null)
                {
                    double iCnt = 0;
                    double.TryParse(dt.Rows[0]["CNT"]?.ToString(), out iCnt);

                    if (iCnt > 0) //마감된 자료가 있으면 마감됨 true
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 오류");
            }

            return result;
        }

        private bool SaveGDAYF()
        {
            bool result = false;
            string sBASDT = DateBas.EditValue?.ToString();
            string sEMPID = TxtEmpNo.EditValue?.ToString();
            string sWKGUB = LkupWkgb.EditValue?.ToString();
            //string sSRTGB = LkupStrgb.EditValue?.ToString();
            //string sENDGB = LkupEndgb.EditValue?.ToString();

            string sWKINTM = TxtWKINTM.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sWKINTM))
            {
                DateTime dateTime = DateTime.Parse(sWKINTM);
                sWKINTM = dateTime.ToString("HH:mm");
            }

            string sWKOTTM = TxtWKOTTM.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sWKOTTM))
            {
                DateTime dateTime = DateTime.Parse(sWKOTTM);
                sWKOTTM = dateTime.ToString("HH:mm");
            }

            string sGOINTM = TxtGOINTM.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sGOINTM))
            {
                DateTime dateTime = DateTime.Parse(sGOINTM);
                sGOINTM = dateTime.ToString("HH:mm");
            }

            string sGOOTTM = TxtGOOTTM.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sGOOTTM))
            {
                DateTime dateTime = DateTime.Parse(sGOOTTM);
                sGOOTTM = dateTime.ToString("HH:mm");
            }

            double dWKBASE = 0;
            double dGWKTM1 = 0;
            double dGWKTM2 = 0;
            double dGWKTM3 = 0;
            double dGWKTM4 = 0;
            double dGWKTM5 = 0;
            double dGWKTM6 = 0;
            double dGWKTM7 = 0;
            double dGWKTM8 = 0;
            double dGWKTM9 = 0;

            double.TryParse(TxtWkbas.EditValue?.ToString(), out dWKBASE);//근무시간
            double.TryParse(TxtGWKTM1.EditValue?.ToString(), out dGWKTM1);
            double.TryParse(TxtGWKTM2.EditValue?.ToString(), out dGWKTM2);
            double.TryParse(TxtGWKTM3.EditValue?.ToString(), out dGWKTM3);
            double.TryParse(TxtGWKTM4.EditValue?.ToString(), out dGWKTM4);
            double.TryParse(TxtGWKTM5.EditValue?.ToString(), out dGWKTM5);
            double.TryParse(TxtGWKTM6.EditValue?.ToString(), out dGWKTM6);
            double.TryParse(TxtGWKTM7.EditValue?.ToString(), out dGWKTM7);
            double.TryParse(TxtGWKTM8.EditValue?.ToString(), out dGWKTM8);
            double.TryParse(TxtGWKTM9.EditValue?.ToString(), out dGWKTM9);

            string sRK1 = MemoRk1.EditValue?.ToString();
            string sUSRCD = FmMainToolBar2.drUser["USRCD"]?.ToString();


            if (string.IsNullOrEmpty(sBASDT))
            {
                XtraMessageBox.Show("근태일자를 선택해주세요.");
                DateBas.Focus();
                return false;
            }
            else
            {
                sBASDT = sBASDT.Substring(0, 10);
            }

            if (string.IsNullOrEmpty(sEMPID))
            {
                XtraMessageBox.Show("사원을 선택해주세요.");
                lkupUsr.Focus();
                return false;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(0) AS CNT");
            strSql.AppendLine("   FROM GDAYF          ");
            strSql.AppendLine("  WHERE BASDT = '" + sBASDT + "'     ");
            strSql.AppendLine("    AND EMPID = '" + sEMPID + "'     ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt != null && int.Parse(dt.Rows[0]["CNT"]?.ToString()) > 0)
            {
                XtraMessageBox.Show("해당 사원의 " + sBASDT + "일자의 근태 데이터가 이미 존재합니다.");
                return false;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine("INSERT INTO GDAYF( BASDT   ");
                strSql.AppendLine("                 , EMPID   ");
                strSql.AppendLine("                 , WKGUB   ");
                strSql.AppendLine("                 , WKINTM  ");
                strSql.AppendLine("                 , WKOTTM  ");
                strSql.AppendLine("                 , GOINTM  ");
                strSql.AppendLine("                 , GOOTTM  ");
                strSql.AppendLine("                 , WKBASE  ");
                strSql.AppendLine("                 , GWKTM1  ");
                strSql.AppendLine("                 , GWKTM2  ");
                strSql.AppendLine("                 , GWKTM3  ");
                strSql.AppendLine("                 , GWKTM4  ");
                strSql.AppendLine("                 , GWKTM5  ");
                strSql.AppendLine("                 , GWKTM6  ");
                strSql.AppendLine("                 , GWKTM7  ");
                strSql.AppendLine("                 , GWKTM8  ");
                strSql.AppendLine("                 , GWKTM9  ");
                strSql.AppendLine("                 , RK1     ");
                strSql.AppendLine("                 , CUSER)  ");
                strSql.AppendLine("           VALUES( '" + sBASDT + "'      ");
                strSql.AppendLine("                 , '" + sEMPID + "'");
                strSql.AppendLine("                 , '" + sWKGUB + "'");
                strSql.AppendLine("                 , '" + sWKINTM + "'");
                strSql.AppendLine("                 , '" + sWKOTTM + "'");
                strSql.AppendLine("                 , '" + sGOINTM + "'");
                strSql.AppendLine("                 , '" + sGOOTTM + "'");
                strSql.AppendLine("                 , " + dWKBASE);
                strSql.AppendLine("                 , " + dGWKTM1);
                strSql.AppendLine("                 , " + dGWKTM2);
                strSql.AppendLine("                 , " + dGWKTM3);
                strSql.AppendLine("                 , " + dGWKTM4);
                strSql.AppendLine("                 , " + dGWKTM5);
                strSql.AppendLine("                 , " + dGWKTM6);
                strSql.AppendLine("                 , " + dGWKTM7);
                strSql.AppendLine("                 , " + dGWKTM8);
                strSql.AppendLine("                 , " + dGWKTM9);
                strSql.AppendLine("                 , '" + sRK1 + "'");
                strSql.AppendLine("                 , '" + sUSRCD + "')     "); 

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                GE001F00 P_frm = (GE001F00)this.Owner;
                P_frm._BASDT = sBASDT;
                P_frm._EMPID = sEMPID;

                P_frm.BtnRetr_Click(null, null);

                result = true;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);

                result = false;
            }

            return result;
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void GE001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnSvMt.PerformClick();
        }

        #region 출퇴근 관련 이벤트
        private void TxtTktm1_Leave(object sender, EventArgs e)
        {
            string sBASDT = DateBas.EditValue?.ToString().Substring(0,10);
            string sWKINTM = TxtWKINTM.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sWKINTM))
                sWKINTM = DateTime.Parse(sWKINTM).ToString("HH:mm");
            string sWKOTTM = TxtWKOTTM.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sWKOTTM))
                sWKOTTM = DateTime.Parse(sWKOTTM).ToString("HH:mm");
            string sWkgub = LkupWkgb.EditValue?.ToString();

            //string sSRTGB = string.Empty;//출근구분
            //string sENDGB = string.Empty;//퇴근구분
            //string sWKGUB = string.Empty;//근태구분

            double dAlWKBS = 0.0;//총근무시간
            double dWKBASE = 0.0;//근무시간
            double dGWKTM2 = 0.0;//기본잔업 남직원:1.5, 여직원:0.0
            double dGWKTM3 = 0.0;//추가잔업 
            double dGWKTM4 = 0.0;//당직
            double dGWKTM5 = 0.0;//지각시간
            //double dGWKTM6 = 0.0;//조퇴시간

            DateTime chulTime = new DateTime();//출근시간
            DateTime chulTime2 = new DateTime();//퇴근시간
            DateTime janTime = DateTime.Parse(sBASDT + " " + _NTBAB);//잔업시작시간

            if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("1"))//남직원. 일반
            {
                chulTime = DateTime.Parse(sBASDT + " " + _ILBAN1);
                chulTime2 = DateTime.Parse(sBASDT + " " + _ILBAN2);

                dGWKTM2 = 1.5;
            }
            else if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("2"))//여직원
            {
                chulTime = DateTime.Parse(sBASDT + " " + _YUKIC1);
                chulTime2 = DateTime.Parse(sBASDT + " " + _YUKIC2);

                dGWKTM2 = 0.0;
            }

            //출근시간 비교 출근시간 = 7:00
            DateTime.TryParse(sBASDT + " " + sWKINTM, out DateTime dtTime);

            //퇴근시간 비교 퇴근시간 = 17:30
            DateTime.TryParse(sBASDT + " " + sWKOTTM, out DateTime dtTime2);

            if (dtTime.Hour > _YAGAN)
            {
                //야간근무인 경우 퇴근 다음날
                dtTime2 = dtTime2.AddDays(1);
            }

            if (!string.IsNullOrEmpty(sWKINTM))
            {
                //sSRTGB = "11";//출근

                if (dtTime.Hour < _YAGAN && chulTime < dtTime) //야근조가 아니고(야근조: 18시 이후 출근), 지각
                {
                    TimeSpan ts = dtTime - chulTime;

                    //지각
                    //sSRTGB = "LN";
                    dGWKTM5 = Math.Round(((double)((ts.Hours * 60) + ts.Minutes) / 60), 1, MidpointRounding.AwayFromZero);

                    //#0001
                    double dTemp = dGWKTM5 - Math.Truncate(dGWKTM5);
                    if (dTemp < 0.5)
                    {
                        dGWKTM5 = dGWKTM5 - dTemp + 0.5;
                    }
                    else if (dTemp > 0.5)
                    {
                        dGWKTM5 = dGWKTM5 - dTemp + 1.0;
                    }
                }
                else if (dtTime.Hour > _YAGAN)
                {
                    //sWKGUB = "YA";//야간근무
                }
            }

            if (!string.IsNullOrEmpty(sWKOTTM))
            {
                //sENDGB = "21";//퇴근

                if (!string.IsNullOrEmpty(sWKINTM))
                {
                    //총근무시간
                    TimeSpan ts = dtTime2 - dtTime;
                    dAlWKBS = Math.Round((double)(((ts.Hours * 60) + ts.Minutes) / 60.0), 1, MidpointRounding.AwayFromZero);
                    dAlWKBS = Math.Round(dAlWKBS * 2, MidpointRounding.AwayFromZero) / 2;

                    if (dAlWKBS >= 8)
                    {
                        dWKBASE = 8;
                    }
                    else
                    {
                        dWKBASE = dAlWKBS;
                    }

                    //추가잔업 18시 이후 근무시간. 야간근무제외
                    if (dtTime2 > janTime && dtTime.Hour < _YAGAN)
                    {
                        TimeSpan tsJan = dtTime2 - janTime;

                        dGWKTM2 = 1.5;
                        dGWKTM3 = Math.Round((double)(((tsJan.Hours * 60) + tsJan.Minutes) / 60.0), 1, MidpointRounding.AwayFromZero);
                        dGWKTM3 = Math.Round(dGWKTM3 * 2, MidpointRounding.AwayFromZero) / 2;

                        if (dGWKTM3 > 0)
                        {
                            if (!string.IsNullOrEmpty(sWkgub) && sWkgub.Equals("DJ")) //당직
                            {
                                dGWKTM4 = dGWKTM3;
                                dGWKTM3 = 0;
                            }
                            else
                            {
                                //sENDGB = "JA";//잔업
                            }
                        }
                    }
                    //else if (dtTime2 >= chulTime2 && dtTime.Hour < _YAGAN) //추가잔업x 야간근무x
                    //{
                    //    dGWKTM3 = 0;
                    //    dGWKTM2 = dAlWKBS - 9;

                    //    if (dGWKTM2 > 1.5)
                    //    {
                    //        dGWKTM2 = 1.5;
                    //    }
                    //    else if (dGWKTM2 < 0)
                    //    {
                    //        dGWKTM2 = 0;
                    //    }
                    //    else
                    //    {
                    //        dGWKTM2 = Math.Round(dGWKTM2 * 2, MidpointRounding.AwayFromZero) / 2;
                    //    }
                    //}

                    //조퇴(야간제외) //수기입력한다함 2022-08-24
                    //if (dtTime2 < chulTime2 && dtTime.Hour < _YAGAN) //조퇴
                    //{
                    //    TimeSpan tsJote = chulTime2 - dtTime2;

                    //    dGWKTM6 = Math.Round(((double)((tsJote.Hours * 60) + tsJote.Minutes) / 60), 1, MidpointRounding.AwayFromZero);
                    //    dGWKTM6 = Math.Round(dGWKTM6 * 2, MidpointRounding.AwayFromZero) / 2;

                    //    if (dGWKTM6 > 0)
                    //    {
                    //        //sENDGB = "ET";//조퇴
                    //    }
                    //}
                }
            }

            //LkupStrgb.EditValue = sSRTGB;//출근구분
            //LkupEndgb.EditValue = sENDGB;//퇴근구분

            TxtWkbas.EditValue = dWKBASE;//근무시간
            TxtGWKTM2.EditValue = dGWKTM2;//기본잔업 1.5
            TxtGWKTM3.EditValue = dGWKTM3;//추가잔업 
            TxtGWKTM4.EditValue = dGWKTM4;//당직
            TxtGWKTM5.EditValue = dGWKTM5;//지각시간
            //TxtGWKTM6.EditValue = dGWKTM6;//조퇴시간
        }
        #endregion

        #region 외출관련 이벤트
        private void TxtGOINTM_Leave(object sender, EventArgs e)
        {
            string sTktm1 = TxtGOINTM.EditValue?.ToString();
            string sTktm2 = TxtGOOTTM.EditValue?.ToString();

            if (string.IsNullOrEmpty(sTktm1) || string.IsNullOrEmpty(sTktm2))
                return;

            DateTime dtTktm1 = DateTime.Parse(sTktm1);
            DateTime dtTktm2 = DateTime.Parse(sTktm2);

            TimeSpan resultTime = dtTktm2 - dtTktm1;

            double dGwktm7 = 0;
            dGwktm7 = Math.Round((double)(((resultTime.Hours * 60) + resultTime.Minutes)/ 60.0), 1, MidpointRounding.AwayFromZero);
            dGwktm7 = Math.Round(dGwktm7 * 2, MidpointRounding.AwayFromZero) / 2;

            TxtGWKTM7.EditValue = dGwktm7;
        }
        #endregion

        #region 연차관련 이벤트 사용x
        private void LkupStrgb_EditValueChanged(object sender, EventArgs e)
        {
            string sVal = LkupStrgb.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
                return;

            if (sVal.Equals("YG") || sVal.Equals("BG")) //연차(개인), 병가
            {
                TxtGWKTM1.EditValue = 1.0;
            }
            else if (sVal.Equals("YA")) //오전반차
            {
                TxtGWKTM1.EditValue = 0.5;
            }
        }

        private void LkupEndgb_EditValueChanged(object sender, EventArgs e)
        {
            string sVal = LkupStrgb.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
                return;

            if (sVal.Equals("YP")) //오후반차
            {
                TxtGWKTM1.EditValue = 0.5;
                TxtGWKTM2.EditValue = 0.0;
                TxtGWKTM3.EditValue = 0.0;
            }
        }
        #endregion
        
        private void lkupUsr_EditValueChanged(object sender, EventArgs e)
        {
            string sVal = lkupUsr.EditValue?.ToString();
            TxtEmpNo.EditValue = sVal;

            if (!string.IsNullOrEmpty(sVal))
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT GENDER_CD   ");
                strSql.AppendLine("   FROM HR_EMP_BASIS");
                strSql.AppendLine("  WHERE EMP_ID = '"+ sVal + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if(dt != null)
                {
                   sGENDER = dt.Rows[0]["GENDER_CD"]?.ToString();
                }
            }
        }
        
    }
}