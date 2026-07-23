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
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.Data;
using DevExpress.XtraGrid;

namespace AccAdm
{
    public partial class SC010F00 : DevExpress.XtraEditors.XtraForm
    {
        public SC010F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_SC010F00";

        private void SC010F00_Load(object sender, EventArgs e)
        {
            KeyPreview = true;
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            ComnEtcFunc.SetDateFromToValue(Dt_DateF, Dt_DateT);
            //설비가동률
            ComnGridFunc.GridStyleBasicSetting(GridViewGILO1);
            ComnGridFunc.GridStyleBasicSetting(GridViewWK1);
            //시간당생산량
            ComnGridFunc.GridStyleBasicSetting(GridViewMeChul);
            ComnGridFunc.GridStyleBasicSetting(GridViewGILO2);
            //작업공수
            ComnGridFunc.GridStyleBasicSetting(GridViewRetrGILO3);
            ComnGridFunc.GridStyleBasicSetting(GridViewRetrWK2);
            //ComnGridFunc.GridStyleBasicSetting(GridViewMeIp);
            //ComnGridFunc.GridStyleBasicSetting(GridViewMeChul2);

            Bt_Retr.PerformClick();
        }

        // 초기화(F1)
        private void Bt_Reset_Click(object sender, EventArgs e)
        {

        }

        // 조회(F5)
        private void Bt_Retr_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            KPIRetr();
            // P1
            if (xtraTabControl1.SelectedTabPageIndex == 0)
                Retr("P1");
            // P2
            else if (xtraTabControl1.SelectedTabPageIndex == 1)
                Retr("P2");
            // C
            else if (xtraTabControl1.SelectedTabPageIndex == 2)
                Retr("C");

            Cursor = Cursors.Default;
        }

        // 목표수정(F2)
        private void Bt_Update_Click(object sender, EventArgs e)
        {
            SC010F01 frm = new SC010F01();
            frm.Owner = this;
            frm.DataRowSendEvent += new SC010F01.SendDataHandler(KPIRefresh);
            frm.ShowDialog();
        }

        // 닫기
        private void Bt_Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        // KPI 성과지표 조회 메서드
        private void KPIRetr()
        {
            string sDateF = Dt_DateF.EditValue?.ToString().Substring(0, 7);
            string sDateT = Dt_DateT.EditValue?.ToString().Substring(0, 7);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            try
            {
                //KPI성과지표
                dicParams.Clear();
                dicParams.Add("CMD", "KPI_Retr");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetrKPI.DataSource = dt;

                DataTable dtKPI = (DataTable)GridRetrKPI.DataSource;

                //KPI성과지표 헌재, 이전 산출 조회 PASTR:이전, PRETR:현재, GRATE:달성률
                dicParams.Clear();
                dicParams.Add("CMD", "KPI_Detail");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            foreach (DataRow drKPI in dtKPI.Rows)
                            {
                                if (dr["KNO"].Equals(drKPI["KNO"]))
                                {
                                    string sPRETR = dr["PRETR"]?.ToString();
                                    string sGOALR = drKPI["GOALR"]?.ToString();
                                    string sPASTR = dr["PASTR"]?.ToString();

                                    if (string.IsNullOrEmpty(sPRETR))
                                    {
                                        return;
                                    }

                                    double dPRETR = 0;
                                    double dGOALR = 0;
                                    double dPASTR = 0;

                                    double.TryParse(sPRETR, out dPRETR);//현재
                                    double.TryParse(sGOALR, out dGOALR);//목표
                                    double.TryParse(sPASTR, out dPASTR);//이전

                                    drKPI["PASTR"] = dr["PASTR"];
                                    drKPI["PRETR"] = dr["PRETR"];

                                    #region 목표 대비 달성률
                                    /// 성장률 계산식?
                                    /// 증가 = (현재-이전) / (목표-이전) * 100
                                    /// 감소 = (이전-현재) / (이전-목표) * 100
                                    /// 달성률 계산식
                                    /// 증가 = (현재/목표) * 100
                                    /// 감소 = (목표/현재) * 100
                                    // P1, 증가
                                    if (dr["KNO"].Equals("1"))
                                        drKPI["GRATE"] = Math.Round(dPRETR / dGOALR * 100);
                                    // P2, 증가
                                    else if (dr["KNO"].Equals("2"))
                                        drKPI["GRATE"] = Math.Round(dPRETR / dGOALR * 100);
                                    // C, 감소
                                    else if (dr["KNO"].Equals("3"))
                                        drKPI["GRATE"] = Math.Round(dGOALR / dPRETR * 100);
                                    #endregion

                                    #region 이전 대비 달성률
                                    //// P1, 증가
                                    //if (dr["KNO"].Equals("1"))
                                    //    drKPI["GRATE"] = Math.Round(dPRETR / dPASTR * 100);
                                    //// P2, 증가
                                    //else if (dr["KNO"].Equals("2"))
                                    //    drKPI["GRATE"] = Math.Round(dPRETR / dPASTR * 100);
                                    //// C, 감소
                                    //else if (dr["KNO"].Equals("3"))
                                    //    drKPI["GRATE"] = Math.Round(dPASTR / dPRETR * 100);
                                    #endregion

                                    drKPI["GRATE"] = Math.Round(Convert.ToDouble(drKPI["GRATE"]), 2).ToString() + "%";
                                }
                            }
                        }
                        GridRetrKPI.DataSource = dtKPI;
                    }
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        // KPI 성과지표 산출근거 그리드 조회 메서드
        private void Retr(string GB)
        {
            string sDateF = Dt_DateF.EditValue?.ToString().Substring(0, 7);
            string sDateT = Dt_DateT.EditValue?.ToString().Substring(0, 7);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            // P1 : 설비가동률 조회
            if (GB.Equals("P1"))
            {
                //길로틴작업시간
                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetrGILO1.DataSource = dt1;

                //길로틴 작업자 근무시간
                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetrWK1.DataSource = dt2;
            }
            // P2 : 시간당 생산량
            else if (GB.Equals("P2"))
            {
                //길로틴작업시간
                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetrGILO2.DataSource = dt1;

                //생산량(매출량)
                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                DataTable dt3 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridMeChul.DataSource = dt3;
            }
            // C : 작업공수
            else if (GB.Equals("C"))
            {
                ////매입량
                //dicParams.Clear();
                //dicParams.Add("CMD", "LIST4");
                //dicParams.Add("DATE_F", sDateF);
                //dicParams.Add("DATE_T", sDateT);

                //DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                //GridMeIp.DataSource = dt1;

                ////작업량(매출량)
                //dicParams.Clear();
                //dicParams.Add("CMD", "LIST3");
                //dicParams.Add("DATE_F", sDateF);
                //dicParams.Add("DATE_T", sDateT);

                //DataTable dt3 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                //GridMeChul2.DataSource = dt3;

                //길로틴 작업자 근무시간
                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetrWK2.DataSource = dt2;

                //길로틴작업시간
                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("DATE_F", sDateF);
                dicParams.Add("DATE_T", sDateT);

                DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetrGILO3.DataSource = dt1;
            }
        }

        private void KPIRefresh(string sVal)
        {
            KPIRetr();
        }

        #region [ Row 스타일 ]
        private void GridViewRetrKPI_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetrKPI_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }
        #endregion

        // 탭 변경 시 자동 조회 이벤트
        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            // P1
            if (xtraTabControl1.SelectedTabPageIndex == 0)
                Retr("P1");
            // P2
            else if (xtraTabControl1.SelectedTabPageIndex == 1)
                Retr("P2");
            // C
            else if (xtraTabControl1.SelectedTabPageIndex == 2)
                Retr("C");

        }

        #region [ 키 이벤트 ]
        private void SC010F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5) { Bt_Retr.PerformClick(); }
            else if (e.KeyCode == Keys.F1) { Bt_Reset.PerformClick(); }
            else if (e.KeyCode == Keys.F2) { Bt_Update.PerformClick(); }
            else if (e.KeyCode == Keys.Escape) {  }
        }
        #endregion

        private void Dt_DateT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Bt_Retr.PerformClick();
        }

        //길로틴 작업시간 foot합산 
        double tot_time = 0;
        private void GridViewGILO1_CustomSummaryCalculate(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            GridSummaryItem gridSummaryItem = e.Item as GridSummaryItem;

            GridView gridView = sender as GridView;

            switch (e.SummaryProcess)
            {
                //calculation entry point
                case CustomSummaryProcess.Start:
                    tot_time = 0;
                    break;
                //consequent calculations
                case CustomSummaryProcess.Calculate:
                    string sVal = gridView.GetRowCellValue(e.RowHandle, gridView.Columns["DIFFTM"])?.ToString();
                    double dVal = 0;

                    int i = 0;
                    if (!string.IsNullOrEmpty(sVal))
                    {
                        string[] arr = sVal.Split(':');

                        foreach (string str in arr)
                        {
                            if (double.TryParse(str, out double dResult))
                            {
                                switch (i)
                                {
                                    case 0:
                                        dVal += dResult * 3600;
                                        break;
                                    case 1:
                                        dVal += dResult * 60;
                                        break;
                                    case 2:
                                        dVal += dResult;
                                        break;
                                }
                            }
                            i++;
                        }

                    }

                    tot_time += dVal;

                    break;
                //final summary value
                case CustomSummaryProcess.Finalize:
                    //시, 분, 초 선언
                    double hours, minute, second;

                    //시간공식
                    hours = Math.Truncate(tot_time / 3600);//시 공식
                    minute = Math.Truncate(tot_time % 3600 / 60);//분을 구하기위해서 입력되고 남은값에서 또 60을 나눈다.
                    second = Math.Truncate(tot_time % 3600 % 60);//마지막 남은 시간에서 분을 뺀 나머지 시간을 초로 계산함

                    string sTempVal1 = Dt_DateF.EditValue?.ToString(); //감리위해 임시로 더하기
                    string sTempVal2 = Dt_DateT.EditValue?.ToString();

                    //DateTime temp0 = new DateTime(2023, 01, 01);
                    //DateTime.TryParse(sTempVal1, out DateTime temp1);
                    //DateTime.TryParse(sTempVal2, out DateTime temp2);
                    //
                    //if (temp1<= temp0 && temp2>= temp0)
                    //{
                    //    hours += 110;
                    //}
                    //e.TotalValue = string.Format("{0}:{1}:{2}", hours, minute, second);
                    e.TotalValue = string.Format("{0}", hours);
                    break;
            }
        }
    }
}