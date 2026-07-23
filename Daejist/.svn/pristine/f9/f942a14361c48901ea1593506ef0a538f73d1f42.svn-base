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
using System.IO;
using DevExpress.XtraGrid;
using DevExpress.Data;
using System.Text.RegularExpressions;

namespace AccAdm
{
    public partial class EQ001F00 : DevExpress.XtraEditors.XtraForm
    {
        public EQ001F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_EQ001F00";
        private string PROCEDURE_AL = "DP_MONIALARM";

        private void EQ001F00_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();

            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.gp_SetColorFocused(layoutControl3);
            ComnEtcFunc.gp_SetColorFocused(layoutControl5);
            ComnEtcFunc.gp_SetColorFocused(layoutControl7);

            ComnEtcFunc.SetDateFromToValue(DateFromT, DateToT);
            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);
            ComnEtcFunc.SetDateFromToValue(DateFrom2, DateTo2);
            ComnEtcFunc.SetDateFromToValue(DateFrom3, DateTo3);
            ComnEtcFunc.SetDateToValue(DateFrom4);
            ComnEtcFunc.SetDateFromToValue(DateFrom5, DateTo5);

            //ComnGridFunc.GridStyleBasicSetting(GridViewRetr2);
            //ComnGridFunc.GridStyleBasicSetting(GridViewRetr3);
            //ComnGridFunc.GridStyleBasicSetting(GridViewRetr4);

            SetLoadFormLayout();

            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr, GridViewRetr2, GridViewRetr3, GridViewRetr4, GridViewErrorMsg };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
        }
        #endregion

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if(xtraTabControl1.SelectedTabPageIndex == 0)
            {
                GetRetrData();
                SetInit();
            }
            else if(xtraTabControl1.SelectedTabPageIndex == 1)
            {
                GetRetr2DataN();
            }
            else if (xtraTabControl1.SelectedTabPageIndex == 2)
            {
                GetRetr3Data();
            }
            else if (xtraTabControl1.SelectedTabPageIndex == 3)
            {
                GetRetr4Data();
            }
            else if (xtraTabControl1.SelectedTabPageIndex == 4)
            {
                GetRetr5Data();
            }
        }

        private void SetInit()
        {
            RepoChkEditErrorImg.ImageOptions.ImageChecked = Properties.Resources.warning__1_;
            RepoChkEditErrorImg.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.UserDefined;

        }

        #region [ 설비 I/F 조회 ]
        private void GetRetrData()
        {
            try
            {
                string sFdate = DateFromT.EditValue?.ToString().Substring(0, 10);
                string sTdate = DateToT.EditValue?.ToString().Substring(0, 10);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null && dt.Rows.Count > 0)
                {
                    GridRetr.Focus();

                    int iDayCharge = 1;
                    foreach(DataRow row in dt.Rows)
                    {
                        string sGub = row["GUB"]?.ToString();
                        if (sGub.Equals("S"))
                        {
                            iDayCharge = 1;
                        }

                        row["GHARG2"] = iDayCharge.ToString();

                        iDayCharge++;
                    }
                }
                else
                {
                    DateFromT.Focus();
                }

                GridRetr.DataSource = dt;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region [ 작업자별 생산내역 조회 ]
        private void GetRetr2Data()
        {
            try
            {
                string sFdate = DateFromT.EditValue?.ToString().Substring(0, 10);
                string sTdate = DateToT.EditValue?.ToString().Substring(0, 10);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                dt.Columns.Add("WORKT", typeof(string));

                if (dt != null && dt.Rows.Count > 0)
                {
                    double xWKTIM_dt = 0;
                    string xWORKT = string.Empty;

                    GridRetr2.Focus();
                    DataTable WkTb = GetWorkTimeDataTable(sFdate, sTdate);

                    if(WkTb != null && WkTb.Rows.Count > 0)
                    {
                        foreach (DataRow rowDt in dt.Rows)
                        {
                            //더한 근무시간 값 dt에 삽입
                            string sWorker_dt = rowDt["WORKER"]?.ToString();
                            string sEMP_NM_dt = rowDt["EMP_NM"]?.ToString();
                            string sDIFF2 = rowDt["DIFF2"]?.ToString();//작업시간

                            double dDIFF2 = 0;
                            double dWKTIM_dt = 0;

                            double.TryParse(sDIFF2, out dDIFF2);

                            if(string.IsNullOrEmpty(sEMP_NM_dt) || sEMP_NM_dt.Equals("소계"))
                            {
                                rowDt["WKTIM"] = xWKTIM_dt;
                                rowDt["WORKT"] = xWORKT;
                                continue;
                            }
                            //if (string.IsNullOrEmpty(sEMP_NM_dt) || sEMP_NM_dt.Equals("합계"))
                            //{
                            //    rowDt["WKTIM"] = xWKTIM_dt;
                            //    rowDt["WORKT"] = xWORKT;
                            //    continue;
                            //}

                            foreach (DataRow rowWk in WkTb.Rows)
                            {
                                //같은 작업자 근무시간 값 더하기
                                string sWorker_Wk = rowWk["WORKER"]?.ToString();

                                if (sWorker_dt.Equals(sWorker_Wk))
                                {
                                    if(double.TryParse(rowWk["WKTIM"]?.ToString(), out double dWKTIM_Wk))
                                    {
                                        dWKTIM_dt = dWKTIM_Wk;
                                    }
                                }
                            }
                            ///가동률 계산해야함
                            rowDt["DIFF2"] = dDIFF2.ToString("n1");

                            int seconds = Convert.ToInt32(dWKTIM_dt); // 변환할 초

                            int hour = seconds / 3600;
                            int min = (seconds % 3600) / 60;
                            int sec = (seconds % 3600) % 60;
                            
                            string shour = Convert.ToInt32(hour).ToString("00");
                            string smin = Convert.ToInt32(min).ToString("00");
                            string ssec = Convert.ToInt32(sec).ToString("00");

                            string sWORKT = string.Concat(shour + ":" + smin + ":" + ssec);

                            double dGADON = dDIFF2 / dWKTIM_dt * 100;
                            string sGADON = string.Format("{0:0.#}", dGADON);

                            rowDt["WKTIM"] = dWKTIM_dt;
                            rowDt["WORKT"] = sWORKT;
                            rowDt["GADON"] = sGADON + "%";

                            xWORKT = sWORKT;
                            xWKTIM_dt = dWKTIM_dt;


                        }
                    }
                }
                else
                {
                    DateFromT.Focus();
                }

                GridRetr2.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GetRetr2DataN()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                string sFdate = DateFromT.EditValue?.ToString().Substring(0, 10);
                string sTdate = DateToT.EditValue?.ToString().Substring(0, 10);

                DataTable dt = new DataTable();
                DataTable dtWkAndGb = GetWorkAndGbTimeDataTableNEWNEW(sFdate, sTdate);//작업자,품목 근무시간

                if (dtWkAndGb != null && dtWkAndGb.Rows.Count > 0)
                {
                    string sJSON1 = ComnEtcFunc.DataTableToJsonWithNewtonSoft(dtWkAndGb);

                    dicParams.Clear();
                    dicParams.Add("CMD", "LIST8TS");
                    dicParams.Add("FDATE", sFdate);
                    dicParams.Add("TDATE", sTdate);
                    dicParams.Add("JSON1", sJSON1);

                    dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    GridRetr2.Focus();
                    foreach (DataRow rowDt in dt.Rows)
                    {
                        //더한 근무시간 값 dt에 삽입
                        string sWorker_dt = rowDt["WORKER"]?.ToString();
                        string sEMP_NM_dt = rowDt["EMP_NM"]?.ToString();
                        string sMGUBUN_dt = rowDt["MGUBUN"]?.ToString();
                        string sWKTIMT = rowDt["WKTIMT"]?.ToString();//작업시간

                        double dWKTIMT = 0;
                        double dWKTIM_dt = 0;
                        double dWKTIM_dtt = 0;

                        double.TryParse(sWKTIMT, out dWKTIMT);

                        if (sEMP_NM_dt.Equals("소계"))
                        {
                            foreach (DataRow rowWk in dt.Rows)
                            {
                                //같은 작업자 근무시간 값 더하기
                                string sWorker_Wk = rowWk["WORKER"]?.ToString();

                                if (sWorker_dt.Equals(sWorker_Wk))
                                {
                                    if (double.TryParse(rowWk["WKTIMT"]?.ToString(), out double dWKTIM_Wk))
                                    {
                                        dWKTIM_dt += dWKTIM_Wk;
                                    }
                                }
                            }
                            //소계 합
                            rowDt["WKTIMT"] = dWKTIM_dt;

                            //소계합을 00:00:00으로 변환
                            int wktim = Convert.ToInt32(dWKTIM_dt);
                            string timeString = string.Format("{0:00}:{1:00}:{2:00}", 
                                wktim / 3600,
                                (wktim % 3600) / 60,
                                (wktim % 3600) % 60 
                            );

                            string result = timeString; 

                            rowDt["WKTIM"] = result;


                            string dDIFF = rowDt["DIFF"]?.ToString();
                            
                            Match match = Regex.Match(dDIFF, @"(\d+):(\d+):(\d+)");

                            
                            int hours = int.Parse(match.Groups[1].Value);
                            int minutes = int.Parse(match.Groups[2].Value);
                            int seconds = int.Parse(match.Groups[3].Value);
                            
                            int totalSeconds = hours * 3600 + minutes * 60 + seconds;

                            
                            double dwww = totalSeconds / dWKTIM_dt * 100;
                            string wwww = dwww.ToString("0.0") + "%";

                            rowDt["GADON"] = wwww;

                        }
                        else if (sEMP_NM_dt.Equals("합계"))
                        {
                            foreach (DataRow rowWkt in dt.Rows)
                            {
                                //같은 작업자 근무시간 값 더하기
                                string tt = "소계";
                                string sWorker_Wk = rowWkt["EMP_NM"]?.ToString();

                                if (tt.Equals(sWorker_Wk))
                                {
                                    if (double.TryParse(rowWkt["WKTIMT"]?.ToString(), out double dWKTIM_Wkt))
                                    {
                                        dWKTIM_dtt += dWKTIM_Wkt;
                                    }
                                }
                            }
                            rowDt["WKTIMT"] = dWKTIM_dtt;

                            int wktim = Convert.ToInt32(dWKTIM_dtt);
                            string timeString = string.Format("{0:00}:{1:00}:{2:00}",
                                wktim / 3600,
                                (wktim % 3600) / 60,
                                (wktim % 3600) % 60
                            );

                            string result = timeString; 

                            rowDt["WKTIM"] = result;

                            string dDIFF = rowDt["DIFF"]?.ToString();

                            Match match = Regex.Match(dDIFF, @"(\d+):(\d+):(\d+)");


                            int hours = int.Parse(match.Groups[1].Value);
                            int minutes = int.Parse(match.Groups[2].Value);
                            int seconds = int.Parse(match.Groups[3].Value);

                            int totalSeconds = hours * 3600 + minutes * 60 + seconds;


                            double dwww = totalSeconds / dWKTIM_dtt * 100;
                            string wwww = dwww.ToString("0.0") + "%";

                            rowDt["GADON"] = wwww;

                        }
                        else
                        {
                            continue;
                        }
                    }


                }
                else
                {
                    DateFromT.Focus();
                }

                GridRetr2.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }



        private void GridViewRetr2_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sEmpNm1 = GridViewRetr2.GetRowCellValue(e.RowHandle1, "EMP_NM")?.ToString();
            string sEmpNm2 = GridViewRetr2.GetRowCellValue(e.RowHandle2, "EMP_NM")?.ToString();

            var val1 = GridViewRetr2.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = GridViewRetr2.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("EMP_NM") || e.Column.FieldName.Equals("DIFF2") || e.Column.FieldName.Equals("WKTIM") || e.Column.FieldName.Equals("GADON"))
            {
                e.Merge = (sEmpNm1.Equals(sEmpNm2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
            }
        }
        #endregion

        #region [ 품목별 생산내역 조회 ]
        private void GetRetr3Data()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                string sFdate = DateFromT.EditValue?.ToString().Substring(0, 10);
                string sTdate = DateToT.EditValue?.ToString().Substring(0, 10);

                DataTable dt = new DataTable();
                DataTable dtWkAndGb = GetWorkAndGbTimeDataTableNEW(sFdate, sTdate);//작업자,품목 근무시간

                if (dtWkAndGb != null && dtWkAndGb.Rows.Count > 0)
                {
                    string sJSON1 = ComnEtcFunc.DataTableToJsonWithNewtonSoft(dtWkAndGb);

                    dicParams.Clear();
                    dicParams.Add("CMD", "LIST8");
                    dicParams.Add("FDATE", sFdate);
                    dicParams.Add("TDATE", sTdate);
                    dicParams.Add("JSON1", sJSON1);

                    dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    GridRetr3.Focus();
                    foreach (DataRow rowDt in dt.Rows)
                    {
                        //더한 근무시간 값 dt에 삽입
                        string sWorker_dt = rowDt["WORKER"]?.ToString();
                        string sEMP_NM_dt = rowDt["EMP_NM"]?.ToString();
                        string sMGUBUN_dt = rowDt["MGUBUN"]?.ToString();
                        string sMGUBUNTS_dt = rowDt["MGUBUNTS"]?.ToString();
                        string sWKTIMT = rowDt["WKTIMT"]?.ToString();//작업시간

                        double dWKTIMT = 0;
                        double dWKTIM_dt = 0;
                        double dWKTIM_dtt = 0;

                        double.TryParse(sWKTIMT, out dWKTIMT);

                        if (sMGUBUN_dt.Equals("소계"))
                        {
                            foreach (DataRow rowWk in dt.Rows)
                            {
                                string sMGUBUN_Wk = rowWk["MGUBUNTS"]?.ToString();

                                if (sMGUBUNTS_dt.Equals(sMGUBUN_Wk))
                                {
                                    if (double.TryParse(rowWk["WKTIMT"]?.ToString(), out double dWKTIM_Wk))
                                    {
                                        dWKTIM_dt += dWKTIM_Wk;
                                    }
                                }
                            }
                            //소계 합
                            rowDt["WKTIMT"] = dWKTIM_dt;

                            //소계합을 00:00:00으로 변환
                            int wktim = Convert.ToInt32(dWKTIM_dt);
                            string timeString = string.Format("{0:00}:{1:00}:{2:00}",
                                wktim / 3600,
                                (wktim % 3600) / 60,
                                (wktim % 3600) % 60
                            );

                            string result = timeString;

                            rowDt["WKTIM"] = result;


                            string dDIFF = rowDt["DIFF"]?.ToString();

                            Match match = Regex.Match(dDIFF, @"(\d+):(\d+):(\d+)");


                            int hours = int.Parse(match.Groups[1].Value);
                            int minutes = int.Parse(match.Groups[2].Value);
                            int seconds = int.Parse(match.Groups[3].Value);

                            int totalSeconds = hours * 3600 + minutes * 60 + seconds;


                            double dwww = totalSeconds / dWKTIM_dt * 100;
                            string wwww = dwww.ToString("0.0") + "%";

                            rowDt["GADON"] = wwww;

                        }
                        else if (sMGUBUN_dt.Equals("합계"))
                        {
                            foreach (DataRow rowWkt in dt.Rows)
                            {
                                //같은 작업자 근무시간 값 더하기
                                string tt = "소계";
                                string sMGUBUN_Wk = rowWkt["MGUBUN"]?.ToString();

                                if (tt.Equals(sMGUBUN_Wk))
                                {
                                    if (double.TryParse(rowWkt["WKTIMT"]?.ToString(), out double dWKTIM_Wkt))
                                    {
                                        dWKTIM_dtt += dWKTIM_Wkt;
                                    }
                                }
                            }
                            rowDt["WKTIMT"] = dWKTIM_dtt;

                            int wktim = Convert.ToInt32(dWKTIM_dtt);
                            string timeString = string.Format("{0:00}:{1:00}:{2:00}",
                                wktim / 3600,
                                (wktim % 3600) / 60,
                                (wktim % 3600) % 60
                            );

                            string result = timeString;

                            rowDt["WKTIM"] = result;

                            string dDIFF = rowDt["DIFF"]?.ToString();

                            Match match = Regex.Match(dDIFF, @"(\d+):(\d+):(\d+)");


                            int hours = int.Parse(match.Groups[1].Value);
                            int minutes = int.Parse(match.Groups[2].Value);
                            int seconds = int.Parse(match.Groups[3].Value);

                            int totalSeconds = hours * 3600 + minutes * 60 + seconds;


                            double dwww = totalSeconds / dWKTIM_dtt * 100;
                            string wwww = dwww.ToString("0.0") + "%";

                            rowDt["GADON"] = wwww;

                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                else
                {
                    DateFromT.Focus();
                }

                GridRetr3.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr3_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sEmpNm1 = GridViewRetr3.GetRowCellValue(e.RowHandle1, "MGUBUN")?.ToString();
            string sEmpNm2 = GridViewRetr3.GetRowCellValue(e.RowHandle2, "MGUBUN")?.ToString();

            var val1 = GridViewRetr3.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = GridViewRetr3.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("MGUBUN") || e.Column.FieldName.Equals("DIFF2") || e.Column.FieldName.Equals("WKTIM") || e.Column.FieldName.Equals("GADON"))
            {
                e.Merge = (sEmpNm1.Equals(sEmpNm2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
            }
        }
        #endregion

        #region [ 작업자별 일일 생산내역 조회 ]
        private void GetRetr4Data()
        {
            try
            {
                string sFdate = DateFrom4.EditValue?.ToString().Substring(0, 10);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sFdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                DataTable newDt = dt.Clone();
                DataTable dtSum = new DataTable();
                
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 1;
                    int rowCnt = 1;
                    string sPrevCd = string.Empty;
                    string sPrevGB = string.Empty;
                    DataRow newRow = null;

                    //차이 초 합
                    double dSec = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        string sWORKER = row["WORKER"]?.ToString();
				        string sEMP_NM = row["EMP_NM"]?.ToString();
				        string sMGUBUN = row["MGUBUN"]?.ToString();
				        string sSTRDT  = row["STRDT"]?.ToString();
                        string sENDDT  = row["ENDDT"]?.ToString();
                        string sCHARG =  row["CHARG"]?.ToString();
                        string sDIFF = row["DIFF"]?.ToString();

                        if (sPrevCd.Equals(sWORKER) && sPrevGB.Equals(sMGUBUN))//이전 작업자와 품목이 같을때
                        {
                            
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["MGUBUN"] = sMGUBUN;
                            //newRow["STRDT"] = sSTRDT;
                            newRow["ENDDT"] = sENDDT;
                            newRow["CHARG"] = i;

                            if (double.TryParse(sDIFF, out double result))
                            {
                                dSec += result;
                            }

                            newRow["DIFF"] = dSec;

                            if(rowCnt == dt.Rows.Count)
                            {
                                double dAvgSec = Math.Round(dSec / i);

                                TimeSpan t = TimeSpan.FromSeconds(dAvgSec);
                                string str = string.Format("{0:D2}:{1:D2}:{2:D2}",
                                                           t.Hours,
                                                           t.Minutes,
                                                           t.Seconds);

                                newRow["AVGTM"] = str;

                                newDt.Rows.Add(newRow);
                            }
                        }
                        else //이전 작업자 또는 품목이 다를때
                        {
                            if (newRow != null)
                            {
                                double dAvgSec = Math.Round(dSec / (i-1));

                                TimeSpan t = TimeSpan.FromSeconds(dAvgSec);
                                string str = string.Format("{0:D2}:{1:D2}:{2:D2}",
                                                           t.Hours,
                                                           t.Minutes,
                                                           t.Seconds);

                                newRow["AVGTM"] = str;

                                newDt.Rows.Add(newRow);
                                dSec = 0;
                            }

                            i = 1;
                            sPrevCd = sWORKER;
                            sPrevGB = sMGUBUN;
                            newRow = newDt.NewRow();

                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["MGUBUN"] = sMGUBUN;
                            newRow["STRDT"]  = sSTRDT;
                            newRow["ENDDT"]  = sENDDT;
                            newRow["CHARG"] = i;
                            newRow["DIFF"] = dSec;

                            if (double.TryParse(sDIFF, out double result))
                            {
                                dSec = result;
                            }
                        }
                        i++;
                        rowCnt++;
                    }

                    if(newDt != null && newDt.Rows.Count > 0)
                    {
                        DataTable WkTb = GetWorkTimeDataTable(sFdate, sFdate);

                        string sJSON1 = ComnEtcFunc.DataTableToJsonWithNewtonSoft(newDt);
                        string sJSON2 = ComnEtcFunc.DataTableToJsonWithNewtonSoft(WkTb);

                        dicParams.Clear();
                        dicParams.Add("CMD", "LIST6");
                        dicParams.Add("JSON1", sJSON1);
                        dicParams.Add("JSON2", sJSON2);

                        dtSum = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                    }

                    GridRetr4.Focus();
                }
                else
                {
                    DateFrom4.Focus();
                }

                GridRetr4.DataSource = dtSum;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr4_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sEmpNm1 = GridViewRetr4.GetRowCellValue(e.RowHandle1, "EMP_NM")?.ToString();
            string sEmpNm2 = GridViewRetr4.GetRowCellValue(e.RowHandle2, "EMP_NM")?.ToString();

            var val1 = GridViewRetr4.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = GridViewRetr4.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("EMP_NM") || e.Column.FieldName.Equals("DIFF2") || e.Column.FieldName.Equals("WKTIM") || e.Column.FieldName.Equals("GADON"))
            {
                e.Merge = (sEmpNm1.Equals(sEmpNm2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
            }
        }
        #endregion

        #region [ 근무시간 ]
        /// <summary>
        /// 다음 근무자까지의 근무시간데이터 작업자별 ex) 김병진 1시간, 유진균1시간, 김병진3시간 ...
        /// </summary>
        /// <param name="sFdate">시작일</param>
        /// <param name="sTdate">종료일</param>
        /// <returns>DataTable</returns>
        private DataTable GetWorkTimeDataTable(string sFdate, string sTdate)
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                DataTable newDt = new DataTable();

                newDt.Columns.Add("WORKER");//작업자코드
                newDt.Columns.Add("EMP_NM");//작업자명
                newDt.Columns.Add("TDATE");//작업일
                newDt.Columns.Add("STRDT");//시작시간
                newDt.Columns.Add("ENDDT");//종료시간
                newDt.Columns.Add("WKTIM");//근무시간

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sSTRDTN = string.Empty;
                    string sTDATEN = string.Empty;
                    string sPrevCd = "1";
                    string sPreTdate = string.Empty;
                    int rowCnt = 0;
                    
                    DataRow newRow = null;
                    foreach (DataRow row in dt.Rows)
                    {
                        string sWORKER = row["WORKER"]?.ToString();
                        string sEMP_NM = row["EMP_NM"]?.ToString();
                        string sTDATE = row["TDATE"]?.ToString();
                        string sSTRDT = row["STRDT"]?.ToString();
                        string sENDDT = row["ENDDT"]?.ToString();
                        double dWKTIM = 0;
                        
                        if (sPrevCd.Equals(sWORKER))//이전 일자와 작업자 이름이 같을때
                        {

                            if (!string.IsNullOrEmpty(sSTRDTN))
                            {
                                if (DateTime.TryParse(sTDATE + " " + sSTRDT, out DateTime stDate))
                                {
                                    if (DateTime.TryParse(sTDATEN + " " + sSTRDTN, out DateTime stDateN))
                                    {
                                        if (stDate > stDateN)
                                        {
                                            sSTRDT = sSTRDTN;
                                        }
                                    }
                                }
                            }

                            if (DateTime.TryParse(sTDATE + " " + sSTRDT, out DateTime sDate))
                            {
                                if (DateTime.TryParse(sTDATE + " " + sENDDT, out DateTime eDate))
                                {
                                    if (sDate > eDate)
                                    {
                                        sDate = sDate.AddDays(-1);
                                    }

                                    TimeSpan dateDiff = eDate - sDate;
                                    double diffTotalSeconds = dateDiff.TotalSeconds;

                                    dWKTIM = diffTotalSeconds;
                                }
                            }

                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["ENDDT"] = sENDDT;
                            double newWkTim = 0;
                            double.TryParse(newRow["WKTIM"]?.ToString(), out newWkTim);
                            //newRow["WKTIM"] = newWkTim + dWKTIM;
                            newRow["WKTIM"] = dWKTIM;
                            sSTRDTN = sSTRDT;

                        }
                        else //이전 작업자가 다를때
                        {
                            if (newRow != null)
                            {
                                newDt.Rows.Add(newRow);
                            }

                            sPrevCd = sWORKER;
                            sPreTdate = sTDATE;
                            newRow = newDt.NewRow();

                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["STRDT"] = sSTRDT;
                            newRow["ENDDT"] = sENDDT;
                            newRow["WKTIM"] = dWKTIM;

                            //if (!string.IsNullOrEmpty(sSTRDTN))
                            //{
                            //    sSTRDTN = sSTRDT;
                            //}
                            sSTRDTN = sSTRDT;
                            sTDATEN = sTDATE;
                        }
                        rowCnt++;

                        if(rowCnt == dt.Rows.Count)//마지막 행 추가
                        {
                            if(newRow != null)
                                newDt.Rows.Add(newRow);
                        }
                    }

                    if(newDt != null && newDt.Rows.Count > 0)
                    {
                        //작업자 근무시간 총합
                        string sJSON = ComnEtcFunc.DataTableToJsonWithNewtonSoft(newDt);

                        if (string.IsNullOrEmpty(sJSON))
                            return null;

                        dicParams.Clear();
                        dicParams.Add("CMD", "LIST5");
                        dicParams.Add("JSON1", sJSON);

                        DataTable sumDt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                        return sumDt;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
        }

        private DataTable GetWorkAndGbTimeDataTableNEW(string sFdate, string sTdate)
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                DataTable newDt = new DataTable();

                newDt.Columns.Add("WORKER");//작업자코드
                newDt.Columns.Add("EMP_NM");//작업자명
                newDt.Columns.Add("MGUBUN");//품목명
                newDt.Columns.Add("TDATE");//작업일
                newDt.Columns.Add("STRDT");//시작시간
                newDt.Columns.Add("ENDDT");//종료시간
                newDt.Columns.Add("WKTIM");//근무시간

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sSTRDTN = string.Empty;
                    string sTDATEN = string.Empty;
                    string sPrevCd = string.Empty;
                    string sPrevWK = string.Empty;
                    string sPreTdate = string.Empty;
                    int rowCnt = 0;

                    DataRow newRow = null;
                    foreach (DataRow row in dt.Rows)
                    {
                        string sWORKER = row["WORKER"]?.ToString();
                        string sEMP_NM = row["EMP_NM"]?.ToString();
                        string sMGUBUN = row["MGUBUN"]?.ToString();
                        string sTDATE = row["TDATE"]?.ToString();
                        string sSTRDT = row["STRDT"]?.ToString();
                        string sENDDT = row["ENDDT"]?.ToString();
                        double dWKTIM = 0;

                        if (sPrevCd.Equals(sWORKER) && sPrevWK.Equals(sMGUBUN))
                        {

                            if (!string.IsNullOrEmpty(sSTRDTN))
                            {
                                if (DateTime.TryParse(sTDATE + " " + sSTRDT, out DateTime stDate))
                                {
                                    if (DateTime.TryParse(sTDATEN + " " + sSTRDTN, out DateTime stDateN))
                                    {
                                        if (stDate > stDateN)
                                        {
                                            sSTRDT = sSTRDTN;
                                        }
                                    }
                                }
                            }

                            if (DateTime.TryParse(sTDATE + " " + sSTRDT, out DateTime sDate))
                            {
                                if (DateTime.TryParse(sTDATE + " " + sENDDT, out DateTime eDate))
                                {
                                    if (sDate > eDate)
                                    {
                                        sDate = sDate.AddDays(-1);
                                    }

                                    TimeSpan dateDiff = eDate - sDate;
                                    double diffTotalSeconds = dateDiff.TotalSeconds;

                                    dWKTIM = diffTotalSeconds;
                                }
                            }

                            newRow["MGUBUN"] = sMGUBUN;
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["ENDDT"] = sENDDT;
                            double newWkTim = 0;
                            double.TryParse(newRow["WKTIM"]?.ToString(), out newWkTim);
                            //newRow["WKTIM"] = newWkTim + dWKTIM;
                            newRow["WKTIM"] = dWKTIM;
                            sSTRDTN = sSTRDT;

                        }
                        else //이전 
                        {
                            if (newRow != null)
                            {
                                newDt.Rows.Add(newRow);
                            }

                            sPrevCd = sWORKER;
                            sPrevWK = sMGUBUN;
                            sPreTdate = sTDATE;
                            newRow = newDt.NewRow();

                            newRow["MGUBUN"] = sMGUBUN;
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["STRDT"] = sSTRDT;
                            newRow["ENDDT"] = sENDDT;
                            //newRow["WKTIM"] = dWKTIM;
                            newRow["WKTIM"] = row["DIFF"];

                            //if (!string.IsNullOrEmpty(sSTRDTN))
                            //{
                            //    sSTRDTN = sSTRDT;
                            //}
                            sSTRDTN = sSTRDT;
                            sTDATEN = sTDATE;
                        }
                        rowCnt++;

                        if (rowCnt == dt.Rows.Count)//마지막 행 추가
                        {
                            if (newRow != null)
                                newDt.Rows.Add(newRow);
                        }
                    }

                    if (newDt != null && newDt.Rows.Count > 0)
                    {
                        //작업자 품목별 근무시간 총합
                        string sJSON = ComnEtcFunc.DataTableToJsonWithNewtonSoft(newDt);

                        if (string.IsNullOrEmpty(sJSON))
                            return null;

                        dicParams.Clear();
                        dicParams.Add("CMD", "LIST7");
                        dicParams.Add("JSON1", sJSON);

                        DataTable sumDt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                        return sumDt;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
        }
        private DataTable GetWorkAndGbTimeDataTableNEWNEW(string sFdate, string sTdate)
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                DataTable newDt = new DataTable();

                newDt.Columns.Add("WORKER");//작업자코드
                newDt.Columns.Add("EMP_NM");//작업자명
                newDt.Columns.Add("MGUBUN");//품목명
                newDt.Columns.Add("TDATE");//작업일
                newDt.Columns.Add("STRDT");//시작시간
                newDt.Columns.Add("ENDDT");//종료시간
                newDt.Columns.Add("WKTIM");//근무시간

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sSTRDTN = string.Empty;
                    string sTDATEN = string.Empty;
                    string sPrevCd = string.Empty;
                    string sPrevWK = string.Empty;
                    string sPreTdate = string.Empty;
                    int rowCnt = 0;

                    DataRow newRow = null;
                    foreach (DataRow row in dt.Rows)
                    {
                        string sWORKER = row["WORKER"]?.ToString();
                        string sEMP_NM = row["EMP_NM"]?.ToString();
                        string sMGUBUN = row["MGUBUN"]?.ToString();
                        string sTDATE = row["TDATE"]?.ToString();
                        string sSTRDT = row["STRDT"]?.ToString();
                        string sENDDT = row["ENDDT"]?.ToString();
                        double dWKTIM = 0;

                        if (sPrevCd.Equals(sWORKER) && sPrevWK.Equals(sMGUBUN))
                        {

                            if (!string.IsNullOrEmpty(sSTRDTN))
                            {
                                if (DateTime.TryParse(sTDATE + " " + sSTRDT, out DateTime stDate))
                                {
                                    if (DateTime.TryParse(sTDATEN + " " + sSTRDTN, out DateTime stDateN))
                                    {
                                        if (stDate > stDateN)
                                        {
                                            sSTRDT = sSTRDTN;
                                        }
                                    }
                                }
                            }

                            if (DateTime.TryParse(sTDATE + " " + sSTRDT, out DateTime sDate))
                            {
                                if (DateTime.TryParse(sTDATE + " " + sENDDT, out DateTime eDate))
                                {
                                    if (sDate > eDate)
                                    {
                                        sDate = sDate.AddDays(-1);
                                    }

                                    TimeSpan dateDiff = eDate - sDate;
                                    double diffTotalSeconds = dateDiff.TotalSeconds;

                                    dWKTIM = diffTotalSeconds;
                                }
                            }

                            newRow["MGUBUN"] = sMGUBUN;
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["ENDDT"] = sENDDT;
                            double newWkTim = 0;
                            double.TryParse(newRow["WKTIM"]?.ToString(), out newWkTim);
                            //newRow["WKTIM"] = newWkTim + dWKTIM;
                            newRow["WKTIM"] = dWKTIM;
                            sSTRDTN = sSTRDT;

                        }
                        else //이전 
                        {
                            if (newRow != null)
                            {
                                newDt.Rows.Add(newRow);
                            }

                            sPrevCd = sWORKER;
                            sPrevWK = sMGUBUN;
                            sPreTdate = sTDATE;
                            newRow = newDt.NewRow();

                            newRow["MGUBUN"] = sMGUBUN;
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["STRDT"] = sSTRDT;
                            newRow["ENDDT"] = sENDDT;
                            //newRow["WKTIM"] = dWKTIM;
                            newRow["WKTIM"] = row["DIFF"];

                            //if (!string.IsNullOrEmpty(sSTRDTN))
                            //{
                            //    sSTRDTN = sSTRDT;
                            //}
                            sSTRDTN = sSTRDT;
                            sTDATEN = sTDATE;
                        }
                        rowCnt++;

                        if (rowCnt == dt.Rows.Count)//마지막 행 추가
                        {
                            if (newRow != null)
                                newDt.Rows.Add(newRow);
                        }
                    }

                    if (newDt != null && newDt.Rows.Count > 0)
                    {
                        //작업자 품목별 근무시간 총합
                        string sJSON = ComnEtcFunc.DataTableToJsonWithNewtonSoft(newDt);

                        if (string.IsNullOrEmpty(sJSON))
                            return null;

                        dicParams.Clear();
                        dicParams.Add("CMD", "LIST7TS");
                        dicParams.Add("JSON1", sJSON);

                        DataTable sumDt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                        return sumDt;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 다음 근무자까지의 근무시간데이터 작업자,품목별
        /// </summary>
        /// <param name="sFdate">시작일</param>
        /// <param name="sTdate">종료일</param>
        /// <returns>DataTable</returns>
        private DataTable GetWorkAndGbTimeDataTable(string sFdate, string sTdate)
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                DataTable newDt = new DataTable();

                newDt.Columns.Add("WORKER");//작업자코드
                newDt.Columns.Add("EMP_NM");//작업자명
                newDt.Columns.Add("MGUBUN");//품목명
                newDt.Columns.Add("TDATE");//작업일
                newDt.Columns.Add("STRDT");//시작시간
                newDt.Columns.Add("ENDDT");//종료시간
                newDt.Columns.Add("WKTIM");//근무시간

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sPrevCd = string.Empty;
                    string sPrevWK = string.Empty;
                    string sPreTdate = string.Empty;
                    int rowCnt = 0;

                    DataRow newRow = null;
                    foreach (DataRow row in dt.Rows)
                    {
                        string sWORKER = row["WORKER"]?.ToString();
                        string sEMP_NM = row["EMP_NM"]?.ToString();
                        string sMGUBUN = row["MGUBUN"]?.ToString();
                        string sTDATE = row["TDATE"]?.ToString();
                        string sSTRDT = row["STRDT"]?.ToString();
                        string sENDDT = row["ENDDT"]?.ToString();
                        double dWKTIM = 0;

                        if (DateTime.TryParse(sTDATE+" "+sSTRDT, out DateTime stDate))
                        {
                            if (DateTime.TryParse(sTDATE + " " + sENDDT, out DateTime enDate))
                            {
                                if (stDate > enDate)
                                {
                                    stDate = stDate.AddDays(-1);
                                }
                        
                                TimeSpan dateDiff = enDate - stDate;
                                double diffTotalSeconds = dateDiff.TotalSeconds;
                        
                                dWKTIM = diffTotalSeconds;
                            }
                        }

                        if (sPrevWK.Equals(sMGUBUN) && sPreTdate.Equals(sTDATE))//이전 일자와 작업자 이름이 같을때
                        {
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["ENDDT"] = sENDDT;
                            double newWkTim = 0;
                            double.TryParse(newRow["WKTIM"]?.ToString(), out newWkTim);

                            newRow["WKTIM"] = newWkTim + dWKTIM;
                        }
                        else //이전 작업자가 다를때
                        {
                            if (newRow != null)
                            {
                                newDt.Rows.Add(newRow);
                            }

                            sPrevCd = sWORKER;
                            sPrevWK = sMGUBUN;
                            sPreTdate = sTDATE;
                            newRow = newDt.NewRow();


                            newRow["MGUBUN"] = sMGUBUN;
                            newRow["WORKER"] = sWORKER;
                            newRow["EMP_NM"] = sEMP_NM;
                            newRow["TDATE"] = sTDATE;
                            newRow["STRDT"] = sSTRDT;
                            newRow["ENDDT"] = sENDDT;
                            newRow["WKTIM"] = dWKTIM;
                        }
                        rowCnt++;

                        if (rowCnt == dt.Rows.Count)//마지막 행 추가
                        {
                            if (newRow != null)
                                newDt.Rows.Add(newRow);
                        }
                    }

                    if (newDt != null && newDt.Rows.Count > 0)
                    {
                        //작업자 품목별 근무시간 총합
                        string sJSON = ComnEtcFunc.DataTableToJsonWithNewtonSoft(newDt);

                        if (string.IsNullOrEmpty(sJSON))
                            return null;

                        dicParams.Clear();
                        dicParams.Add("CMD", "LIST7");
                        dicParams.Add("JSON1", sJSON);

                        DataTable sumDt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                        return sumDt;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
        }
        
        #endregion
        
        #region [ 주야별 생산내역 조회 ]
        private void GetRetr5Data()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                string sFdate = DateFromT.EditValue?.ToString().Substring(0, 10);
                string sTdate = DateToT.EditValue?.ToString().Substring(0, 10);

                dicParams.Clear();
                dicParams.Add("CMD", "LIST9");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null && dt.Rows.Count > 0)
                {
                    GridRetr5.Focus();
                }
                else
                {
                    DateFromT.Focus();
                }

                GridRetr5.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr5_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sTITLE1 = GridViewRetr5.GetRowCellValue(e.RowHandle1, "TITLE")?.ToString();
            string sTITLE2 = GridViewRetr5.GetRowCellValue(e.RowHandle2, "TITLE")?.ToString();

            var val1 = GridViewRetr5.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = GridViewRetr5.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("TITLE"))
            {
                e.Merge = (sTITLE1.Equals(sTITLE2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
            }
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnRetr.PerformClick();
            }
        }

        private void EQ001F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void RepoPicError_EditValueChanged(object sender, EventArgs e)
        {
            PictureEdit pictureEdit = sender as PictureEdit;

            string sVal = pictureEdit.EditValue?.ToString();

            if (sVal.Equals("Y"))
            {
                pictureEdit.Image = Properties.Resources.warning;
            }
            else
            {
                pictureEdit.Image = null;
            }

        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();

                dt.Columns.Add("BUNGI");
                dt.Columns.Add("ANUM2");
                dt.Columns.Add("AMSG");

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                for (int i = 0; i < 14; i++)
                {
                    if (i > 5)
                        continue;

                    string sNum = GridViewRetr.GetFocusedRowCellValue("ALARM" + i)?.ToString();

                    if (string.IsNullOrEmpty(sNum))
                        continue;

                    dicParams.Clear();
                    dicParams.Add("BUNGI", "ALARM" + i);
                    dicParams.Add("NUM", sNum);

                    DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_AL, dicParams);

                    if(dtResult != null && dtResult.Rows.Count > 0)
                    {
                        foreach (DataRow row in dtResult.Rows)
                        {
                            dt.ImportRow(row);
                        }
                    }
                }

                GridErrorMsg.DataSource = dt;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewErrorMsg_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            if (e.Column.FieldName == "EQNM")
            {
                var dr1 = GridViewErrorMsg.GetDataRow(e.RowHandle1);
                var dr2 = GridViewErrorMsg.GetDataRow(e.RowHandle2);

                e.Merge = dr1["EQNM"].ToString().Equals(dr2["EQNM"].ToString());
            }
        }

        private void EQ001F00_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];
                string path = ComnEtcFunc.GetLayoutPath();
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            BtnRetr.PerformClick();
            if (xtraTabControl1.SelectedTabPageIndex == 3)
            {
                DateFromT.Enabled = false;
                DateToT.Enabled = false;
            }
            else
            {
                DateFromT.Enabled = true;
                DateToT.Enabled = true;
            } 
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateFrom4.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateFrom4.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateFrom4.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateFrom4.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = GridViewRetr2.GetRowCellValue(e.RowHandle, "EMP_NM")?.ToString();

            if (!string.IsNullOrEmpty(sVal) && sVal.Equals("소계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (!string.IsNullOrEmpty(sVal) && sVal.Equals("합계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr3_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = GridViewRetr3.GetRowCellValue(e.RowHandle, "MGUBUN")?.ToString();

            if (!string.IsNullOrEmpty(sVal) && sVal.Equals("소계") )
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (!string.IsNullOrEmpty(sVal) && sVal.Equals("합계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                fileDlg.FileName = "길로틴 생산내역";
                if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage1"))
                    ComnEtcFunc.ExportExcelFile("길로틴 생산 내역", GridRetr);
                if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage2"))
                    ComnEtcFunc.ExportExcelFile("길로틴 생산 내역", GridRetr2);
                if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage3"))
                    ComnEtcFunc.ExportExcelFile("길로틴 생산 내역", GridRetr3);
                if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage4"))
                    ComnEtcFunc.ExportExcelFile("길로틴 생산 내역", GridRetr4);
                if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage5"))
                    ComnEtcFunc.ExportExcelFile("길로틴 생산 내역", GridRetr5);

                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }
    }
}