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
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using ComLib;
using DevExpress.XtraGrid.Views.BandedGrid;
using System.Collections;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.LookAndFeel;
using DevExpress.Utils;
using DevExpress.Data;
using static AccAdm.ComnGridFunc;
using DevExpress.XtraGrid.Helpers;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraCharts;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using DevExpress.XtraBars;
using DevExpress.XtraSplashScreen;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class AllReport : DevExpress.XtraEditors.XtraForm
    {
        public AllReport()
        {
            InitializeComponent();
        }
        private string PROCEDURE_ID = "DP_AllReport";
        private string PROCEDURE_EQ = "DP_EQ001F00";
        private string PROCEDURE_IM = "DP_IMAGEALLRE";
        string stMagam;
        DataGridViewRow selectRow;

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr1, GridViewRetr2, GridViewRetr3, GridViewRetr4
                                        , GridViewRetr5};
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

        private void AllReport_Load(object sender, EventArgs e)
        {
            InitControls01();//경영EXCEL
            InitControls11();//생산EXCEL
            InitControls21();//영업EXCEL
            ComnEtcFunc.SetBoundGridLookUp(RepoM3_OP, "HR_EMP_BASIS", "EMP_ID", "EMP_NM");
            ComnEtcFunc.SetBoundGridLookUp(RepoM3_PUT, "HR_EMP_BASIS", "EMP_ID", "EMP_NM");
            ComnEtcFunc.SetBoundGridLookUp(RepoM3_LINE, "HR_EMP_BASIS", "EMP_ID", "EMP_NM");
            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);
            layoutControlItem1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            layoutControlItem32.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            ComnEtcFunc.SetDateToValue(DateFrom4);


            SetLoadFormLayout();
            //BtnRetr.PerformClick();
        }

        #region 버튼이벤트
        private void AllReport_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnEXCEL01.PerformClick();
        }
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }
        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            if (xtraTabControl1.SelectedTabPageIndex == 0)
            {
                //경영보고 조회
                GetRetr1();//재고자산
                GetRetr2();//재무 -현금자산
                GetRetr3();//재무 -입금예정
                GetRetr4();//재무 -지급예정
                GetRetr5();//여신현황
            }
            else if (xtraTabControl1.SelectedTabPageIndex == 1)
            {
                //생산보고 조회
                GridControl1();//매출
                GridControl2();//매입
                GridControl3();//생산-길로틴
                GridControl4();//생산-슈레더
                GridControl5();//정비및고장내용
                IMAGERE();//전력비
                IMAGERE1();//계절별요금표
            }
            else if (xtraTabControl1.SelectedTabPageIndex == 2)
            {
                //영업보고 조회
                GridControl8();//매출
                GridControl7();//매입
                GridControl9();//감량조정
                GridControl10();//기준단가 보다 높은 업체
                GridControl11();//신규업체
            }
            SplashScreenManager.CloseForm();
        }

        #endregion

        #region 조회(최신)

        //재고자산
        private void GetRetr1()
        {
            string sFrom = DateFrom4.EditValue?.ToString().Substring(0, 10);
            string sTo = DateFrom4.EditValue?.ToString().Substring(0, 10);
            DateTime tDateTime = Convert.ToDateTime(DateFrom4.EditValue?.ToString().Substring(0, 10));
            stMagam = Convert.ToString(tDateTime.AddMonths(1))?.Substring(0,8) + "01";
                
           
            
            string stFrom = Convert.ToString(tDateTime.ToString())?.Substring(0, 8) + "01";

            // int iFrom = Convert.ToInt32(stFrom) - 1;
            // string xFrom = Convert.ToString(iFrom)?.ToString();
            //string sxFrom = string.Concat(stFrom, "-01-01");

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            
            dicParams.Clear();
            dicParams.Add("CMD", "LIST4");
            dicParams.Add("SDATE", stFrom);
            
            DataTable dtC = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            
            if (dtC.Rows.Count == 0)
            {
                GridRetr1.DataSource = null;
                XtraMessageBox.Show("'" + stFrom + "'월 마감이 되어있지 않습니다. 마감 후 조회해주세요");
                return;
            }

            dicParams.Clear();
            dicParams.Add("CMD", "LIST1");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            GridRetr1.DataSource = dt;
            
        }

        //재무 -현금자산
        private void GetRetr2()
        {
            //string sFrom = DateFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sTo = DateFrom4.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sFrom = sTo.ToString()?.Substring(0, 6) + "01";

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST5");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            GridRetr2.DataSource = dt;
        }
        //재무 -입금예정
        private void GetRetr3()
        {
            //string sFrom = DateFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sTo = DateFrom4.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sFrom = sTo.ToString()?.Substring(0, 6) + "01";

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST6");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            GridRetr3.DataSource = dt;
        }
        private void GetRetr4()
        {
            //string sFrom = DateFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sTo = DateFrom4.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sFrom = sTo.ToString()?.Substring(0, 6) + "01";

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST7");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            GridRetr4.DataSource = dt;
        }
        private void GetRetr5()
        {
            
            string sTo = DateFrom4.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sFrom = sTo.ToString()?.Substring(0, 6) + "01";
            
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST8");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            GridRetr5.DataSource = dt;
        }

        private void GridControl1()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "CLIST1");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl1.DataSource = dt;
        }
        private void GridControl2()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "CLIST2");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl2.DataSource = dt;
        }
        private void GridControl3()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                string sFdate = DateFrom.EditValue?.ToString().Substring(0, 10);
                string sTdate = DateTo.EditValue?.ToString().Substring(0, 10);

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

                    dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_EQ, dicParams);
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

                }

                gridControl3.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void GridControl4()
        {
            string sFrom = DateFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sTo = DateTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "CLIST4");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl6.DataSource = dt;
        }
        private void GridControl5()
        {
            string sFrom = DateFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sTo = DateTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "CLIST5");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl4.DataSource = dt;
        }
        private void IMAGERE()
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 7);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST1");
            dicParams.Add("yymm", sDate);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_IM, dicParams);

            if (dt.Rows.Count == 0)
            {
                PicInsa.Image = null;
                TEXT1.EditValue = null;
            }
            else
            {
                if (string.IsNullOrEmpty(dt.Rows[0]["IMIMAGE"]?.ToString()))
                {
                    PicInsa.Image = null;
                    TEXT1.EditValue = null;
                }
                else
                {
                    byte[] ImgData = (byte[])dt.Rows[0]["IMIMAGE"];
                    MemoryStream ms = new MemoryStream(ImgData);
                    Image returnImage = Image.FromStream(ms);
                    //PicInsa.Image = ResizeImage(returnImage, 300, 200);
                    PicInsa.Image = returnImage;
                    TEXT1.EditValue = dt.Rows[0]["RK"]?.ToString();
                }

            }

        }
        private void IMAGERE1()
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 4);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST2");
            dicParams.Add("yymm", sDate);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_IM, dicParams);

            if (dt.Rows.Count == 0)
            {
                PicInsa1.Image = null;
                TEXT2.EditValue = null;
            }
            else
            {
                if (string.IsNullOrEmpty(dt.Rows[0]["IMIMAGE"]?.ToString()))
                {
                    PicInsa1.Image = null;
                    TEXT2.EditValue = null;
                }
                else
                {
                    byte[] ImgData = (byte[])dt.Rows[0]["IMIMAGE"];
                    MemoryStream ms = new MemoryStream(ImgData);
                    Image returnImage = Image.FromStream(ms);
                    //PicInsa.Image = ResizeImage(returnImage, 300, 200);
                    PicInsa1.Image = returnImage;
                    TEXT2.EditValue = dt.Rows[0]["RK"]?.ToString();
                }
            }

        }
        private void GridControl8()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            string sDFrom = DateFrom.EditValue?.ToString().Replace("-", "").Substring(0, 6);
            string sDTo = DateTo.EditValue?.ToString().Replace("-", "").Substring(0, 6);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "YULIST1");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);
            dicParams.Add("FDDATE", sDFrom);
            dicParams.Add("TDDATE", sDTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl8.DataSource = dt;
        }
        private void GridControl7()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "YULIST2");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl7.DataSource = dt;
        }
        private void GridControl9()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "YULIST3");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl9.DataSource = dt;
        }
        private void GridControl10()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "YULIST4");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl10.DataSource = dt;
        }
        private void GridControl11()
        {
            string sFrom = DateFrom.EditValue?.ToString();
            string sTo = DateTo.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "YULIST5");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            gridControl11.DataSource = dt;
        }




        private void GridViewRetr3_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sEmpNm1 = gridView3.GetRowCellValue(e.RowHandle1, "MGUBUN")?.ToString();
            string sEmpNm2 = gridView3.GetRowCellValue(e.RowHandle2, "MGUBUN")?.ToString();

            var val1 = gridView3.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = gridView3.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("MGUBUN") || e.Column.FieldName.Equals("DIFF2") || e.Column.FieldName.Equals("WKTIM") || e.Column.FieldName.Equals("GADON"))
            {
                e.Merge = (sEmpNm1.Equals(sEmpNm2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
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

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_EQ, dicParams);

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

                        DataTable sumDt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_EQ, dicParams);

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

        private void BusinessReport_TextChanged(object sender, EventArgs e)
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
            //BtnRetr.PerformClick();
            if (xtraTabControl1.SelectedTabPageIndex == 0)
            {
                
                layoutControlItem1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem26.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                layoutControlItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem32.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;


            }
            else if (xtraTabControl1.SelectedTabPageIndex == 1)
            {
                
                layoutControlItem1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem26.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                layoutControlItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem32.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }
            else if (xtraTabControl1.SelectedTabPageIndex == 2)
            {
                
                layoutControlItem1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem26.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                layoutControlItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem32.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }

        }
        

        private void SaveIm()
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 7);
            string sRK = DateFrom.EditValue?.ToString().Substring(0, 7);

            StringBuilder strSql = new StringBuilder();

            if (PicInsa.Image == null)
            {
                if (XtraMessageBox.Show(string.Format("이미지가 없습니다. 그래도 진행하시겠습니까?"), "확인", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
            }

            byte[] byteImg = null;
            if (PicInsa.Image != null)
            {
                Image img = PicInsa.Image;
                MemoryStream ms = new MemoryStream();
                img.Save(ms, img.RawFormat);
                byteImg = ms.ToArray();
            }

            Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

            dicParams.Clear();
            dicParams.Add("CMD", "SAVEALLRE1");
            dicParams.Add("IMIMAGE", byteImg);
            dicParams.Add("MAINCD", '1');
            dicParams.Add("yymm", sDate);
            dicParams.Add("RK", sRK);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_IM, dicParams);
            string sResult = dt.Rows[0]["RESULT"]?.ToString();


            if (sResult.Equals("-1"))
            {
                XtraMessageBox.SmartTextWrap = true;
                XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString(), "등록오류");
            }
            else
            {
                XtraMessageBox.Show("이미지가 등록되었습니다.");
                BtnRetr.PerformClick();
            }
        }
        private void SaveIm1()
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 4);
            string sRK = DateFrom.EditValue?.ToString().Substring(0, 4);

            StringBuilder strSql = new StringBuilder();

            if (PicInsa1.Image == null)
            {
                if (XtraMessageBox.Show(string.Format("이미지가 없습니다. 그래도 진행하시겠습니까?"), "확인", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
            }

            byte[] byteImg = null;
            if (PicInsa1.Image != null)
            {
                Image img = PicInsa1.Image;
                MemoryStream ms = new MemoryStream();
                img.Save(ms, img.RawFormat);
                byteImg = ms.ToArray();
            }

            Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

            dicParams.Clear();
            dicParams.Add("CMD", "SAVEALLRE2");
            dicParams.Add("IMIMAGE", byteImg);
            dicParams.Add("MAINCD", '2');
            dicParams.Add("yymm", sDate);
            dicParams.Add("RK", sRK);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_IM, dicParams);
            string sResult = dt.Rows[0]["RESULT"]?.ToString();


            if (sResult.Equals("-1"))
            {
                XtraMessageBox.SmartTextWrap = true;
                XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString(), "등록오류");
            }
            else
            {
                XtraMessageBox.Show("이미지가 등록되었습니다.");
                BtnRetr.PerformClick();
            }
        }
        private void layoutControlGroup19_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            if (e.Button.Properties.Tag.Equals("T_ADD"))
            {
                string sFile = null;
                OpenFileDialog dialog = new OpenFileDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    sFile = dialog.FileName;
                    PicInsa.Image = Image.FromFile(sFile);
                }
                else
                {
                    return;
                }
            }
            else if (e.Button.Properties.Tag.Equals("T_DEL"))
            {
                PicInsa.EditValue = null;
            }
            else if (e.Button.Properties.Tag.Equals("T_SAVE"))
            {
                SaveIm();
            }
        }

        private void layoutControlGroup20_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            if (e.Button.Properties.Tag.Equals("T_ADD"))
            {
                string sFile = null;
                OpenFileDialog dialog = new OpenFileDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    sFile = dialog.FileName;
                    PicInsa1.Image = Image.FromFile(sFile);
                }
                else
                {
                    return;
                }
            }
            else if (e.Button.Properties.Tag.Equals("T_DEL"))
            {
                PicInsa1.EditValue = null;
            }
            else if (e.Button.Properties.Tag.Equals("T_SAVE"))
            {
                SaveIm1();
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

        


        private void PicInsa_Click(object sender, EventArgs e)
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 7);

            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("yymm", sDate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_IM, dicParams);

                AllReportIM01 fm = new AllReportIM01();
                fm.Owner = this;
                fm.ImgData = (byte[])dt.Rows[0]["IMIMAGE"];
                fm.Show();
            }
            catch
            {
                //XtraMessageBox.Show("이미지가 없습니다. ");
            }
        }

        private void PicInsa1_Click(object sender, EventArgs e)
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 4);

            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("yymm", sDate);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_IM, dicParams);
                AllReportIM01 fm = new AllReportIM01();
                fm.Owner = this;
                fm.ImgData = (byte[])dt.Rows[0]["IMIMAGE"];
                fm.Show();
            }
            catch
            {
                //XtraMessageBox.Show("이미지가 없습니다. ");
            }
        }


        #region[RowStyle]
        private void GridViewRetr3_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = gridView3.GetRowCellValue(e.RowHandle, "MGUBUN")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }
        private void gridView7_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = gridView7.GetRowCellValue(e.RowHandle, "GU")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void gridView1_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = gridView1.GetRowCellValue(e.RowHandle, "DaeGubun")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void gridView2_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = gridView2.GetRowCellValue(e.RowHandle, "DaeGubun")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = GridViewRetr2.GetRowCellValue(e.RowHandle, "GU")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }

            }
        }

        private void GridViewRetr3_RowStyle_1(object sender, RowStyleEventArgs e)
        {
            string sVal = GridViewRetr3.GetRowCellValue(e.RowHandle, "GU")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void GridViewRetr4_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = GridViewRetr4.GetRowCellValue(e.RowHandle, "GU")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void GridViewRetr5_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sVal = GridViewRetr5.GetRowCellValue(e.RowHandle, "GU")?.ToString();

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
                //ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
                GridView gridView = sender as GridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void bandedGridView2_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string sVal = bandedGridView2.GetRowCellValue(e.RowHandle, "CHRG_NM")?.ToString();

            if (!string.IsNullOrEmpty(sVal) && sVal.Equals("소계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
                e.Appearance.ForeColor = Color.Black;
            }
            else if (!string.IsNullOrEmpty(sVal) && sVal.Equals("합계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                BandedGridView gridView = sender as BandedGridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                    e.Appearance.ForeColor = Color.Black;
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void GridViewRetr1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string sVal = GridViewRetr1.GetRowCellValue(e.RowHandle, "GU")?.ToString();

            if (!string.IsNullOrEmpty(sVal) && sVal.Equals("소계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
                e.Appearance.ForeColor = Color.Black;
            }
            else if (!string.IsNullOrEmpty(sVal) && sVal.Equals("합계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                BandedGridView gridView = sender as BandedGridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                    e.Appearance.ForeColor = Color.Black;
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        private void bandedGridView1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            BandedGridView gridView = sender as BandedGridView;

            if (e.RowHandle == gridView.FocusedRowHandle)
            {
                // 클릭된 행의 배경색 변경
                e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                e.Appearance.BackColor = SystemColors.Window;
            }
        }

        private void gridView4_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView gridView = sender as GridView;

            if (e.RowHandle == gridView.FocusedRowHandle)
            {
                // 클릭된 행의 배경색 변경
                e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                e.Appearance.BackColor = SystemColors.Window;
            }
        }

        private void gridView8_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView gridView = sender as GridView;

            if (e.RowHandle == gridView.FocusedRowHandle)
            {
                // 클릭된 행의 배경색 변경
                e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                e.Appearance.BackColor = SystemColors.Window;
            }
        }

        private void gridView9_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView gridView = sender as GridView;

            if (e.RowHandle == gridView.FocusedRowHandle)
            {
                // 클릭된 행의 배경색 변경
                e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                e.Appearance.BackColor = SystemColors.Window;
            }
        }

        private void gridView10_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView gridView = sender as GridView;

            if (e.RowHandle == gridView.FocusedRowHandle)
            {
                // 클릭된 행의 배경색 변경
                e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                e.Appearance.BackColor = SystemColors.Window;
            }
        }
        #endregion

        #region[엑셀]
        private void BtnEXCEL_Click_1(object sender, EventArgs e)
        {
             string tag = BtnEXCEL01.Tag?.ToString();
             
             if (tag == "재고자산")
             {
                 EX01();
             }
             else if (tag == "재무-현금자산")
             {
                 EX02();
             }
             else if (tag == "재무-입금예정")
             {
                 EX03();
             }
             else if (tag == "재무-지급예정")
             {
                 EX04();
             }
             else if (tag == "여신현황")
             {
                 EX05();
             }
        }

        private void BtnEXCEL02_Click(object sender, EventArgs e)
        {
            //생산보고
            string tag = BtnEXCEL02.Tag?.ToString();

            if (tag == "매출내역")
            {
                EX11();
            }
            else if (tag == "매입내역")
            {
                EX12();
            }
            else if (tag == "생산-길로틴")
            {
                EX13();
            }
            else if (tag == "생산-슈레더")
            {
                EX14();
            }
            else if (tag == "정비및고장내용")
            {
                EX15();
            }
        }
        private void BtnEXCEL03_Click(object sender, EventArgs e)
        {
            string tag = BtnEXCEL03.Tag?.ToString();

            if (tag == "매출")
            {
                EX21();
            }
            else if (tag == "매입")
            {
                EX22();
            }
            else if (tag == "감량조정")
            {
                EX23();
            }
            else if (tag == "기준단가보다높은업체")
            {
                EX24();
            }
            else if (tag == "신규업체")
            {
                EX25();
            }
        }
        BarManager barManager01;
        PopupMenu popupMenu01;
        BarButtonItem BtnSummary01;
        BarButtonItem BtnSummary02;
        BarButtonItem BtnSummary03;
        BarButtonItem BtnSummary04;
        BarButtonItem BtnSummary05;

        BarManager barManager11;
        PopupMenu popupMenu11;
        BarButtonItem BtnSummary11;
        BarButtonItem BtnSummary12;
        BarButtonItem BtnSummary13;
        BarButtonItem BtnSummary14;
        BarButtonItem BtnSummary15;

        BarManager barManager21;
        PopupMenu popupMenu21;
        BarButtonItem BtnSummary21;
        BarButtonItem BtnSummary22;
        BarButtonItem BtnSummary23;
        BarButtonItem BtnSummary24;
        BarButtonItem BtnSummary25;

        private void InitControls01()
        {
            barManager01 = new BarManager();
            barManager01.Form = this;

            popupMenu01 = new PopupMenu(barManager01);
            BtnSummary01 = new BarButtonItem(barManager01, "재고자산");
            BtnSummary02 = new BarButtonItem(barManager01, "재무-현금자산");
            BtnSummary03 = new BarButtonItem(barManager01, "재무-입금예정");
            BtnSummary04 = new BarButtonItem(barManager01, "재무-지급예정");
            BtnSummary05 = new BarButtonItem(barManager01, "여신현황");
            popupMenu01.AddItem(BtnSummary01);
            popupMenu01.AddItem(BtnSummary02);
            popupMenu01.AddItem(BtnSummary03);
            popupMenu01.AddItem(BtnSummary04);
            popupMenu01.AddItem(BtnSummary05);

            BtnEXCEL01.DropDownControl = popupMenu01;

            BtnSummary01.Tag = "재고자산";
            BtnSummary01.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary01_ItemClick);

            BtnSummary02.Tag = "재무-현금자산";
            BtnSummary02.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary02_ItemClick);

            BtnSummary03.Tag = "재무-입금예정";
            BtnSummary03.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary03_ItemClick);

            BtnSummary04.Tag = "재무-지급예정";
            BtnSummary04.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary04_ItemClick);

            BtnSummary05.Tag = "여신현황";
            BtnSummary05.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary05_ItemClick);
        }
        private void BtnSummary01_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton01(e.Item);
            EX01();
        }
        private void BtnSummary02_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton01(e.Item);
            EX02();
        }
        private void BtnSummary03_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton01(e.Item);
            EX03();
        }
        private void BtnSummary04_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton01(e.Item);
            EX04();
        }
        private void BtnSummary05_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton01(e.Item);
            EX05();
        }
        private void UpdateDropDownButton01(BarItem submenuItem)
        {
            BtnEXCEL01.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            BtnEXCEL01.ImageOptions.SvgImageSize = new Size(16, 16);
            BtnEXCEL01.Tag = submenuItem.Tag;
        }
        private void EX01()
        {
            string date = DateFrom4.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("재고자산("+ date + ")", GridRetr1);
        }
        private void EX02()
        {
            string date = DateFrom4.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("재무-현금자산(" + date + ")", GridRetr2);
        }
        private void EX03()
        {
            string date = DateFrom4.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("재무-입금예정(" + date + ")", GridRetr3);
        }
        private void EX04()
        {
            string date = DateFrom4.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("재무-지급예정(" + date + ")", GridRetr4);
        }
        private void EX05()
        {
            string date = DateFrom4.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("여신현황(" + date + ")", GridRetr5);
        }


        private void InitControls11()
        {
            barManager11 = new BarManager();
            barManager11.Form = this;

            popupMenu11 = new PopupMenu(barManager11);
            BtnSummary11 = new BarButtonItem(barManager11, "매출내역");
            BtnSummary12 = new BarButtonItem(barManager11, "매입내역");
            BtnSummary13 = new BarButtonItem(barManager11, "생산-길로틴");
            BtnSummary14 = new BarButtonItem(barManager11, "생산-슈레더");
            BtnSummary15 = new BarButtonItem(barManager11, "정비및고장내용");
            popupMenu11.AddItem(BtnSummary11);
            popupMenu11.AddItem(BtnSummary12);
            popupMenu11.AddItem(BtnSummary13);
            popupMenu11.AddItem(BtnSummary14);
            popupMenu11.AddItem(BtnSummary15);

            BtnEXCEL02.DropDownControl = popupMenu11;

            BtnSummary11.Tag = "매출내역";
            BtnSummary11.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary11_ItemClick);

            BtnSummary12.Tag = "매입내역";
            BtnSummary12.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary12_ItemClick);

            BtnSummary13.Tag = "생산-길로틴";
            BtnSummary13.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary13_ItemClick);

            BtnSummary14.Tag = "생산-슈레더";
            BtnSummary14.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary14_ItemClick);

            BtnSummary15.Tag = "정비및고장내용";
            BtnSummary15.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary15_ItemClick);
        }
        private void BtnSummary11_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton11(e.Item);
            EX11();
        }
        private void BtnSummary12_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton11(e.Item);
            EX12();
        }
        private void BtnSummary13_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton11(e.Item);
            EX13();
        }
        private void BtnSummary14_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton11(e.Item);
            EX14();
        }
        private void BtnSummary15_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton11(e.Item);
            EX15();
        }
        private void UpdateDropDownButton11(BarItem submenuItem)
        {
            BtnEXCEL02.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            BtnEXCEL02.ImageOptions.SvgImageSize = new Size(16, 16);
            BtnEXCEL02.Tag = submenuItem.Tag;
        }
        private void EX11()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("매출내역(" + fdate + " ~ "+ tdate + ")", gridControl1);
        }
        private void EX12()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("매입내역(" + fdate + " ~ " + tdate + ")", gridControl2);
        }
        private void EX13()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("생산-길로틴(" + fdate + " ~ " + tdate + ")", gridControl3);
        }
        private void EX14()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("생산-슈레더(" + fdate + " ~ " + tdate + ")", gridControl6);
        }
        private void EX15()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("정비및고장내용(" + fdate + " ~ " + tdate + ")", gridControl4);
        }

        private void InitControls21()
        {
            barManager21 = new BarManager();
            barManager21.Form = this;

            popupMenu21 = new PopupMenu(barManager21);
            BtnSummary21 = new BarButtonItem(barManager21, "매출");
            BtnSummary22 = new BarButtonItem(barManager21, "매입");
            BtnSummary23 = new BarButtonItem(barManager21, "감량조정");
            BtnSummary24 = new BarButtonItem(barManager21, "기준단가보다높은업체");
            BtnSummary25 = new BarButtonItem(barManager21, "신규업체");
            popupMenu21.AddItem(BtnSummary21);
            popupMenu21.AddItem(BtnSummary22);
            popupMenu21.AddItem(BtnSummary23);
            popupMenu21.AddItem(BtnSummary24);
            popupMenu21.AddItem(BtnSummary25);

            BtnEXCEL03.DropDownControl = popupMenu21;

            BtnSummary21.Tag = "매출";
            BtnSummary21.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary21_ItemClick);

            BtnSummary22.Tag = "매입";
            BtnSummary22.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary22_ItemClick);

            BtnSummary23.Tag = "감량조정";
            BtnSummary23.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary23_ItemClick);

            BtnSummary24.Tag = "기준단가보다높은업체";
            BtnSummary24.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary24_ItemClick);

            BtnSummary25.Tag = "신규업체";
            BtnSummary25.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSummary25_ItemClick);
        }
        private void BtnSummary21_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton21(e.Item);
            EX21();
        }
        private void BtnSummary22_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton21(e.Item);
            EX22();
        }
        private void BtnSummary23_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton21(e.Item);
            EX23();
        }
        private void BtnSummary24_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton21(e.Item);
            EX24();
        }
        private void BtnSummary25_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton21(e.Item);
            EX25();
        }
        private void UpdateDropDownButton21(BarItem submenuItem)
        {
            BtnEXCEL03.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            BtnEXCEL03.ImageOptions.SvgImageSize = new Size(16, 16);
            BtnEXCEL03.Tag = submenuItem.Tag;
        }
        private void EX21()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("매출(" + fdate + " ~ " + tdate + ")", gridControl8);
        }
        private void EX22()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("매입(" + fdate + " ~ " + tdate + ")", gridControl7);
        }
        private void EX23()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("감량조정(" + fdate + " ~ " + tdate + ")", gridControl9);
        }
        private void EX24()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("기준단가보다높은업체(" + fdate + " ~ " + tdate + ")", gridControl10);
        }
        private void EX25()
        {
            string fdate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string tdate = DateTo.EditValue?.ToString().Substring(0, 10);
            ComnEtcFunc.ExportExcelFile("신규업체(" + fdate + " ~ " + tdate + ")", gridControl11);
        }
        #endregion

        private void DateFrom4_EditValueChanged(object sender, EventArgs e)
        {
           // DateFrom.EditValue = DateFrom4.EditValue?.ToString().Substring(0,7) + "-01";
        }
    }
}