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

namespace AccAdm
{
    public partial class BusinessReport : DevExpress.XtraEditors.XtraForm
    {
        public BusinessReport()
        {
            InitializeComponent();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr1, GridViewRetr2, GridViewRetr3, GridViewRetr4
                                        , GridViewRetr5, BGridViewRetr1, BGridViewRetr2, BGridViewYMRetr
                                        , GridViewRetr8, GridViewRetr9, GridViewRetr10 };
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

        private void BusinessReport_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);
            SetLoadFormLayout();
        }

        #region 버튼이벤트
        private void BusinessReport_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            //주간내역
            GridRetr1.DataSource = GetRetr1();
            GridRetr2.DataSource = GetRetr2();
            GridRetr3.DataSource = GetRetr3();
            GridRetr4.DataSource = GetRetr4();
            GridRetr5.DataSource = GetRetr5();

            //재고
            GetJaeGoData();

            //영업부-실적
            //주간
            SetRetr6();
            //월간
            SetRetr7();
            //매출실적
            GetRetr8();

            //매출매입현황
            DataTable dt9 = GetRetr9();
            GridRetr9.DataSource = dt9;
            SetChart1(dt9);

            GridRetr10.DataSource = GetRetr10();
            SetChart2();
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            FileInfo_1 fileInfo = new FileInfo_1("8");

            Cursor = Cursors.WaitCursor;
            string[] sPath = fileInfo.CheckFileInfo();
            Cursor = Cursors.Default;

            if (sPath != null)
            {
                SetWeeklyReportForm(sPath[0], sPath[1]);
            }
        }

        #region 엑셀 양식
        Excel.Application ExcelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;
        private void SetWeeklyReportForm(string StandardPath, string SavePath)
        {
            try
            {
                if (!File.Exists(StandardPath))
                {
                    XtraMessageBox.Show("엑셀파일 양식이 존재하지 않습니다.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
                string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

                int jucha = ComnEtcFunc.GetWeekOfMonth(DateTime.Parse(sTo), null);

                string newFileName = "업무보고서 - "+ sFrom.Substring(5,2)+"월 " + jucha + "주차"+" ("+ sFrom.Substring(5,5) + " ~ "+ sTo.Substring(5, 5) + ").xlsx";

                ExcelApp = new Excel.Application();
                wb = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                DataTable dt = new DataTable();

                #region 주간내역
                ws = wb.Worksheets.get_Item(1);

                //일자부분 세팅
                ws.Range["B1"].Value = sFrom;
                ws.Range["D1"].Value = sTo;

                //스크랩사업부
                dt = (DataTable)GridRetr1.DataSource;

                if (dt != null)
                {
                    int iStartRow = 4;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sDANJUNG = dt.Rows[i]["DANJUNG"]?.ToString();
                        string sDANGA = dt.Rows[i]["DANGA"]?.ToString();
                        string sKONGKEP = dt.Rows[i]["KONGKEP"]?.ToString();
                        string sBIGO = dt.Rows[i]["BIGO"]?.ToString();

                        int iApplyRowIdx = iStartRow + i;

                        ws.Range["E" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["F" + iApplyRowIdx].Value = sDANGA;
                        ws.Range["G" + iApplyRowIdx].Value = sKONGKEP;
                        ws.Range["H" + iApplyRowIdx].Value = sBIGO;
                    }
                }

                //슈레더 사업부
                dt = (DataTable)GridRetr2.DataSource;

                if (dt != null)
                {
                    int iStartRow = 4;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sDANJUNG = dt.Rows[i]["DANJUNG"]?.ToString();
                        string sDANGA = dt.Rows[i]["DANGA"]?.ToString();
                        string sKONGKEP = dt.Rows[i]["KONGKEP"]?.ToString();

                        int iApplyRowIdx = iStartRow + i;

                        ws.Range["N" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["O" + iApplyRowIdx].Value = sDANGA;
                        ws.Range["P" + iApplyRowIdx].Value = sKONGKEP;
                    }
                }

                //슈레더별
                dt = (DataTable)GridRetr3.DataSource;

                if (dt != null)
                {
                    int iStartRow = 11;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sGUBUN    = dt.Rows[i]["GUBUN"]?.ToString();
                        string sRECOVERY = dt.Rows[i]["RECOVERY"]?.ToString();
                        string sDANJUNG  = dt.Rows[i]["DANJUNG"]?.ToString();
                        string sDANGA    = dt.Rows[i]["DANGA"]?.ToString();
                        string sSANBI    = dt.Rows[i]["SANBI"]?.ToString();
                        string sSDANJUNG = dt.Rows[i]["SDANJUNG"]?.ToString();
                        string sMDANGA   = dt.Rows[i]["MDANGA"]?.ToString();
                        string sMECHIC = dt.Rows[i]["MECHIC"]?.ToString();

                        int iApplyRowIdx = iStartRow + i;

                        if (sGUBUN.Equals("합계"))
                        {
                            ws.Range["K" + 25].Value = sGUBUN;
                            ws.Range["L" + 25].Value = sRECOVERY;
                            ws.Range["N" + 25].Value = sDANJUNG;
                            ws.Range["O" + 25].Value = sDANGA;
                            ws.Range["P" + 25].Value = sSANBI;
                            ws.Range["Q" + 25].Value = sSDANJUNG;
                            ws.Range["R" + 25].Value = sMDANGA;
                            ws.Range["S" + 25].Value = sMECHIC;
                        }
                        else
                        {
                            ws.Range["K" + iApplyRowIdx].Value = sGUBUN;
                            ws.Range["L" + iApplyRowIdx].Value = sRECOVERY;
                            ws.Range["N" + iApplyRowIdx].Value = sDANJUNG;
                            ws.Range["O" + iApplyRowIdx].Value = sDANGA;
                            ws.Range["P" + iApplyRowIdx].Value = sSANBI;
                            ws.Range["Q" + iApplyRowIdx].Value = sSDANJUNG;
                            ws.Range["R" + iApplyRowIdx].Value = sMDANGA;
                            ws.Range["S" + iApplyRowIdx].Value = sMECHIC;
                        }
                    }
                }

                //직납매입출
                dt = (DataTable)GridRetr4.DataSource;

                if (dt != null)
                {
                    int iStartRow = 26;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sDANJUNG = dt.Rows[i]["DANJUNG"]?.ToString();
                        string sDANGA = dt.Rows[i]["DANGA"]?.ToString();
                        string sKONGKEP = dt.Rows[i]["KONGKEP"]?.ToString();

                        int iApplyRowIdx = iStartRow + i;

                        ws.Range["D" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["E" + iApplyRowIdx].Value = sDANGA;
                        ws.Range["F" + iApplyRowIdx].Value = sKONGKEP;
                    }
                }

                //총 매출
                dt = (DataTable)GridRetr5.DataSource;

                if (dt != null)
                {
                    int iStartRow = 28;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sWEIGHT = dt.Rows[i]["WEIGHT"]?.ToString();
                        string sPRICE = dt.Rows[i]["PRICE"]?.ToString();
                        string sMONMCH = dt.Rows[i]["MONMCH"]?.ToString();

                        int iApplyRowIdx = iStartRow + i;

                        ws.Range["L" + iApplyRowIdx].Value = sWEIGHT;
                        ws.Range["N" + iApplyRowIdx].Value = sPRICE;
                        ws.Range["P" + iApplyRowIdx].Value = sMONMCH;
                    }
                }
                #endregion

                #region 재고현황
                ws = wb.Worksheets.get_Item(2);

                dt = (DataTable)BGridYMRetr.DataSource;

                if (dt != null)
                {
                    //일자부분 세팅
                    ws.Range["B3"].Value = DateTime.Parse(sTo).ToString("yyyy년 MM월");
                    int iStartRow = 6;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sWEIGHT1 = dt.Rows[i]["WEIGHT1"]?.ToString();
                        string sDANGA1 = dt.Rows[i]["DANGA1"]?.ToString();
                        string sKONGKEP1 = dt.Rows[i]["KONGKEP1"]?.ToString();
                        string sWEIGHT2 = dt.Rows[i]["WEIGHT2"]?.ToString();
                        string sDANGA2 = dt.Rows[i]["DANGA2"]?.ToString();
                        string sKONGKEP2 = dt.Rows[i]["KONGKEP2"]?.ToString();
                        string sWEIGHT3 = dt.Rows[i]["WEIGHT3"]?.ToString();
                        string sDANGA3 = dt.Rows[i]["DANGA3"]?.ToString();
                        string sKONGKEP3 = dt.Rows[i]["KONGKEP3"]?.ToString();
                        string sWEIGHT4 = dt.Rows[i]["WEIGHT4"]?.ToString();
                        string sDANGA4 = dt.Rows[i]["DANGA4"]?.ToString();
                        string sKONGKEP4 = dt.Rows[i]["KONGKEP4"]?.ToString();
                        //string sWEIGHT5  = dt.Rows[i]["WEIGHT5"]?.ToString();
                        //string sDANGA5   = dt.Rows[i]["DANGA5"]?.ToString();
                        //string sKONGKEP5 = dt.Rows[i]["KONGKEP5"]?.ToString();
                        string sWEIGHT6 = dt.Rows[i]["WEIGHT6"]?.ToString();
                        string sDANGA6 = dt.Rows[i]["DANGA6"]?.ToString();
                        string sKONGKEP6 = dt.Rows[i]["KONGKEP6"]?.ToString();

                        int iApplyRowIdx = iStartRow + (i + 1);

                        ws.Range["C" + iApplyRowIdx].Value = sWEIGHT1;
                        ws.Range["D" + iApplyRowIdx].Value = sDANGA1;
                        ws.Range["E" + iApplyRowIdx].Value = sKONGKEP1;
                        ws.Range["F" + iApplyRowIdx].Value = sWEIGHT2;
                        ws.Range["G" + iApplyRowIdx].Value = sDANGA2;
                        ws.Range["H" + iApplyRowIdx].Value = sKONGKEP2;
                        ws.Range["I" + iApplyRowIdx].Value = sWEIGHT3;
                        ws.Range["J" + iApplyRowIdx].Value = sDANGA3;
                        ws.Range["K" + iApplyRowIdx].Value = sKONGKEP3;
                        ws.Range["L" + iApplyRowIdx].Value = sWEIGHT4;
                        ws.Range["M" + iApplyRowIdx].Value = sDANGA4;
                        ws.Range["N" + iApplyRowIdx].Value = sKONGKEP4;
                        //ws.Range["O" + iApplyRowIdx].Value = sWEIGHT5;
                        //ws.Range["P" + iApplyRowIdx].Value = sDANGA5;
                        //ws.Range["Q" + iApplyRowIdx].Value = sKONGKEP5;
                        ws.Range["O" + iApplyRowIdx].Value = sWEIGHT6;
                        ws.Range["P" + iApplyRowIdx].Value = sDANGA6;
                        ws.Range["Q" + iApplyRowIdx].Value = sKONGKEP6;
                    }
                }
                #endregion

                #region 영업부-실적
                ws = wb.Worksheets.get_Item(3);

                string[] strList = { "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
                                    , "AA","AB","AC","AD","AE","AF","AG", "AH","AI","AJ","AK","AL","AM","AN", "AO","AP","AQ","AR","AS","AT","AU"};
                int nameidex = 0;
                int addCnt = 0;

                dt = (DataTable)BGridRetr1.DataSource;

                if (dt != null)
                {
                    int iStartRow = 4;

                    if (dt.Columns.Count - 3 > 6)
                    {
                        for (int j = 0; j < dt.Columns.Count - 9; j++)
                        {
                            Excel.Range range = ws.Columns["I"];
                            range.Insert();
                            addCnt += 1;
                        }
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        int iApplyRowIdx = iStartRow + i;
                        nameidex = 0;

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            string sName = dt.Columns[j].ColumnName;

                            if (!sName.Equals("TITLE1") && !sName.Equals("TITLE2"))
                            {
                                if (!sName.Equals("GITA"))
                                {
                                    ws.Range[strList[nameidex] + 3].Value = sName;
                                }

                                string sVal = row[j]?.ToString();

                                ws.Range[strList[nameidex++] + iApplyRowIdx].Value = sVal;
                            }
                        }
                    }
                }

                dt = (DataTable)BGridRetr2.DataSource;

                if (dt != null)
                {
                    int iStartRow = 4;
                    int nameidx2 = nameidex + addCnt + 3;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        int iApplyRowIdx = iStartRow + i;
                        nameidex = nameidx2;

                        if (dt.Columns.Count - 3 > 6)
                        {
                            addCnt = 0;
                            for (int j=0;j< dt.Columns.Count - 9; j++)
                            {
                                Excel.Range range = ws.Columns[strList[15+ addCnt]];
                                range.Insert();

                                addCnt += 1;
                            }
                        }

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            string sName = dt.Columns[j].ColumnName;

                            if (!sName.Equals("TITLE1") && !sName.Equals("TITLE2"))
                            {
                                if (!sName.Equals("GITA"))
                                {
                                    ws.Range[strList[nameidex] + 3].Value = sName;
                                }

                                string sVal = row[j]?.ToString();

                                ws.Range[strList[nameidex++] + iApplyRowIdx].Value = sVal;
                            }
                        }
                    }
                }

                dt = (DataTable)GridRetr8.DataSource;

                if(dt != null)
                {
                    int iStartRow = 3;
                    int nameidx2 = nameidex + addCnt + 3;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        int iApplyRowIdx = iStartRow + i;
                        nameidex = nameidx2;

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            string sName = dt.Columns[j].ColumnName;

                            if (!sName.Equals("TITLE"))
                            {
                                string sVal = row[j]?.ToString();

                                ws.Range[strList[nameidex++] + iApplyRowIdx].Value = sVal;
                            }
                        }
                    }
                }
                #endregion

                #region 매출매입현황
                ws = wb.Worksheets.get_Item(4);

                //매출
                dt = (DataTable)GridRetr9.DataSource;
                int addRowCnt = 0;

                if(dt != null)
                {
                    int iStartRow = 4;

                    if (dt.Rows.Count > 20)
                    {
                        for(int i=0;i< dt.Rows.Count - 20; i++)
                        {
                            Excel.Range range = ws.Rows[23];
                            range.Insert();
                            addRowCnt += 1;
                        }
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sGUBUN1  = dt.Rows[i]["GUBUN1"]?.ToString();
                        string sCNT     = dt.Rows[i]["CNT"]?.ToString();
                        string sSANCNT  = dt.Rows[i]["SANCNT"]?.ToString();
                        string sDANJUNG = dt.Rows[i]["DANJUNG"]?.ToString();
                        string sAVGDJ   = dt.Rows[i]["AVGDJ"]?.ToString();
                        string sHALIN   = dt.Rows[i]["HALIN"]?.ToString();
                        string sNAPJAN = dt.Rows[i]["NAPJAN"]?.ToString();

                        int iApplyRowIdx = iStartRow + i;

                        ws.Range["A" + iApplyRowIdx].Value = sGUBUN1;
                        ws.Range["B" + iApplyRowIdx].Value = sCNT;
                        ws.Range["C" + iApplyRowIdx].Value = sSANCNT;
                        ws.Range["D" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["E" + iApplyRowIdx].Value = sAVGDJ;
                        ws.Range["F" + iApplyRowIdx].Value = sHALIN;
                        ws.Range["G" + iApplyRowIdx].Value = sNAPJAN;
                    }
                }

                //매입
                dt = (DataTable)GridRetr10.DataSource;
                if (dt != null)
                {
                    int scrapCnt = 0;
                    int surederCnt = 0;

                    for(int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sNum = dt.Rows[i]["NUM"]?.ToString();

                        if (sNum.Equals("1"))
                            scrapCnt++;
                        else if (sNum.Equals("3"))
                            surederCnt++;
                    }

                    int addRowCntScrap = 0;
                    if (scrapCnt > 33)
                    {
                        for (int i = 0; i < scrapCnt - 33; i++)
                        {
                            Excel.Range range = ws.Rows[61];
                            range.Insert();
                            addRowCntScrap ++;
                        }
                    }

                    int addRowCntSureder = 0;
                    if (surederCnt > 12)
                    {
                        for (int i = 0; i < surederCnt - 12; i++)
                        {
                            Excel.Range range = ws.Rows[74 + addRowCntScrap];
                            range.Insert();
                            addRowCntSureder++;
                        }
                    }

                    int iStartRowScrep = 29 + addRowCnt;
                    int iStartRowSureder = 63 + addRowCntScrap;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sNum = dt.Rows[i]["NUM"]?.ToString();
                        string sGUBUN1 = dt.Rows[i]["GUBUN1"]?.ToString();
                        string sDANJUNG = dt.Rows[i]["DANJUNG"]?.ToString();
                        string sTDANGA = dt.Rows[i]["TDANGA"]?.ToString();
                        string sAVGDAN = dt.Rows[i]["AVGDAN"]?.ToString();
                        string sHALIN = dt.Rows[i]["HALIN"]?.ToString();
                        string sBANCNT = dt.Rows[i]["BANCNT"]?.ToString();

                        int iApplyRowIdx = 0;

                        if (sNum.Equals("2") || sNum.Equals("4") || sNum.Equals("5"))//스크랩 합계
                        {
                            continue;
                        }
                        else if (sNum.Equals("1"))
                        {
                            iApplyRowIdx = iStartRowScrep++;
                        }
                        else if (sNum.Equals("3"))
                        {
                            iApplyRowIdx = iStartRowSureder++;
                        }

                        ws.Range["A" + iApplyRowIdx].Value = sGUBUN1;
                        ws.Range["B" + iApplyRowIdx].Value = sDANJUNG;
                        ws.Range["D" + iApplyRowIdx].Value = sTDANGA;
                        ws.Range["E" + iApplyRowIdx].Value = sAVGDAN;
                        ws.Range["F" + iApplyRowIdx].Value = sHALIN;
                        ws.Range["G" + iApplyRowIdx].Value = sBANCNT;
                    }
                }
                #endregion

                if (File.Exists(SavePath))
                    File.Delete(SavePath);

                Cursor = Cursors.Default;

                SavePath = SavePath.Replace("업무보고.xlsx", newFileName);
                wb.SaveAs(SavePath);
                wb.Close(false, Type.Missing, Type.Missing);
                wb = null;
                ExcelApp.Quit();

                Process.Start(SavePath);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                XtraMessageBox.Show(ex.Message);

                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
            finally
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
        }

        private Image GetChartImage(ChartControl chart, ImageFormat format)
        {
            // Create an image. 
            Image image = null;

            // Create an image of the chart. 
            using (MemoryStream s = new MemoryStream())
            {
                chart.ExportToImage(s, format);
                image = Image.FromStream(s);
            }

            // Return the image. 
            return image;
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }
        #endregion

        #region 그리드 이벤트
        private void GridViewRetr1_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            if (e.Column.FieldName == "TITLE")
            {
                var dr1 = GridViewRetr1.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr1.GetDataRow(e.RowHandle2);

                e.Merge = dr1["TITLE"].ToString().Equals(dr2["TITLE"].ToString());
            }

            if (e.Column.FieldName == "GUBUN")
            {
                var dr1 = GridViewRetr1.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr1.GetDataRow(e.RowHandle2);

                e.Merge = dr1["GUBUN"].ToString().Equals(dr2["GUBUN"].ToString());
            }
        }
        private void GridViewRetr2_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            if (e.Column.FieldName == "TITLE")
            {
                var dr1 = GridViewRetr2.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr2.GetDataRow(e.RowHandle2);

                e.Merge = dr1["TITLE"].ToString().Equals(dr2["TITLE"].ToString());
            }

            if (e.Column.FieldName == "GUBUN")
            {
                var dr1 = GridViewRetr2.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr2.GetDataRow(e.RowHandle2);

                e.Merge = dr1["GUBUN"].ToString().Equals(dr2["GUBUN"].ToString());
            }
        }

        private void GridViewRetr4_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            if (e.Column.FieldName == "TITLE")
            {
                var dr1 = GridViewRetr4.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr4.GetDataRow(e.RowHandle2);

                e.Merge = dr1["TITLE"].ToString().Equals(dr2["TITLE"].ToString());
            }

            if (e.Column.FieldName == "COMP")
            {
                var dr1 = GridViewRetr4.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr4.GetDataRow(e.RowHandle2);

                e.Merge = dr1["COMP"].ToString().Equals(dr2["COMP"].ToString());
            }

            if (e.Column.FieldName == "DANJUNG")
            {
                var dr1 = GridViewRetr4.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr4.GetDataRow(e.RowHandle2);

                e.Merge = dr1["DANJUNG"].ToString().Equals(dr2["DANJUNG"].ToString());
            }
        }

        private void BGridViewRetr1_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            if (e.Column.FieldName == "TITLE1")
            {
                var dr1 = BGridViewRetr1.GetDataRow(e.RowHandle1);
                var dr2 = BGridViewRetr1.GetDataRow(e.RowHandle2);

                e.Merge = dr1["TITLE1"].ToString().Equals(dr2["TITLE1"].ToString());
            }

            if (e.Column.FieldName == "TITLE2")
            {
                var dr1 = BGridViewRetr1.GetDataRow(e.RowHandle1);
                var dr2 = BGridViewRetr1.GetDataRow(e.RowHandle2);

                e.Merge = dr1["TITLE2"].ToString().Equals(dr2["TITLE2"].ToString());
            }

            if (e.Column.FieldName == "GITA")
            {
                var dr1 = BGridViewRetr1.GetDataRow(e.RowHandle1);
                var dr2 = BGridViewRetr1.GetDataRow(e.RowHandle2);

                e.Merge = dr1["GITA"].ToString().Equals(dr2["GITA"].ToString());
            }
        }

        private void GridViewRetr1_RowStyle_1(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string title = view.GetRowCellValue(e.RowHandle, "TITLE")?.ToString();

            if (title == null)
                return;

            if (title.Equals("차액"))
            {
                e.Appearance.BackColor = Color.FromArgb(253,233,217);
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }

            string gubun = view.GetRowCellValue(e.RowHandle, "GUBUN")?.ToString();
            if (gubun != null && gubun.Equals("합계"))
            {
                e.Appearance.BackColor = Color.LightGray;
            }
            
        }

        private void GridViewRetr1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string col = view.GetRowCellValue(e.RowHandle, "COL")?.ToString();
            if (col != null)
            {
                if (col.Equals("Y"))
                {
                    if (e.Column.FieldName.Equals("BIGO"))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
            }

            string val = e.CellValue?.ToString();

            double.TryParse(val, out double dVal);

            if(!double.IsNaN(dVal) && dVal < 0)
            {
                e.Appearance.ForeColor = Color.Red;
            }
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string GUBUN = view.GetRowCellValue(e.RowHandle, "GUBUN")?.ToString();

            if (GUBUN == null)
                return;

            if (GUBUN.Equals("ASR"))
            {
                e.Appearance.BackColor = Color.Yellow;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr3_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string GUBUN = view.GetRowCellValue(e.RowHandle, "GUBUN")?.ToString();

            if (GUBUN == null)
                return;

            if (GUBUN.Equals("합계"))
            {
                e.Appearance.BackColor = Color.LightGray;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr4_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string TITLE = view.GetRowCellValue(e.RowHandle, "TITLE")?.ToString();

            if (TITLE == null)
                return;

            if (TITLE.Equals("차액"))
            {
                e.Appearance.BackColor = Color.LightGray;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr5_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string TITLE = view.GetRowCellValue(e.RowHandle, "TITLE")?.ToString();

            if (TITLE == null)
                return;

            if (TITLE.Equals("총 매출 현황"))
            {
                e.Appearance.BackColor = Color.LightGray;
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void BGridViewYMRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            string title = BGridViewYMRetr.GetRowCellValue(e.RowHandle, GridColTitle)?.ToString();

            if (DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName != "DevExpress Dark Style")
            {
                if (!string.IsNullOrEmpty(title) && (title.Equals("월초재고") || title.Equals("월말재고")))
                {
                    e.Appearance.BackColor = Color.LightGreen;
                }
                else if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }
            }
            else if (DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName == "DevExpress Dark Style")
            {
                if (!string.IsNullOrEmpty(title) && (title.Equals("월초재고") || title.Equals("월말재고")))
                {
                    e.Appearance.BackColor = Color.LightGreen;
                }
                else if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.ControlDark;
                }
            }
        }

        private void BGridViewRetr1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            e.Appearance.BackColor = Color.FromArgb(255, 255, 204);

            string val = e.CellValue?.ToString();

            double.TryParse(val, out double dVal);

            if (!double.IsNaN(dVal) && dVal < 0)
            {
                e.Appearance.ForeColor = Color.Red;
            }

            string gubun = view.GetRowCellValue(e.RowHandle, "TITLE2")?.ToString();

            if(!string.IsNullOrEmpty(gubun) && gubun.Contains("목표"))
            {
                if(!e.Column.FieldName.Equals("TITLE1") && !e.Column.FieldName.Equals("TITLE2") && !e.Column.FieldName.Equals("GITA"))
                {
                    e.Appearance.BackColor = Color.FromArgb(146, 208, 80);
                }
            }

            string title = view.GetRowCellValue(e.RowHandle, "TITLE1")?.ToString();

            if(!string.IsNullOrEmpty(title) && title.Equals("전체"))
            {
                if (e.Column.FieldName.Equals("TITLE1") || e.Column.FieldName.Equals("TITLE2"))
                {
                    e.Appearance.BackColor = Color.FromArgb(252, 213, 180);
                }
            }
            else
            {
                if (e.Column.FieldName.Equals("TITLE1") || e.Column.FieldName.Equals("TITLE2"))
                {
                    e.Appearance.BackColor = Color.FromArgb(183, 222, 232);
                }
            }
        }

        private void BGridViewRetr1_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string TITLE2 = view.GetRowCellValue(e.RowHandle, "TITLE2")?.ToString();

            if(TITLE2 != null && TITLE2.Equals("달성율"))
            {
                CellDrawHelper.DoDefaultDrawCell(view, e);
                CellDrawHelper.DrawCellBorderBottom(e);
                e.Handled = true;
            }

            if(e.Column.FieldName.Equals("TITLE1"))
            {
                CellDrawHelper.DoDefaultDrawCell(view, e);
                CellDrawHelper.DrawCellBorderRigthBottom(e);
                e.Handled = true;
            }

            if (e.Column.FieldName.Equals("GITA"))
            {
                string title1 = view.GetRowCellValue(e.RowHandle, "TITLE1")?.ToString();

                if (!string.IsNullOrEmpty(title1) && !title1.Equals("전체"))
                {
                    CellDrawHelper.DoDefaultDrawCell(view, e);
                    CellDrawHelper.DrawCellBorderBottom(e);
                    e.Handled = true;
                }
            }
        }

        public static class CellDrawHelper
        {
            public static void DrawCellBorderBottom(DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
            {
                Brush brush = Brushes.Black;
                //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.Right - 1, e.Bounds.Y - 1, 2, e.Bounds.Height + 2));  // 오른쪽
                //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Y - 1, e.Bounds.Width + 2, 2));
                e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Bottom - 1, e.Bounds.Width + 2, 2));
                //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Y - 1, 2, e.Bounds.Height + 2));

            }

            public static void DrawCellBorderRigthBottom(DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
            {
                Brush brush = Brushes.Black;
                e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.Right - 1, e.Bounds.Y - 1, 2, e.Bounds.Height + 2));  // 오른쪽
                //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Y - 1, e.Bounds.Width + 2, 2));
                e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Bottom - 1, e.Bounds.Width + 2, 2));
                //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Y - 1, 2, e.Bounds.Height + 2));
            }

            public static void DoDefaultDrawCell(GridView view, RowCellCustomDrawEventArgs e)
            {
                e.Appearance.FillRectangle(e.Cache, e.Bounds);
                ((IViewController)view.GridControl).EditorHelper.DrawCellEdit(new GridViewDrawArgs(e.Cache, (view.GetViewInfo() as GridViewInfo), e.Bounds), (e.Cell as GridCellInfo).Editor, (e.Cell as GridCellInfo).ViewInfo, e.Appearance, (e.Cell as GridCellInfo).CellValueRect.Location);
            }
        }

        private void GridViewRetr10_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null)
                return;

            string GUBUN1 = view.GetRowCellValue(e.RowHandle, "GUBUN1")?.ToString();

            if (GUBUN1 == null)
                return;

            if (GUBUN1.Contains("합계"))
            {
                e.Appearance.BackColor = Color.FromArgb(235,241,222);
            }
            else
            {
                ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
            }
        }

        private void GridViewRetr1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion

        #region 데이터 select function

        #region 주간내역
        //스크랩 사업부
        private DataTable GetRetr1()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0,10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                                                                       ");
            strSql.AppendLine(" SELECT 'YK,대한(고철A)' AS COMP                                                                                    ");
            strSql.AppendLine("      , (SUM(A.DANJUNG) + SUM(A.HALIN))/ 1000 AS DANJUNG--톤수                                                      ");
		    strSql.AppendLine("      , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / 1000) / ((SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000) AS DANGA --평균단가  ");
		    strSql.AppendLine("      , (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / 1000 AS KONGKEP--금액                                                  ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                                           ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                                     ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                                         ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                              ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                            ");
            strSql.AppendLine("    AND(A.J_ID1 = '6531121044' OR A.J_ID1 = '6541233044')                                                           ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철A'                                                                                        ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브'                                                                                      ");
            strSql.AppendLine("  UNION ALL                                                                                                         ");
            strSql.AppendLine(" SELECT 'YK,대한(고철B)' AS COMP                                                                                    ");
            strSql.AppendLine("      , (SUM(A.DANJUNG) + SUM(A.HALIN))/ 1000 AS DANJUNG--톤수                                                      ");
		    strSql.AppendLine("      , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / 1000) / ((SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000) AS DANGA --평균단가  ");
		    strSql.AppendLine("      , (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / 1000 AS KONGKEP--금액                                                  ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                                           ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                                     ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                                         ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                              ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                            ");
            strSql.AppendLine("    AND(A.J_ID1 = '6531121044' OR A.J_ID1 = '6541233044')                                                           ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철B'                                                                                        ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브'                                                                                      ");
            strSql.AppendLine("  UNION ALL                                                                                                         ");
            strSql.AppendLine(" SELECT '타업체'                                                                                                    ");
		    strSql.AppendLine("      , (SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000 AS DANJUNG                                                           ");
            strSql.AppendLine("      , ((SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / 1000)/ ((SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000) AS DANGA              ");
            strSql.AppendLine("      , (SUM(A.KONGKEP) - SUM(A.CKONGKEP))/ 1000 AS KONGKEP                                                         ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                                           ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                                     ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                                         ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                              ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                            ");
            strSql.AppendLine("    AND A.J_ID1 NOT IN('6531121044', '6541233044')                                                                  ");
            strSql.AppendLine("    AND B.DAEGUBUN <> '슈레더'                                                                                      ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                                                         ");
            strSql.AppendLine(" SELECT '매입' AS GUBUN                                                                   ");
            strSql.AppendLine("      , '고철A' AS COMP                                                                   ");
            strSql.AppendLine("      , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                                   ");
            strSql.AppendLine("      , SUM(A.HALIN) AS HALIN                                                             ");
            strSql.AppendLine("      , ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / 1000)/ (SUM(A.DANJUNG) / 1000) AS DANGA    ");
            strSql.AppendLine("      , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ 1000 AS IKONGKEP                             ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                 ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                           ");
            strSql.AppendLine("    AND A.KERATYPE = '매입'                                                               ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                    ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철A'                                                              ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                  ");
            strSql.AppendLine("    AND B.DAEGUBUN <> '슈레더'                                                            ");
            strSql.AppendLine("  UNION ALL                                                                               ");
            strSql.AppendLine(" SELECT '매입' AS GUBUN                                                                   ");
            strSql.AppendLine("      , '고철B' AS COMP                                                                   ");
            strSql.AppendLine("      , SUM(A.DANJUNG)/ 1000 AS DANJUNG                                                   ");
            strSql.AppendLine("       , SUM(A.HALIN) AS HALIN                                                            ");
            strSql.AppendLine("       , ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / 1000)/ (SUM(A.DANJUNG) / 1000) AS DANGA   ");
            strSql.AppendLine("            , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ 1000 AS IKONGKEP                       ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                 ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                           ");
            strSql.AppendLine("    AND A.KERATYPE = '매입'                                                               ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                    ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철B'                                                              ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                  ");
            strSql.AppendLine("    AND B.DAEGUBUN <> '슈레더'                                                            ");
            strSql.AppendLine(" ), TEMP3 AS(                                                                                                         ");
            strSql.AppendLine("     SELECT '스크랩 사업부' AS TITLE                                                                                  ");
            strSql.AppendLine("          , Z1.*                                                                                                      ");
            strSql.AppendLine("       FROM(SELECT '매출' AS GUBUN                                                                                    ");
            strSql.AppendLine("                   , COMP                                                                                             ");
            strSql.AppendLine("                   , DANJUNG                                                                                          ");
            strSql.AppendLine("                   , DANGA                                                                                            ");
            strSql.AppendLine("                   , KONGKEP                                                                                          ");
            strSql.AppendLine("                FROM TEMP1                                                                                            ");
            strSql.AppendLine("               UNION ALL                                                                                              ");
            strSql.AppendLine("              SELECT '합계'                                                                                           ");
            strSql.AppendLine("                   , ''                                                                                               ");
            strSql.AppendLine("                   , SUM(DANJUNG)                                                                                     ");
            strSql.AppendLine("                   , SUM(KONGKEP) / SUM(DANJUNG)                                                                      ");
            strSql.AppendLine("                   , SUM(KONGKEP)                                                                                     ");
            strSql.AppendLine("                FROM TEMP1                                                                                            ");
            strSql.AppendLine("               UNION ALL                                                                                              ");
            strSql.AppendLine("              SELECT GUBUN                                                                                            ");
            strSql.AppendLine("                   , COMP                                                                                             ");
            strSql.AppendLine("                   , DANJUNG                                                                                          ");
            strSql.AppendLine("                   , DANGA                                                                                            ");
            strSql.AppendLine("                   , IKONGKEP                                                                                         ");
            strSql.AppendLine("                FROM TEMP2                                                                                            ");
            strSql.AppendLine("               UNION ALL                                                                                              ");
            strSql.AppendLine("              SELECT '합계'                                                                                           ");
            strSql.AppendLine("                   , ''                                                                                               ");
            strSql.AppendLine("                   , SUM(DANJUNG)                                                                                     ");
            strSql.AppendLine("                   , SUM(IKONGKEP) / SUM(DANJUNG)                                                                     ");
            strSql.AppendLine("                   , SUM(IKONGKEP)                                                                                    ");
            strSql.AppendLine("                FROM TEMP2) Z1                                                                                        ");
            strSql.AppendLine(" ), TEMP4 AS(                                                                                       ");
            strSql.AppendLine("     --재고이동(스크랩)                                                                             ");
            strSql.AppendLine("     SELECT ROUND((SUM(CASE WHEN A1.J_SERIAL IN(5059071, 5059072) THEN A1.WEIGHT ELSE 0 END)                   ");   
            strSql.AppendLine("              -SUM(CASE WHEN A1.J_SERIAL IN(5059073, 5059074) THEN A1.WEIGHT ELSE 0 END)) / 1000,0) AS WEIGHT  ");
            strSql.AppendLine("          , ROUND((SUM(CASE WHEN A1.J_SERIAL IN(5059071, 5059072) THEN A1.PRICE ELSE 0 END)                    ");
            strSql.AppendLine("             - SUM(CASE WHEN A1.J_SERIAL IN(5059073, 5059074) THEN A1.PRICE ELSE 0 END)) / 1000, 0) AS PRICE   ");
            strSql.AppendLine("       FROM J_MAGAM A1                                                                              ");
            strSql.AppendLine("       LEFT JOIN jajae A2                                                                           ");
            strSql.AppendLine("         ON A1.J_SERIAL = A2.J_Serial                                                               ");
            strSql.AppendLine("      WHERE A2.DaeGubun = '재고이동'                                                                ");
            strSql.AppendLine("        AND A1.BASDT BETWEEN '"+sFrom+"' AND '"+sTo+"'                                          ");
            strSql.AppendLine(" ), TEMP5 AS(                                                                                       ");
            strSql.AppendLine("     SELECT*                                                                                        ");
            strSql.AppendLine("       FROM TEMP4                                                                                   ");
            strSql.AppendLine("      UNION ALL                                                                                     ");
            strSql.AppendLine("     SELECT SUM(DANJUNG)                                                                            ");
            strSql.AppendLine("         , SUM(IKONGKEP)                                                                            ");
            strSql.AppendLine("      FROM TEMP2                                                                                    ");
            strSql.AppendLine(" )                                                                                                  ");
            strSql.AppendLine("                                                                                                    ");
            strSql.AppendLine("                                                                                                    ");
            strSql.AppendLine(" SELECT 'N' AS COL, *                                                                                           ");
            strSql.AppendLine("      , NULL AS BIGO                                                                                ");
            strSql.AppendLine("   FROM TEMP3                                                                                       ");
            strSql.AppendLine("  UNION ALL                                               ");
            strSql.AppendLine(" SELECT 'N' AS COL, '차액'                                            ");
            strSql.AppendLine("      , ' '                                               ");
            strSql.AppendLine("      , ''                                                ");
            strSql.AppendLine("      , NULL                                              ");
            strSql.AppendLine("      , (SELECT SUM(KONGKEP) / SUM(DANJUNG) FROM TEMP1)   ");
            strSql.AppendLine("         -(SELECT SUM(IKONGKEP) / SUM(DANJUNG) FROM TEMP2)");
            strSql.AppendLine("      , NULL                                              ");
            strSql.AppendLine("      , NULL                                              ");
            strSql.AppendLine("  UNION ALL                                                                                         ");
            strSql.AppendLine(" SELECT 'Y' AS COL, '재고이동 포함'                                                                             ");
            strSql.AppendLine("      , '  '                                                                                         ");
            strSql.AppendLine("      , ''                                                                                          ");
            strSql.AppendLine("      , SUM(WEIGHT)                                                                                 ");
            strSql.AppendLine("      , SUM(PRICE) / SUM(WEIGHT)                                                                    ");
            strSql.AppendLine("      , SUM(PRICE)                                                                                  ");
            strSql.AppendLine("      , (SELECT FORMAT(CONVERT(NUMERIC, WEIGHT), N'#,0') FROM TEMP4)                                                ");                                                            
            strSql.AppendLine("   FROM TEMP5                                                                                       ");
            strSql.AppendLine("  UNION ALL                                                                                         ");
            strSql.AppendLine(" SELECT 'Y' AS COL, '차액'                                                                                      ");
            strSql.AppendLine("      , ''                                                                                          ");
            strSql.AppendLine("      , ''                                                                                          ");
            strSql.AppendLine("      , NULL                                                                                        ");
            strSql.AppendLine("      , (SELECT SUM(KONGKEP) / SUM(DANJUNG) FROM TEMP1)                                             ");
            strSql.AppendLine("         -(SUM(PRICE) / SUM(WEIGHT))                                                                ");
            strSql.AppendLine("      , NULL                                                                                        ");
            strSql.AppendLine("      , (SELECT FORMAT(CONVERT(NUMERIC,PRICE), N'#,0') FROM TEMP4)                                                ");
            strSql.AppendLine("   FROM TEMP5                                                                                       ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //슈레더 사업부
        private DataTable GetRetr2()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" SELECT '슈레더 사업부' AS TITLE                                                                        ");
            strSql.AppendLine("      , '매출' AS GUBUN                                                                                 ");
            strSql.AppendLine("      , 'YK스틸' AS COMP                                                                                ");
            strSql.AppendLine("      , (SUM(A.DANJUNG) + SUM(A.HALIN))/ 1000 AS DANJUNG                                                ");
            strSql.AppendLine("      , ((SUM(A.KONGKEP) + SUM(A.CKONGKEP)) / 1000)/ ((SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000) AS DANGA  ");
            strSql.AppendLine("      , (SUM(A.KONGKEP) + SUM(A.CKONGKEP))/ 1000 AS KONGKEP                                             ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                               ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                         ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                             ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                  ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                ");
            strSql.AppendLine("    AND A.J_ID1 = '6531121044'                                                                          ");
            strSql.AppendLine("    AND B.DAEGUBUN = '슈레더'                                                                           ");
            strSql.AppendLine("  UNION ALL                                                                                             ");
            strSql.AppendLine(" SELECT '슈레더 사업부' AS TITLE                                                                        ");
            strSql.AppendLine("      , '매출' AS GUBUN                                                                                 ");
            strSql.AppendLine("      , '기타' AS COMP                                                                                  ");
            strSql.AppendLine("      , (SUM(A.DANJUNG) + SUM(A.HALIN))/ 1000 AS DANJUNG                                                ");
            strSql.AppendLine("      , ((SUM(A.KONGKEP) + SUM(A.CKONGKEP)) / 1000)/ ((SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000) AS DANGA  ");
            strSql.AppendLine("      , (SUM(A.KONGKEP) + SUM(A.CKONGKEP))/ 1000 AS KONGKEP                                             ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                               ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                         ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                             ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                  ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                ");
            strSql.AppendLine("    AND A.J_ID1 <> '6531121044'                                                                         ");
            strSql.AppendLine("    AND B.DAEGUBUN = '슈레더'                                                                           ");
            strSql.AppendLine("  UNION ALL                                                                                             ");
            strSql.AppendLine(" SELECT '슈레더 사업부'                                                                                 ");
            strSql.AppendLine("      , 'ASR'                                                                                           ");
            strSql.AppendLine("      , ''                                                                                              ");
            strSql.AppendLine("      , SUM(WEIGHT) / 1000                                                                              ");
            strSql.AppendLine("      , SUM(PRICE) / SUM(WEIGHT) / 1000                                                                 ");
            strSql.AppendLine("      , SUM(PRICE) / 1000                                                                               ");
            strSql.AppendLine("   FROM J_MAGAM                                                                                         ");
            strSql.AppendLine("  WHERE J_SERIAL = 2025163                                                                              ");
            strSql.AppendLine("    AND BASDT BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                     ");
            strSql.AppendLine("  UNION ALL                                                                                             ");
            strSql.AppendLine(" SELECT '슈레더 사업부'                                                                                 ");
            strSql.AppendLine("      , '매입'                                                                                          ");
            strSql.AppendLine("      , '매입'                                                                                          ");
            strSql.AppendLine("      , SUM(A.DANJUNG) / 1000 AS DANJUNG                                                                ");
            strSql.AppendLine("      , ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / 1000)/ (SUM(A.DANJUNG) / 1000) AS DANGA                  ");
            strSql.AppendLine("      , (SUM(A.IKONGKEP) + SUM(A.CKONGKEP))/ 1000 AS IKONGKEP                                           ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                               ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                         ");
            strSql.AppendLine("    AND A.KERATYPE = '매입'                                                                             ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                  ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                ");
            strSql.AppendLine("    AND B.DAEGUBUN = '슈레더'                                                                           ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //슈레더 매출이익
        private DataTable GetRetr3()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" DECLARE @MDANGA NUMERIC;                                                                               ");
            strSql.AppendLine(" SELECT @MDANGA = ((SUM(A.KONGKEP) + SUM(A.CKONGKEP)) / 1000) / ((SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000)");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                               ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                                         ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                                             ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                                  ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                                ");
            strSql.AppendLine("    AND A.J_ID1 = '6531121044'                                                                          ");
            strSql.AppendLine("    AND B.DAEGUBUN = '슈레더';                                                                          ");
            strSql.AppendLine("             WITH TEMP1 AS(                                                                             ");
            strSql.AppendLine("              SELECT A1.GUBUN                                                                           ");
            strSql.AppendLine("                   , MAX(A1.RECOVERY) AS RECOVERY                                                       ");
            strSql.AppendLine("                   , SUM(B2.DANJUNG) AS DANJUNG                                                         ");
            strSql.AppendLine("                   , (SUM(B2.IKONGKEP) + SUM(B2.CKONGKEP)) / SUM(B2.DANJUNG)AS DANGA                    ");
            strSql.AppendLine("                   , SUM(B2.IKONGKEP) + SUM(B2.CKONGKEP) AS IKONGKEP                                    ");
            strSql.AppendLine("                   , MAX(A1.RECOVERY) * SUM(B2.DANJUNG) AS SDANJUNG                                     ");
            strSql.AppendLine("                   , 36.6 AS SANBI--생산비용                                                            ");
            strSql.AppendLine("                   , @MDANGA AS MDANGA                                                                  ");
            strSql.AppendLine("                FROM MEAIPSILJUK A1                                                                     ");
            strSql.AppendLine("                LEFT JOIN JAJAE B1                                                                      ");
            strSql.AppendLine("                  ON A1.GUBUN = B1.GUBUN1                                                               ");
            strSql.AppendLine("                LEFT JOIN INLIST B2                                                                     ");
            strSql.AppendLine("                  ON B1.J_SERIAL = B2.J_SERIAL                                                          ");
            strSql.AppendLine("                 AND B2.J_DATE BETWEEN'"+sFrom+"' AND '"+sTo+"'                                     ");
            strSql.AppendLine("                 AND B2.KERATYPE = '매입'                                                               ");
            strSql.AppendLine("                 AND B2.J_LOTNO <> '4'                                                                  ");
            strSql.AppendLine("               WHERE A1.DAEGUBUN = '슈레더'                                                             ");
            strSql.AppendLine("                 AND A1.J_DATE <= (SELECT MAX(J_DATE) FROM MEAIPSILJUK WHERE GUBUN = A1.GUBUN AND J_DATE <= '"+sTo+"')");
            strSql.AppendLine("   GROUP BY A1.GUBUN                                                                                                     ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                                                            ");
            strSql.AppendLine("     SELECT A1.*                                                                                                         ");
 	        strSql.AppendLine("          , (SDANJUNG * MDANGA) - (DANJUNG * (DANGA + 36.6)) - ((DANJUNG - SDANJUNG) * 185) AS MECHIC                    ");
            strSql.AppendLine("       FROM TEMP1 A1                                                                                                     ");
            strSql.AppendLine(" )                                                                                                                       ");
            strSql.AppendLine(" SELECT GUBUN                                                                                                            ");
            strSql.AppendLine("      , RECOVERY                                                                                                         ");
            strSql.AppendLine("      , DANJUNG                                                                                                          ");
            strSql.AppendLine("      , DANGA                                                                                                            ");
            strSql.AppendLine("      , SANBI                                                                                                            ");
            strSql.AppendLine("      , SDANJUNG                                                                                                         ");
            strSql.AppendLine("      , MDANGA                                                                                                           ");
            strSql.AppendLine("      , MECHIC                                                                                                           ");
            strSql.AppendLine("   FROM TEMP2                                                                                                            ");
            strSql.AppendLine("  UNION ALL                                                                                                              ");
            strSql.AppendLine(" SELECT '합계'                                                                                                           ");
            strSql.AppendLine("      , NULL                                                                                                             ");
            strSql.AppendLine("      , SUM(DANJUNG)                                                                                                     ");
            strSql.AppendLine("      , SUM(IKONGKEP) / SUM(DANJUNG)                                                                                     ");
            strSql.AppendLine("      , NULL                                                                                                             ");
            strSql.AppendLine("      , SUM(SDANJUNG)                                                                                                    ");
            strSql.AppendLine("      , NULL                                                                                                             ");
            strSql.AppendLine("      , SUM(MECHIC)                                                                                                      ");
            strSql.AppendLine("   FROM TEMP2                                                                                                            ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //직납매입출
        private DataTable GetRetr4()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                                            ");
            strSql.AppendLine("      SELECT 1 AS NUM                                                                      ");
            strSql.AppendLine("           , '직납 매출' AS TITLE                                                          ");
            strSql.AppendLine("           , 'YK' AS COMP                                                                  ");
            strSql.AppendLine("           , ISNULL(SUM(DANJUNG) / 1000, 0) AS DANJUNG                                     ");
            strSql.AppendLine("           , ISNULL(SUM(KONGKEP) / SUM(DANJUNG), 0) AS DANGA                               ");
            strSql.AppendLine("           , ISNULL(SUM(KONGKEP) / 1000, 0) AS KONGKEP-- 매출                              ");
            strSql.AppendLine("        FROM INLIST                                                                        ");
            strSql.AppendLine("       WHERE KERATYPE = '매출'                                                             ");
            strSql.AppendLine("         AND J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                  ");
            strSql.AppendLine("         AND J_LOTNO = '4'                                                                 ");
            strSql.AppendLine("         AND J_ID1 = '6531121044'                                                          ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT 2                                                                             ");
            strSql.AppendLine("           , '직납 매입'                                                                   ");
            strSql.AppendLine("           , 'YK'                                                                          ");
            strSql.AppendLine("           , ISNULL(SUM(DANJUNG) / 1000, 0) AS DANJUNG                                     ");
            strSql.AppendLine("           , ISNULL((SUM(MIKONGKEP) + SUM(CKONGKEP)) / SUM(DANJUNG), 0) AS DANGA           ");
            strSql.AppendLine("           , ISNULL((SUM(MIKONGKEP) + SUM(CKONGKEP)) / 1000, 0) AS MIKONGKEP--매입         ");
            strSql.AppendLine("        FROM INLIST                                                                        ");
            strSql.AppendLine("       WHERE KERATYPE = '매출'                                                             ");
            strSql.AppendLine("         AND J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                  ");
            strSql.AppendLine("         AND J_LOTNO = '4'                                                                 ");
            strSql.AppendLine("         AND J_ID1 = '6531121044'                                                          ");
            strSql.AppendLine("  ), TEMP2 AS(                                                                             ");
            strSql.AppendLine("      SELECT 4 AS NUM                                                                      ");
            strSql.AppendLine("           , '직납 매출' AS TITLE                                                          ");
            strSql.AppendLine("           , '대한' AS COMP                                                                ");
            strSql.AppendLine("           , ISNULL(SUM(DANJUNG) / 1000, 0) AS DANJUNG                                     ");
            strSql.AppendLine("           , ISNULL(SUM(KONGKEP) / SUM(DANJUNG), 0) AS DANGA                               ");
            strSql.AppendLine("             , ISNULL(SUM(KONGKEP) / 1000, 0) AS KONGKEP                                   ");
            strSql.AppendLine("        FROM INLIST                                                                        ");
            strSql.AppendLine("       WHERE KERATYPE = '매출'                                                             ");
            strSql.AppendLine("         AND J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                  ");
            strSql.AppendLine("         AND J_LOTNO = '4'                                                                 ");
            strSql.AppendLine("         AND J_ID1 = '6541233044'                                                          ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT 5                                                                             ");
            strSql.AppendLine("           , '직납 매입'                                                                   ");
            strSql.AppendLine("           , '대한'                                                                        ");
            strSql.AppendLine("           , ISNULL(SUM(DANJUNG) / 1000, 0) AS DANJUNG                                     ");
            strSql.AppendLine("           , ISNULL((SUM(MIKONGKEP) + SUM(CKONGKEP)) / SUM(DANJUNG), 0) AS DANGA           ");
            strSql.AppendLine("             , ISNULL((SUM(MIKONGKEP) + SUM(CKONGKEP)) / 1000, 0) AS MIKONGKEP             ");
            strSql.AppendLine("        FROM INLIST                                                                        ");
            strSql.AppendLine("       WHERE KERATYPE = '매출'                                                             ");
            strSql.AppendLine("         AND J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                  ");
            strSql.AppendLine("         AND J_LOTNO = '4'                                                                 ");
            strSql.AppendLine("         AND J_ID1 = '6541233044'                                                          ");
            strSql.AppendLine("  ), TEMP3 AS(                                                                             ");
            strSql.AppendLine("      SELECT 7 AS NUM                                                                      ");
            strSql.AppendLine("           , '직납 매출' AS TITLE                                                          ");
            strSql.AppendLine("           , '기타' AS COMP                                                                ");
            strSql.AppendLine("           , ISNULL(SUM(DANJUNG) / 1000, 0) AS DANJUNG                                     ");
            strSql.AppendLine("            , ISNULL(SUM(KONGKEP) / SUM(DANJUNG), 0) AS DANGA                              ");
            strSql.AppendLine("              , ISNULL(SUM(KONGKEP) / 1000, 0) AS KONGKEP                                  ");
            strSql.AppendLine("        FROM INLIST                                                                        ");
            strSql.AppendLine("       WHERE KERATYPE = '매출'                                                             ");
            strSql.AppendLine("         AND J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                  ");
            strSql.AppendLine("         AND J_LOTNO = '4'                                                                 ");
            strSql.AppendLine("         AND J_ID1 NOT IN('6531121044', '6541233044')                                      ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT 8                                                                             ");
            strSql.AppendLine("           , '직납 매입'                                                                   ");
            strSql.AppendLine("           , '기타'                                                                        ");
            strSql.AppendLine("           , ISNULL(SUM(DANJUNG) / 1000, 0) AS DANJUNG                                     ");
            strSql.AppendLine("            , ISNULL((SUM(MIKONGKEP) + SUM(CKONGKEP)) / SUM(DANJUNG), 0) AS DANGA          ");
            strSql.AppendLine("              , ISNULL((SUM(MIKONGKEP) + SUM(CKONGKEP)) / 1000, 0) AS MIKONGKEP            ");
            strSql.AppendLine("        FROM INLIST                                                                        ");
            strSql.AppendLine("       WHERE KERATYPE = '매출'                                                             ");
            strSql.AppendLine("        AND J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                                   ");
            strSql.AppendLine("        AND J_LOTNO = '4'                                                                  ");
            strSql.AppendLine("        AND J_ID1 NOT IN('6531121044', '6541233044')                                       ");
            strSql.AppendLine("  ), TEMP4 AS(                                                                             ");
            strSql.AppendLine("      SELECT*                                                                              ");
            strSql.AppendLine("        FROM TEMP1                                                                         ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT 3                                                                             ");
            strSql.AppendLine("           , '차액'                                                                        ");
            strSql.AppendLine("           , ' '                                                                           ");
            strSql.AppendLine("           , NULL                                                                          ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매출' THEN DANGA END)                             ");
            strSql.AppendLine("             - SUM(CASE WHEN TITLE = '직납 매입' THEN DANGA END)                           ");
            strSql.AppendLine("           , NULL                                                                          ");
            strSql.AppendLine("        FROM TEMP1                                                                         ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("       SELECT*                                                                             ");
            strSql.AppendLine("        FROM TEMP2                                                                         ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT 6                                                                             ");
            strSql.AppendLine("           , '차액'                                                                        ");
            strSql.AppendLine("           , ' '                                                                           ");
            strSql.AppendLine("           , NULL                                                                          ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매출' THEN DANGA END)                             ");
            strSql.AppendLine("             - SUM(CASE WHEN TITLE = '직납 매입' THEN DANGA END)                           ");
            strSql.AppendLine("           , NULL                                                                          ");
            strSql.AppendLine("        FROM TEMP2                                                                         ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("                                                                                           ");
            strSql.AppendLine("                                                                                           ");
            strSql.AppendLine("       SELECT*                                                                             ");
            strSql.AppendLine("        FROM TEMP3                                                                         ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT 9                                                                             ");
            strSql.AppendLine("           , '차액'                                                                        ");
            strSql.AppendLine("           , ' '                                                                           ");
            strSql.AppendLine("           , NULL                                                                          ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매출' THEN DANGA END)                             ");
            strSql.AppendLine("             - SUM(CASE WHEN TITLE = '직납 매입' THEN DANGA END)                           ");
            strSql.AppendLine("           , NULL                                                                          ");
            strSql.AppendLine("        FROM TEMP3                                                                         ");
            strSql.AppendLine("  ), TEMP5 AS(                                                                             ");
            strSql.AppendLine("      SELECT 10 AS NUM                                                                     ");
            strSql.AppendLine("           , '직납 매출' AS TITLE                                                          ");
            strSql.AppendLine("           , '합계' AS COMP                                                                ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매출' THEN DANJUNG END) AS DANJUNG                ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매출' THEN KONGKEP END)/ SUM(CASE WHEN TITLE = '직납 매출' THEN DANJUNG END) AS DANGA");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매출' THEN KONGKEP END) AS KONGKEP                                                   ");
            strSql.AppendLine("        FROM TEMP4                                                                                                            ");
            strSql.AppendLine("       UNION ALL                                                                                                              ");
            strSql.AppendLine("      SELECT 11                                                                                                               ");
            strSql.AppendLine("           , '직납 매입'                                                                                                      ");
            strSql.AppendLine("           , '합계'                                                                                                           ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매입' THEN DANJUNG END)                                                              ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매입' THEN KONGKEP END) / SUM(CASE WHEN TITLE = '직납 매입' THEN DANJUNG END)        ");
            strSql.AppendLine("           , SUM(CASE WHEN TITLE = '직납 매입' THEN KONGKEP END)                                                              ");
            strSql.AppendLine("        FROM TEMP4                                                                                                            ");
            strSql.AppendLine("  )                                                                                                                           ");
            strSql.AppendLine("  SELECT Z1.*                                                                                                                 ");
            strSql.AppendLine("    FROM(SELECT *                                                                                                             ");
            strSql.AppendLine("            FROM TEMP4                                                                                                        ");
            strSql.AppendLine("           UNION ALL                                                                                                          ");
            strSql.AppendLine("          SELECT *                                                                                                            ");
            strSql.AppendLine("            FROM TEMP5                                                                                                        ");
            strSql.AppendLine("           UNION ALL                                                                                                          ");
            strSql.AppendLine("          SELECT 12                                                                                                           ");
            strSql.AppendLine("               , '차액'                                                                                                       ");
            strSql.AppendLine("               , ' '                                                                                                          ");
            strSql.AppendLine("               , NULL                                                                                                         ");
            strSql.AppendLine("               , SUM(CASE WHEN TITLE = '직납 매출' THEN DANGA END)                                                            ");
            strSql.AppendLine("                 - SUM(CASE WHEN TITLE = '직납 매입' THEN DANGA END)                                                          ");
            strSql.AppendLine("               , NULL                                                                                                         ");
            strSql.AppendLine("            FROM TEMP5) Z1                                                                                                    ");
            strSql.AppendLine("   ORDER BY Z1.NUM                                                                                                            ");


            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //총 매출현황
        private DataTable GetRetr5()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);
            string sYM = sTo.Substring(0, 7);

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                              ");
            strSql.AppendLine("     SELECT 1 AS NUM ");
            strSql.AppendLine("          , '월 말 재고' AS TITLE                                            ");
            strSql.AppendLine("          , NULL AS WEIGHT                                     ");
            strSql.AppendLine("          , NULL AS PRICE                                       ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 2 AS NUM ");
            strSql.AppendLine("          , 'YK 슈레더 매출'                                                 ");
            strSql.AppendLine("          , (SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000 AS DANJUNG                ");
            strSql.AppendLine("          , SUM(A.KONGKEP) / 1000 AS KONGKEP                                 ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                                ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                          ");
            strSql.AppendLine("        AND A.KERATYPE = '매출'                                              ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sYM + "-01' AND '"+sTo+"'                   ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                 ");
            strSql.AppendLine("        AND A.J_ID1 = '6531121044'                                           ");
            strSql.AppendLine("        AND B.DAEGUBUN = '슈레더'                                            ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 3 AS NUM ");
            strSql.AppendLine("          , 'YK 스크랩 매출'                                                 ");
            strSql.AppendLine("          , (SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000 AS DANJUNG                ");
            strSql.AppendLine("          , SUM(A.KONGKEP) / 1000 AS KONGKEP                                 ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                                ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                          ");
            strSql.AppendLine("        AND A.KERATYPE = '매출'                                              ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sYM + "-01' AND '"+sTo+"'                   ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                 ");
            strSql.AppendLine("        AND A.J_ID1 = '6531121044'                                           ");
            strSql.AppendLine("        AND B.DAEGUBUN <> '슈레더'                                           ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 4 AS NUM ");
            strSql.AppendLine("          , 'YK 직납 매출'                                                   ");
            strSql.AppendLine("          , SUM(DANJUNG) / 1000 AS DANJUNG                                   ");
            strSql.AppendLine("          , SUM(KONGKEP) / 1000 AS KONGKEP                                   ");
            strSql.AppendLine("       FROM INLIST                                                           ");
            strSql.AppendLine("      WHERE KERATYPE = '매출'                                                ");
            strSql.AppendLine("        AND J_DATE BETWEEN '"+ sYM + "-01' AND '"+sTo+"'                     ");
            strSql.AppendLine("        AND J_LOTNO = '4'                                                    ");
            strSql.AppendLine("        AND J_ID1 = '6531121044'                                             ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 5 AS NUM ");
            strSql.AppendLine("          ,  'YK 수입직납 매출'                                               ");
            strSql.AppendLine("          , NULL                                                             ");
            strSql.AppendLine("          , NULL                                                             ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 6 AS NUM ");
            strSql.AppendLine("          , '대한 스크랩 매출'                                               ");
            strSql.AppendLine("          , (SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000 AS DANJUNG                ");
            strSql.AppendLine("          , SUM(A.KONGKEP) / 1000 AS KONGKEP                                 ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                                ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                          ");
            strSql.AppendLine("        AND A.KERATYPE = '매출'                                              ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sYM + "-01' AND  '"+sTo+"'                  ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                 ");
            strSql.AppendLine("        AND A.J_ID1 = '6541233044'                                           ");
            strSql.AppendLine("        AND B.DAEGUBUN <> '슈레더'                                           ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 7 AS NUM ");
            strSql.AppendLine("          , '대한 직납 매출'                                                 ");
            strSql.AppendLine("          , SUM(DANJUNG) / 1000                                              ");
            strSql.AppendLine("          , SUM(KONGKEP) / 1000                                              ");
            strSql.AppendLine("       FROM INLIST                                                           ");
            strSql.AppendLine("      WHERE KERATYPE = '매출'                                                ");
            strSql.AppendLine("        AND J_DATE BETWEEN '"+ sYM + "-01' AND  '"+sTo+"'                    ");
            strSql.AppendLine("        AND J_LOTNO = '4'                                                    ");
            strSql.AppendLine("        AND J_ID1 = '6541233044'                                             ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 8 AS NUM ");
            strSql.AppendLine("          , '기타 스크랩 매출'                                               ");
            strSql.AppendLine("          , (SUM(A.DANJUNG) + SUM(A.HALIN)) / 1000 AS DANJUNG                ");
            strSql.AppendLine("          , SUM(A.KONGKEP) / 1000 AS KONGKEP                                 ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                                ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                          ");
            strSql.AppendLine("        AND A.KERATYPE = '매출'                                              ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sYM + "-01' AND  '"+sTo+"'                  ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                 ");
            strSql.AppendLine("        AND A.J_ID1 NOT IN('6531121044', '6541233044')                       ");
            strSql.AppendLine("        AND B.DAEGUBUN <> '슈레더'                                           ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 9 AS NUM ");
            strSql.AppendLine("          , '기타 직납 매출'                                                 ");
            strSql.AppendLine("          , SUM(DANJUNG) / 1000                                              ");
            strSql.AppendLine("          , SUM(KONGKEP) / 1000                                              ");
            strSql.AppendLine("       FROM INLIST                                                           ");
            strSql.AppendLine("      WHERE KERATYPE = '매출'                                                ");
            strSql.AppendLine("        AND J_DATE BETWEEN '"+ sYM + "-01' AND  '"+sTo+"'                    ");
            strSql.AppendLine("        AND J_LOTNO = '4'                                                    ");
            strSql.AppendLine("        AND J_ID1 NOT IN('6531121044', '6541233044')                         ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                ");
            strSql.AppendLine("     SELECT*                                                                 ");
            strSql.AppendLine("       FROM TEMP1                                                            ");
            strSql.AppendLine("      UNION ALL                                                              ");
            strSql.AppendLine("     SELECT 10 AS NUM ");
            strSql.AppendLine("          , '총 매출 현황'                                                   ");
	        strSql.AppendLine("          , SUM(WEIGHT)                                                      ");
	        strSql.AppendLine("          , SUM(PRICE)                                                       ");
            strSql.AppendLine("       FROM TEMP1                                                            ");
            strSql.AppendLine(" )                                                                           ");
            strSql.AppendLine("                                                                             ");
            strSql.AppendLine(" SELECT *                                                                    ");
            strSql.AppendLine("      , CASE WHEN TITLE NOT IN('월 말 재고', '총 매출 현황')                 ");
            strSql.AppendLine("        THEN WEIGHT/ (SELECT WEIGHT FROM TEMP2 WHERE TITLE = '총 매출 현황') ELSE NULL END AS MONMCH");
            strSql.AppendLine("   FROM TEMP2");
            strSql.AppendLine("  ORDER BY NUM");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        #endregion

        #region 재고현황
        private void GetJaeGoData()
        {
            string sDate = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF                                                                                                       ");
            strSql.AppendLine(" SET ARITHIGNORE ON                                                                                                          ");
            strSql.AppendLine(" SET ARITHABORT OFF;                                                                                                         ");
            strSql.AppendLine("            WITH TEMP1 AS(                                                                                                   ");
            strSql.AppendLine("                  SELECT A3.DaeGubun AS GUBUN                                                                                ");
            strSql.AppendLine("                       , SUBSTRING(A1.J_Date, 1, 7) AS BASDT                                                                 ");
            strSql.AppendLine("                       , A1.KERATYPE                                                                                         ");
            strSql.AppendLine("                       , CASE WHEN A1.KeraType = '매입' THEN SUM(A1.iWeight)                                                 ");
            strSql.AppendLine("                              WHEN A1.KeraType = '매출' THEN SUM(A1.OWeight) END AS WEIGHT                                   ");
            strSql.AppendLine("                       , CASE WHEN A1.KeraType = '매입' THEN SUM(A1.iKongKep)                                                ");
            strSql.AppendLine("                              WHEN A1.KeraType = '매출' THEN SUM(A1.KongKep) END AS KONGKEP                                  ");
            strSql.AppendLine("                       , SUM(CASE WHEN A1.KeraType = '매출' THEN A1.Halin ELSE NULL END) AS LOSSWEIGHT                       ");
            strSql.AppendLine("                       , FLOOR((CASE WHEN A1.KeraType = '매입' THEN SUM(A1.iKongKep) / SUM(A1.iWeight)                       ");
            strSql.AppendLine("                                    WHEN A1.KeraType = '매출' THEN SUM(A1.KongKep) / SUM(A1.OWeight) END) * 10) / 10 AS DANGA");
            strSql.AppendLine("         FROM INLIST A1                                                                    ");
            strSql.AppendLine("         LEFT JOIN MESURING A2                                                             ");
            strSql.AppendLine("           ON A1.J_RID = A2.JUNPYOID                                                       ");
            strSql.AppendLine("         LEFT JOIN JAJAE A3                                                                ");
            strSql.AppendLine("           ON A1.J_Serial = A3.J_Serial                                                    ");
            strSql.AppendLine("        WHERE A2.KeraType <> '직송'                                                        ");
            strSql.AppendLine("          AND A1.J_Date BETWEEN '" + sDate.Substring(0, 7) + "-01' AND '" + sDate + "'                              ");
            strSql.AppendLine("        GROUP BY A3.DaeGubun, SUBSTRING(A1.J_Date, 1, 7), A1.KeraType                      ");
            strSql.AppendLine("       UNION ALL                                                                           ");
            strSql.AppendLine("      SELECT '슈레더비철' AS GUBUN                                                         ");
            strSql.AppendLine("           , SUBSTRING(A1.J_Date, 1, 7) AS BASDT                                           ");
            strSql.AppendLine("           , A1.KERATYPE                                                                   ");
            strSql.AppendLine("           , CASE WHEN A1.KeraType = '매입' THEN SUM(A1.iWeight)                           ");
            strSql.AppendLine("                  WHEN A1.KeraType = '매출' THEN SUM(A1.OWeight) END AS WEIGHT             ");
            strSql.AppendLine("           , CASE WHEN A1.KeraType = '매입' THEN SUM(A1.iKongKep)                          ");
            strSql.AppendLine("                  WHEN A1.KeraType = '매출' THEN SUM(A1.KongKep) END AS KONGKEP            ");
            strSql.AppendLine("           , SUM(CASE WHEN A1.KeraType = '매출' THEN A1.Halin ELSE NULL END) AS LOSSWEIGHT ");
            strSql.AppendLine("           , FLOOR((CASE WHEN A1.KeraType = '매입' THEN SUM(A1.iKongKep) / SUM(A1.iWeight) ");
            strSql.AppendLine("                        WHEN A1.KeraType = '매출' THEN SUM(A1.KongKep) / SUM(A1.OWeight) END)*10)/ 10 AS DANGA");
            strSql.AppendLine("        FROM INLIST A1                                                                                        ");
            strSql.AppendLine("        LEFT JOIN MESURING A2                                                                                 ");
            strSql.AppendLine("          ON A1.J_RID = A2.JUNPYOID                                                                           ");
            strSql.AppendLine("        LEFT JOIN JAJAE A3                                                                                    ");
            strSql.AppendLine("          ON A1.J_Serial = A3.J_Serial                                                                        ");
            strSql.AppendLine("       WHERE A2.KeraType <> '직송'                                                                            ");
            strSql.AppendLine("         AND A1.J_Date BETWEEN '" + sDate.Substring(0, 7) + "-01' AND '" + sDate + "'                                                  ");
            strSql.AppendLine("         AND A1.J_SERIAL IN(650,651,652,653)                                                                  ");
            strSql.AppendLine("       GROUP BY A3.DaeGubun, SUBSTRING(A1.J_Date, 1, 7), A1.KeraType                                          ");
            strSql.AppendLine("       UNION ALL                                                                                              ");
            strSql.AppendLine("      SELECT A2.Gubun1                                                                                        ");
            strSql.AppendLine("           , SUBSTRING(A1.BASDT, 1, 7) AS BASDT                                                               ");
            strSql.AppendLine("           , 'ASR'                                                                                            ");
            strSql.AppendLine("           , SUM(A1.WEIGHT) AS WEIGHT                                                                         ");
            strSql.AppendLine("           , SUM(A1.PRICE) AS PRICE                                                                           ");
            strSql.AppendLine("           , NULL                                                                                             ");
            strSql.AppendLine("           , FLOOR(SUM(A1.PRICE) / SUM(A1.WEIGHT) * 10)/ 10 AS DANGA                                          ");
            strSql.AppendLine("        FROM J_MAGAM A1                                                                                       ");
            strSql.AppendLine("        LEFT JOIN JAJAE A2                                                                                    ");
            strSql.AppendLine("          ON A1.J_SERIAL = A2.J_SERIAL                                                                        ");
            strSql.AppendLine("       WHERE A1.BASDT BETWEEN '" + sDate.Substring(0, 7) + "-01' AND '" + sDate + "'                                                   ");
            strSql.AppendLine("         AND A1.J_SERIAL = 2025163                                                                            ");
            strSql.AppendLine("       GROUP BY A2.Gubun1, SUBSTRING(A1.BASDT, 1, 7)                                                          ");
            strSql.AppendLine("       UNION ALL                                                                                              ");
            strSql.AppendLine("      SELECT A1.Gubun1                                                                                        ");
            strSql.AppendLine("           , SUBSTRING(A2.BASDT, 1, 7) AS BASDT                                                               ");
            strSql.AppendLine("           , '재고이동'                                                                                       ");
            strSql.AppendLine("           , SUM(A2.WEIGHT) AS WEIGHT                                                                         ");
            strSql.AppendLine("           , SUM(A2.PRICE) AS PRICE                                                                           ");
            strSql.AppendLine("           , NULL                                                                                             ");
            strSql.AppendLine("           , FLOOR(SUM(A2.PRICE) / SUM(A2.WEIGHT) * 10)/ 10 AS DANGA                                          ");
            strSql.AppendLine("        FROM JAJAE A1                                                                                         ");
            strSql.AppendLine("        LEFT JOIN J_MAGAM A2                                                                                  ");
            strSql.AppendLine("          ON A1.J_Serial = A2.J_SERIAL                                                                        ");
            strSql.AppendLine("       WHERE A1.DaeGubun = '재고이동'                                                                         ");
            strSql.AppendLine("         AND A2.BASDT BETWEEN '" + sDate.Substring(0, 7) + "-01' AND '" + sDate + "'                                                   ");
            strSql.AppendLine("       GROUP BY A1.J_Serial, A1.Gubun1, SUBSTRING(A2.BASDT, 1, 7)                                             ");
            strSql.AppendLine("       UNION ALL                                                                                              ");
            strSql.AppendLine("      SELECT '슈레더비철'                                                                                     ");
            strSql.AppendLine("           , SUBSTRING(A2.BASDT, 1, 7) AS BASDT                                                               ");
            strSql.AppendLine("           , '재고이동'                                                                                       ");
            strSql.AppendLine("           , SUM(A2.WEIGHT) AS WEIGHT                                                                         ");
            strSql.AppendLine("           , SUM(A2.PRICE) AS PRICE                                                                           ");
            strSql.AppendLine("           , NULL                                                                                             ");
            strSql.AppendLine("           , FLOOR(SUM(A2.PRICE) / SUM(A2.WEIGHT) * 10)/ 10 AS DANGA                                          ");
            strSql.AppendLine("        FROM JAJAE A1                                                                                         ");
            strSql.AppendLine("        LEFT JOIN J_MAGAM A2                                                                                  ");
            strSql.AppendLine("          ON A1.J_Serial = A2.J_SERIAL                                                                        ");
            strSql.AppendLine("       WHERE A1.DaeGubun = '슈레더'                                                                           ");
            strSql.AppendLine("         AND A1.J_SERIAL IN(650,651,652,653)                                                                  ");
            strSql.AppendLine("         AND A2.BASDT BETWEEN '" + sDate.Substring(0, 7) + "-01' AND '" + sDate + "'                                                   ");
            strSql.AppendLine("       GROUP BY SUBSTRING(A2.BASDT, 1, 7)                                                                     ");
            strSql.AppendLine("       UNION ALL                                                                                              ");
            strSql.AppendLine("      SELECT A2.Gubun1                                                                                        ");
            strSql.AppendLine("           , A1.BASYM                                                                                         ");
            strSql.AppendLine("           , '월초재고'                                                                                       ");
            strSql.AppendLine("           , A1.WEIGHT                                                                                        ");
            strSql.AppendLine("           , A1.PRICE                                                                                         ");
            strSql.AppendLine("           , NULL                                                                                             ");
            strSql.AppendLine("           , A1.DANGA                                                                                         ");
            strSql.AppendLine("        FROM YYJEGO A1                                                                                        ");
            strSql.AppendLine("        LEFT JOIN JAJAE A2                                                                                    ");
            strSql.AppendLine("          ON A1.J_SERIAL = A2.J_Serial                                                                        ");
            strSql.AppendLine("         AND A2.DaeGubun = '재고'                                                                             ");
            strSql.AppendLine("       WHERE A1.BASYM = '" + sDate.Substring(0, 7) + "'                                                                             ");
            strSql.AppendLine(" )                                                                                                            ");
            strSql.AppendLine("                                                                                                              ");
            strSql.AppendLine(" SELECT A1.GUBUN                                                                                              ");
            strSql.AppendLine("      , A1.BASDT                                                                                              ");
            strSql.AppendLine("      , A1.KERATYPE                                                                                           ");
            strSql.AppendLine("      , FORMAT(A1.WEIGHT, N'#,0') AS WEIGHT                                                                   ");
            strSql.AppendLine("      , FORMAT(A1.KONGKEP, N'#,0') AS KONGKEP                                                                 ");
            strSql.AppendLine("      , FORMAT(A1.LOSSWEIGHT, N'#,0') AS LOSSWEIGHT                                                           ");
            strSql.AppendLine("      , FORMAT(A1.DANGA, N'#,0.#') AS DANGA                                                                   ");
            strSql.AppendLine("   FROM TEMP1 A1");

            DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //월말재고
            double d1WEIGHT1 = 0;
            double d1DANGA1 = 0;
            //double d1KONGKEP1 = 0;
            double d1WEIGHT2 = 0;
            double d1DANGA2 = 0;
            //double d1KONGKEP2 = 0;
            double d1WEIGHT3 = 0;
            double d1DANGA3 = 0;
            double d1KONGKEP3 = 0;
            double d1WEIGHT4 = 0;
            double d1DANGA4 = 0;
            //double d1KONGKEP4 = 0;
            //double d1WEIGHT5 = 0;
            //double d1DANGA5 = 0;
            //double d1KONGKEP5 = 0;
            double d1WEIGHT6 = 0;

            DataTable dt = new DataTable();
            dt.Columns.Add("TITLE");
            dt.Columns.Add("WEIGHT1");
            dt.Columns.Add("DANGA1");
            dt.Columns.Add("KONGKEP1");
            dt.Columns.Add("WEIGHT2");
            dt.Columns.Add("DANGA2");
            dt.Columns.Add("KONGKEP2");
            dt.Columns.Add("WEIGHT3");
            dt.Columns.Add("DANGA3");
            dt.Columns.Add("KONGKEP3");
            dt.Columns.Add("WEIGHT4");
            dt.Columns.Add("DANGA4");
            dt.Columns.Add("KONGKEP4");
            dt.Columns.Add("WEIGHT5");
            dt.Columns.Add("DANGA5");
            dt.Columns.Add("KONGKEP5");
            dt.Columns.Add("WEIGHT6");
            dt.Columns.Add("DANGA6");
            dt.Columns.Add("KONGKEP6");

            DataRow row = dt.NewRow();
            row["TITLE"] = "월초재고";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                double dKONGKEP1 = 0;
                double dDANGA1 = 0;
                double dWEIGHT2 = 0;
                double dKONGKEP2 = 0;
                double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                double dKONGKEP3 = 0;
                double dDANGA3 = 0;
                //double dWEIGHT4 = 0;
                //double dKONGKEP4 = 0;
                //double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("월초재고"))
                    {
                        if (sGUBUN.Equals("슈레더"))
                        {
                            row["WEIGHT1"] = sWEIGHT;
                            row["DANGA1"] = sDANGA;
                            row["KONGKEP1"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT1);
                            double.TryParse(sDANGA, out dDANGA1);
                            double.TryParse(sKONGKEP, out dKONGKEP1);

                            d1WEIGHT1 = dWEIGHT1;
                        }

                        if (sGUBUN.Equals("슈레더비철"))
                        {
                            row["WEIGHT2"] = sWEIGHT;
                            row["DANGA2"] = sDANGA;
                            row["KONGKEP2"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT2);
                            double.TryParse(sDANGA, out dDANGA2);
                            double.TryParse(sKONGKEP, out dKONGKEP2);

                            d1WEIGHT2 = dWEIGHT2;
                        }

                        if (sGUBUN.Equals("고철A"))
                        {
                            row["WEIGHT3"] = sWEIGHT;
                            row["DANGA3"] = sDANGA;
                            row["KONGKEP3"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT3);
                            double.TryParse(sDANGA, out dDANGA3);
                            double.TryParse(sKONGKEP, out dKONGKEP3);

                            d1WEIGHT3 = dWEIGHT3;
                        }

                        //if (sGUBUN.Equals("고철B"))
                        //{
                        //    row["WEIGHT4"] = sWEIGHT;
                        //    row["DANGA4"] = sDANGA;
                        //    row["KONGKEP4"] = sKONGKEP;

                        //    double.TryParse(sWEIGHT, out dWEIGHT4);
                        //    double.TryParse(sDANGA, out dDANGA4);
                        //    double.TryParse(sKONGKEP, out dKONGKEP4);

                        //    d1WEIGHT4 = dWEIGHT4;
                        //}

                        //if (sGUBUN.Equals("스크랩비철"))
                        //{
                        //    row["WEIGHT5"] = sWEIGHT;
                        //    row["DANGA5"] = sDANGA;
                        //    row["KONGKEP5"] = sKONGKEP;

                        //    double.TryParse(sWEIGHT, out dWEIGHT5);
                        //    double.TryParse(sDANGA, out dDANGA5);
                        //    double.TryParse(sKONGKEP, out dKONGKEP5);

                        //    d1WEIGHT5 = dWEIGHT5;
                        //}

                    }

                    if (i == dtResult.Rows.Count - 1)
                    {
                        double dWEIGHT6 = dWEIGHT1 + dWEIGHT3;//+ dWEIGHT2; //+ dWEIGHT4; //+ dWEIGHT5;
                        double dKONGKEP6 = dKONGKEP1 + dKONGKEP3;//+ dKONGKEP2; //+ dKONGKEP4; //+ dKONGKEP5;
                        double dDANGA6 = Math.Truncate(dKONGKEP6 / dWEIGHT6 * 10) / 10;

                        row["WEIGHT6"] = dWEIGHT6 == 0 ? "" : dWEIGHT6.ToString("n0");
                        row["DANGA6"] = double.IsNaN(dDANGA6) ? "" : dDANGA6.ToString("n1");
                        row["KONGKEP6"] = dKONGKEP6 == 0 ? "" : dKONGKEP6.ToString("n0");

                        d1WEIGHT6 = dWEIGHT6;
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "매출";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                double dKONGKEP1 = 0;
                double dDANGA1 = 0;
                double dWEIGHT2 = 0;
                double dKONGKEP2 = 0;
                double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                double dKONGKEP3 = 0;
                double dDANGA3 = 0;
                double dWEIGHT4 = 0;
                double dKONGKEP4 = 0;
                double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("매출"))
                    {
                        if (sGUBUN.Equals("슈레더"))
                        {
                            row["WEIGHT1"] = sWEIGHT;
                            row["DANGA1"] = sDANGA;
                            row["KONGKEP1"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT1);
                            double.TryParse(sDANGA, out dDANGA1);
                            double.TryParse(sKONGKEP, out dKONGKEP1);

                            d1WEIGHT1 -= dWEIGHT1;
                        }

                        if (sGUBUN.Equals("슈레더비철"))
                        {
                            row["WEIGHT2"] = sWEIGHT;
                            row["DANGA2"] = sDANGA;
                            row["KONGKEP2"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT2);
                            double.TryParse(sDANGA, out dDANGA2);
                            double.TryParse(sKONGKEP, out dKONGKEP2);

                            d1WEIGHT2 -= dWEIGHT2;
                        }

                        if (sGUBUN.Equals("고철A"))
                        {
                            row["WEIGHT3"] = sWEIGHT;
                            row["DANGA3"] = sDANGA;
                            row["KONGKEP3"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT3);
                            double.TryParse(sDANGA, out dDANGA3);
                            double.TryParse(sKONGKEP, out dKONGKEP3);

                            d1WEIGHT3 -= dWEIGHT3;
                        }

                        if (sGUBUN.Equals("고철B"))
                        {
                            row["WEIGHT4"] = sWEIGHT;
                            row["DANGA4"] = sDANGA;
                            row["KONGKEP4"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT4);
                            double.TryParse(sDANGA, out dDANGA4);
                            double.TryParse(sKONGKEP, out dKONGKEP4);

                            d1WEIGHT4 -= dWEIGHT4;
                        }

                        //if (sGUBUN.Equals("스크랩비철"))
                        //{
                        //    row["WEIGHT5"] = sWEIGHT;
                        //    row["DANGA5"] = sDANGA;
                        //    row["KONGKEP5"] = sKONGKEP;

                        //    double.TryParse(sWEIGHT, out dWEIGHT5);
                        //    double.TryParse(sDANGA, out dDANGA5);
                        //    double.TryParse(sKONGKEP, out dKONGKEP5);

                        //    d1WEIGHT5 -= dWEIGHT5;
                        //}

                    }

                    if (i == dtResult.Rows.Count - 1)
                    {
                        double dWEIGHT6 = dWEIGHT1 + dWEIGHT3 + dWEIGHT4;//+ dWEIGHT2; //+ dWEIGHT5;
                        double dKONGKEP6 = dKONGKEP1 + dKONGKEP3 + dKONGKEP4;//+ dKONGKEP2; //+ dKONGKEP5;
                        double dDANGA6 = Math.Truncate(dKONGKEP6 / dWEIGHT6 * 10) / 10;

                        row["WEIGHT6"] = dWEIGHT6 == 0 ? "" : dWEIGHT6.ToString("n0");
                        row["DANGA6"] = double.IsNaN(dDANGA6) ? "" : dDANGA6.ToString("n1");
                        row["KONGKEP6"] = dKONGKEP6 == 0 ? "" : dKONGKEP6.ToString("n0");

                        d1WEIGHT6 -= dWEIGHT6;
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "매출감량";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                //double dKONGKEP1 = 0;
                //double dDANGA1 = 0;
                double dWEIGHT2 = 0;
                //double dKONGKEP2 = 0;
                //double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                //double dKONGKEP3 = 0;
                //double dDANGA3 = 0;
                double dWEIGHT4 = 0;
                //double dKONGKEP4 = 0;
                //double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("매출"))
                    {
                        if (sGUBUN.Equals("슈레더"))
                        {
                            row["WEIGHT1"] = sLOSSWEIGHT;

                            double.TryParse(sLOSSWEIGHT, out dWEIGHT1);

                            d1WEIGHT1 -= dWEIGHT1;
                        }

                        //if (sGUBUN.Equals("슈레더비철"))
                        //{
                        //    row["WEIGHT2"] = sLOSSWEIGHT;

                        //    double.TryParse(sLOSSWEIGHT, out dWEIGHT2);

                        //    d1WEIGHT2 -= dWEIGHT2;
                        //}

                        if (sGUBUN.Equals("고철A"))
                        {
                            row["WEIGHT3"] = sLOSSWEIGHT;

                            double.TryParse(sLOSSWEIGHT, out dWEIGHT3);

                            d1WEIGHT3 -= dWEIGHT3;
                        }

                        if (sGUBUN.Equals("고철B"))
                        {
                            row["WEIGHT4"] = sLOSSWEIGHT;

                            double.TryParse(sLOSSWEIGHT, out dWEIGHT4);

                            d1WEIGHT4 -= dWEIGHT4;
                        }

                        //if (sGUBUN.Equals("스크랩비철"))
                        //{
                        //    row["WEIGHT5"] = sLOSSWEIGHT;

                        //    double.TryParse(sLOSSWEIGHT, out dWEIGHT5);
                        //}

                    }

                    if (i == dtResult.Rows.Count - 1)
                    {
                        double dWEIGHT6 = dWEIGHT1 + dWEIGHT3 + dWEIGHT4;//+ dWEIGHT2;// + dWEIGHT5;

                        row["WEIGHT6"] = dWEIGHT6 == 0 ? "" : dWEIGHT6.ToString("n0");
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "ASR, 폐기물처리";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                double dKONGKEP1 = 0;
                double dDANGA1 = 0;
                //double dWEIGHT2 = 0;
                //double dKONGKEP2 = 0;
                //double dDANGA2 = 0;
                //double dWEIGHT3 = 0;
                //double dKONGKEP3 = 0;
                //double dDANGA3 = 0;
                //double dWEIGHT4 = 0;
                //double dKONGKEP4 = 0;
                //double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("ASR"))
                    {
                        row["WEIGHT1"] = sWEIGHT;
                        row["DANGA1"] = sDANGA;
                        row["KONGKEP1"] = sKONGKEP;

                        double.TryParse(sWEIGHT, out dWEIGHT1);
                        double.TryParse(sDANGA, out dDANGA1);
                        double.TryParse(sKONGKEP, out dKONGKEP1);

                        d1WEIGHT1 -= dWEIGHT1;

                        double dWEIGHT6 = dWEIGHT1;
                        double dKONGKEP6 = dKONGKEP1;
                        double dDANGA6 = Math.Truncate(dKONGKEP6 / dWEIGHT6 * 10) / 10;

                        row["WEIGHT6"] = dWEIGHT6 == 0 ? "" : dWEIGHT6.ToString("n0");
                        row["DANGA6"] = double.IsNaN(dDANGA6) ? "" : dDANGA6.ToString("n1");
                        row["KONGKEP6"] = dKONGKEP6 == 0 ? "" : dKONGKEP6.ToString("n0");

                        d1WEIGHT6 -= dWEIGHT6;
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "재고감모손실";
            if (dtResult != null)
            {
                //double dWEIGHT1 = 0;
                //double dKONGKEP1 = 0;
                //double dDANGA1 = 0;
                double dWEIGHT2 = 0;
                double dKONGKEP2 = 0;
                double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                double dKONGKEP3 = 0;
                double dDANGA3 = 0;
                double dWEIGHT4 = 0;
                double dKONGKEP4 = 0;
                double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("매입"))
                    {
                        if (sGUBUN.Equals("고철A"))
                        {
                            double.TryParse(sWEIGHT, out dWEIGHT3);
                            double.TryParse(sDANGA, out dDANGA3);

                            dWEIGHT3 = Math.Truncate(dWEIGHT3 * 0.02);
                            dKONGKEP3 = Math.Truncate(dWEIGHT3 * dDANGA3);

                            row["WEIGHT3"] = double.IsNaN(dWEIGHT3) ? "" : dWEIGHT3.ToString("n0");
                            row["DANGA3"] = sDANGA;
                            row["KONGKEP3"] = double.IsNaN(dKONGKEP3) ? "" : dKONGKEP3.ToString("n0");

                            //d1WEIGHT3 -= dWEIGHT3; //재고감모손실은 합계에서 제외 2022-02-18
                        }

                        if (sGUBUN.Equals("고철B"))
                        {
                            double.TryParse(sWEIGHT, out dWEIGHT4);
                            double.TryParse(sDANGA, out dDANGA4);

                            dWEIGHT4 = Math.Truncate(dWEIGHT4 * 0.02);
                            dKONGKEP4 = Math.Truncate(dWEIGHT4 * dDANGA4);

                            row["WEIGHT4"] = double.IsNaN(dWEIGHT4) ? "" : dWEIGHT4.ToString("n0");
                            row["DANGA4"] = sDANGA;
                            row["KONGKEP4"] = double.IsNaN(dKONGKEP4) ? "" : dKONGKEP4.ToString("n0");

                            //d1WEIGHT4 -= dWEIGHT4;
                        }

                        //if (sGUBUN.Equals("스크랩비철"))
                        //{
                        //    double.TryParse(sWEIGHT, out dWEIGHT5);
                        //    double.TryParse(sDANGA, out dDANGA5);

                        //    dWEIGHT5 = Math.Truncate(dWEIGHT5 * 0.02, 0);
                        //    dKONGKEP5 = Math.Truncate(dWEIGHT5 * dDANGA5);

                        //    row["WEIGHT5"] = dWEIGHT5.ToString("n0");
                        //    row["DANGA5"] = sDANGA;
                        //    row["KONGKEP5"] = dKONGKEP5.ToString("n0");
                        //}

                    }

                    //(현업요청)슈레더 비철 재고이동량 재고감모손실에 표시 2022-03-02
                    if (sGUBUN.Equals("슈레더비철") && sKERATYPE.Equals("재고이동"))
                    {

                        double.TryParse(sWEIGHT, out dWEIGHT2);
                        double.TryParse(sDANGA, out dDANGA2);
                        double.TryParse(sKONGKEP, out dKONGKEP2);

                        row["WEIGHT2"] = double.IsNaN(dWEIGHT2) || dWEIGHT2 == 0 ? "" : dWEIGHT2.ToString("n0");
                        row["DANGA2"] = sDANGA;
                        row["KONGKEP2"] = double.IsNaN(dKONGKEP2) || dKONGKEP2 == 0 ? "" : dKONGKEP2.ToString("n0");

                        row["WEIGHT1"] = double.IsNaN(dWEIGHT2) || dWEIGHT2 == 0 ? "" : (-dWEIGHT2).ToString("n0");
                        row["DANGA1"] = sDANGA;
                        row["KONGKEP1"] = double.IsNaN(dKONGKEP2) || dKONGKEP2 == 0 ? "" : dKONGKEP2.ToString("n0");

                        d1WEIGHT1 -= dWEIGHT2;//슈레더->슈레더비철(슈레더)
                        d1WEIGHT2 += dWEIGHT2;//슈레더->슈레더비철(슈레더비철)
                    }

                    if (i == dtResult.Rows.Count - 1)
                    {
                        double dWEIGHT6 = dWEIGHT3 + dWEIGHT4;// + dWEIGHT5;
                        double dKONGKEP6 = dKONGKEP3 + dKONGKEP4;// + dKONGKEP5;
                        double dDANGA6 = Math.Truncate(dKONGKEP6 / dWEIGHT6 * 10) / 10;

                        row["WEIGHT6"] = dWEIGHT6 == 0 ? "" : dWEIGHT6.ToString("n0");
                        row["DANGA6"] = double.IsNaN(dDANGA6) ? "" : dDANGA6.ToString("n1");
                        row["KONGKEP6"] = dKONGKEP6 == 0 ? "" : dKONGKEP6.ToString("n0");

                        //d1WEIGHT6 -= dWEIGHT6;
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "매입";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                double dKONGKEP1 = 0;
                double dDANGA1 = 0;
                double dWEIGHT2 = 0;
                double dKONGKEP2 = 0;
                double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                double dKONGKEP3 = 0;
                double dDANGA3 = 0;
                double dWEIGHT4 = 0;
                double dKONGKEP4 = 0;
                double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("매입"))
                    {
                        if (sGUBUN.Equals("슈레더"))
                        {
                            row["WEIGHT1"] = sWEIGHT;
                            row["DANGA1"] = sDANGA;
                            row["KONGKEP1"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT1);
                            double.TryParse(sDANGA, out dDANGA1);
                            double.TryParse(sKONGKEP, out dKONGKEP1);

                            d1WEIGHT1 += dWEIGHT1;
                            d1DANGA1 = dDANGA1;
                        }

                        //if (sGUBUN.Equals("슈레더비철"))
                        //{
                        //    row["WEIGHT2"] = sWEIGHT;
                        //    row["DANGA2"] = sDANGA;
                        //    row["KONGKEP2"] = sKONGKEP;

                        //    double.TryParse(sWEIGHT, out dWEIGHT2);
                        //    double.TryParse(sDANGA, out dDANGA2);
                        //    double.TryParse(sKONGKEP, out dKONGKEP2);

                        //    d1WEIGHT2 += dWEIGHT2;
                        //    d1DANGA2 = dDANGA2;
                        //}

                        if (sGUBUN.Equals("고철A"))
                        {
                            row["WEIGHT3"] = sWEIGHT;
                            row["DANGA3"] = sDANGA;
                            row["KONGKEP3"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT3);
                            double.TryParse(sDANGA, out dDANGA3);
                            double.TryParse(sKONGKEP, out dKONGKEP3);

                            d1WEIGHT3 += dWEIGHT3;
                            //d1DANGA3 = dDANGA3;
                        }

                        if (sGUBUN.Equals("고철B"))
                        {
                            row["WEIGHT4"] = sWEIGHT;
                            row["DANGA4"] = sDANGA;
                            row["KONGKEP4"] = sKONGKEP;

                            double.TryParse(sWEIGHT, out dWEIGHT4);
                            double.TryParse(sDANGA, out dDANGA4);
                            double.TryParse(sKONGKEP, out dKONGKEP4);

                            d1WEIGHT4 += dWEIGHT4;
                            //d1DANGA4 = dDANGA4;
                        }

                        //if (sGUBUN.Equals("스크랩비철"))
                        //{
                        //    row["WEIGHT5"] = sWEIGHT;
                        //    row["DANGA5"] = sDANGA;
                        //    row["KONGKEP5"] = sKONGKEP;

                        //    double.TryParse(sWEIGHT, out dWEIGHT5);
                        //    double.TryParse(sDANGA, out dDANGA5);
                        //    double.TryParse(sKONGKEP, out dKONGKEP5);

                        //    d1WEIGHT5 += dWEIGHT5;
                        //    d1DANGA5 = dDANGA5;
                        //}

                    }

                    if (i == dtResult.Rows.Count - 1)
                    {
                        double dWEIGHT6 = dWEIGHT1 + dWEIGHT3 + dWEIGHT4; //+dWEIGHT2; //+ dWEIGHT5;
                        double dKONGKEP6 = dKONGKEP1 + dKONGKEP3 + dKONGKEP4;//+ dKONGKEP2; //+ dKONGKEP5;
                        double dDANGA6 = Math.Truncate(dKONGKEP6 / dWEIGHT6 * 10) / 10;

                        row["WEIGHT6"] = dWEIGHT6 == 0 ? "" : dWEIGHT6.ToString("n0");
                        row["DANGA6"] = double.IsNaN(dDANGA6) ? "" : dDANGA6.ToString("n1");
                        row["KONGKEP6"] = dKONGKEP6 == 0 ? "" : dKONGKEP6.ToString("n0");

                        d1WEIGHT6 += dWEIGHT6;

                        d1DANGA3 = Math.Truncate((dKONGKEP3 + dKONGKEP4) / (dWEIGHT3 + dWEIGHT4) * 10) / 10;
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "슈 -> 스";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                double dKONGKEP1 = 0;
                double dDANGA1 = 0;
                //double dWEIGHT2 = 0;
                //double dKONGKEP2 = 0;
                //double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                double dKONGKEP3 = 0;
                double dDANGA3 = 0;
                double dWEIGHT4 = 0;
                double dKONGKEP4 = 0;
                double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("재고이동"))
                    {
                        if (sGUBUN.Equals("슈-고철A"))
                        {
                            double.TryParse(sWEIGHT, out dWEIGHT3);
                            double.TryParse(sKONGKEP, out dKONGKEP3);
                            double.TryParse(sDANGA, out dDANGA3);

                            row["WEIGHT3"] = dWEIGHT3.ToString("n0");
                            row["KONGKEP3"] = dKONGKEP3.ToString("n0");
                            row["DANGA3"] = dDANGA3.ToString("n1");

                            d1WEIGHT1 -= dWEIGHT3;
                            d1WEIGHT3 += dWEIGHT3;
                        }

                        if (sGUBUN.Equals("슈-고철B"))
                        {
                            double.TryParse(sWEIGHT, out dWEIGHT4);
                            double.TryParse(sKONGKEP, out dKONGKEP4);
                            double.TryParse(sDANGA, out dDANGA4);

                            row["WEIGHT4"] = dWEIGHT4.ToString("n0");
                            row["KONGKEP4"] = dKONGKEP4.ToString("n0");
                            row["DANGA4"] = dDANGA4.ToString("n1");

                            d1WEIGHT1 -= dWEIGHT4;
                            d1WEIGHT4 += dWEIGHT4;
                        }

                        dWEIGHT1 = -(dWEIGHT3 + dWEIGHT4);
                        dKONGKEP1 = -(dKONGKEP3 + dKONGKEP4);
                        dDANGA1 = Math.Truncate(dKONGKEP1 / dWEIGHT1 * 10) / 10;

                        row["WEIGHT1"] = dWEIGHT1 == 0 ? "" : dWEIGHT1.ToString("n0");
                        row["DANGA1"] = double.IsNaN(dDANGA1) ? "" : dDANGA1.ToString("n1");
                        row["KONGKEP1"] = dKONGKEP1 == 0 ? "" : dKONGKEP1.ToString("n0");
                    }
                }
            }

            dt.Rows.Add(row);

            row = dt.NewRow();
            row["TITLE"] = "스 -> 슈";
            if (dtResult != null)
            {
                double dWEIGHT1 = 0;
                double dKONGKEP1 = 0;
                double dDANGA1 = 0;
                //double dWEIGHT2 = 0;
                //double dKONGKEP2 = 0;
                //double dDANGA2 = 0;
                double dWEIGHT3 = 0;
                double dKONGKEP3 = 0;
                double dDANGA3 = 0;
                double dWEIGHT4 = 0;
                double dKONGKEP4 = 0;
                double dDANGA4 = 0;
                //double dWEIGHT5 = 0;
                //double dKONGKEP5 = 0;
                //double dDANGA5 = 0;

                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
                    string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
                    string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
                    string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
                    string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
                    string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

                    if (sKERATYPE.Equals("재고이동"))
                    {
                        if (sGUBUN.Equals("고철A-슈"))
                        {
                            double.TryParse(sWEIGHT, out dWEIGHT3);
                            double.TryParse(sKONGKEP, out dKONGKEP3);
                            double.TryParse(sDANGA, out dDANGA3);

                            row["WEIGHT3"] = (-dWEIGHT3).ToString("n0");
                            row["KONGKEP3"] = (-dKONGKEP3).ToString("n0");
                            row["DANGA3"] = dDANGA3.ToString("n1");

                            d1WEIGHT1 += dWEIGHT3;
                            d1WEIGHT3 -= dWEIGHT3;
                        }

                        if (sGUBUN.Equals("고철B-슈"))
                        {
                            double.TryParse(sWEIGHT, out dWEIGHT4);
                            double.TryParse(sKONGKEP, out dKONGKEP4);
                            double.TryParse(sDANGA, out dDANGA4);

                            row["WEIGHT3"] = (-dWEIGHT4).ToString("n0");
                            row["KONGKEP3"] = (-dKONGKEP4).ToString("n0");
                            row["DANGA3"] = dDANGA4.ToString("n1");

                            d1WEIGHT1 += dWEIGHT4;
                            d1WEIGHT4 -= dWEIGHT4;
                        }

                        dWEIGHT1 = dWEIGHT3 + dWEIGHT4;
                        dKONGKEP1 = dKONGKEP3 + dKONGKEP4;
                        dDANGA1 = Math.Truncate(dKONGKEP1 / dWEIGHT1 * 10) / 10;

                        row["WEIGHT1"] = dWEIGHT1 == 0 ? "" : dWEIGHT1.ToString("n0");
                        row["DANGA1"] = double.IsNaN(dDANGA1) ? "" : dDANGA1.ToString("n1");
                        row["KONGKEP1"] = dKONGKEP1 == 0 ? "" : dKONGKEP1.ToString("n0");
                    }
                }
            }

            dt.Rows.Add(row);

            #region 고철A<->고철B 고철재고 합산해서 관리하는걸로 변경되면서 사용x
            //row = dt.NewRow();
            //row["TITLE"] = "고철A <-> 고철B";
            //if (dtResult != null)
            //{
            //    //double dWEIGHT1 = 0;
            //    //double dKONGKEP1 = 0;
            //    //double dDANGA1 = 0;
            //    //double dWEIGHT2 = 0;
            //    //double dKONGKEP2 = 0;
            //    //double dDANGA2 = 0;
            //    double dWEIGHT3 = 0;
            //    double dKONGKEP3 = 0;
            //    double dDANGA3 = 0;
            //    double dWEIGHT4 = 0;
            //    double dKONGKEP4 = 0;
            //    double dDANGA4 = 0;
            //    //double dWEIGHT5 = 0;
            //    //double dKONGKEP5 = 0;
            //    //double dDANGA5 = 0;

            //    double dWEIGHT3_1 = 0;
            //    double dWEIGHT4_1 = 0;

            //    double dKONGKEP3_1 = 0;
            //    double dKONGKEP4_1 = 0;

            //    for (int i = 0; i < dtResult.Rows.Count; i++)
            //    {
            //        string sGUBUN = dtResult.Rows[i]["GUBUN"]?.ToString();
            //        string sKERATYPE = dtResult.Rows[i]["KERATYPE"]?.ToString();
            //        string sWEIGHT = dtResult.Rows[i]["WEIGHT"]?.ToString();
            //        string sKONGKEP = dtResult.Rows[i]["KONGKEP"]?.ToString();
            //        string sLOSSWEIGHT = dtResult.Rows[i]["LOSSWEIGHT"]?.ToString();
            //        string sDANGA = dtResult.Rows[i]["DANGA"]?.ToString();

            //        if (sKERATYPE.Equals("재고이동"))
            //        {
            //            if (sGUBUN.Equals("고철A-고철B"))
            //            {
            //                double.TryParse(sWEIGHT, out dWEIGHT3);
            //                double.TryParse(sKONGKEP, out dKONGKEP3);

            //                d1WEIGHT3 -= dWEIGHT3;
            //                d1WEIGHT4 += dWEIGHT3;
            //            }

            //            if (sGUBUN.Equals("고철B-고철A"))
            //            {
            //                double.TryParse(sWEIGHT, out dWEIGHT4);
            //                double.TryParse(sKONGKEP, out dKONGKEP4);

            //                d1WEIGHT3 += dWEIGHT4;
            //                d1WEIGHT4 -= dWEIGHT4;
            //            }

            //            dWEIGHT3_1 = dWEIGHT4 - dWEIGHT3;
            //            dWEIGHT4_1 = dWEIGHT3 - dWEIGHT4;

            //            dKONGKEP3_1 = dKONGKEP4 - dKONGKEP3;
            //            dKONGKEP4_1 = dKONGKEP3 - dKONGKEP4;

            //            dDANGA3 = Math.Truncate(dKONGKEP3_1 / dWEIGHT3_1*10)/10;
            //            dDANGA4 = Math.Truncate(dKONGKEP4_1 / dWEIGHT4_1*10)/10;

            //            row["WEIGHT3"] = dWEIGHT3_1 == 0 ? "" : dWEIGHT3_1.ToString("n0");
            //            row["WEIGHT4"] = dWEIGHT4_1 == 0 ? "" : dWEIGHT4_1.ToString("n0");

            //            row["KONGKEP3"] = dKONGKEP3_1 == 0 ? "" : dKONGKEP3_1.ToString("n0");
            //            row["KONGKEP4"] = dKONGKEP4_1 == 0 ? "" : dKONGKEP4_1.ToString("n0");

            //            row["DANGA3"] = double.IsNaN(dDANGA3) ? "" : dDANGA3.ToString("n1");
            //            row["DANGA4"] = double.IsNaN(dDANGA4) ? "" : dDANGA4.ToString("n1");
            //        }
            //    }
            //}

            //dt.Rows.Add(row);
            #endregion

            row = dt.NewRow();
            row["TITLE"] = "월말재고";

            row["WEIGHT1"] = d1WEIGHT1 == 0 ? "" : d1WEIGHT1.ToString("n0");
            row["DANGA1"] = d1DANGA1 == 0 ? "" : d1DANGA1.ToString("n1");
            double d1KONGKEP1 = Math.Truncate(d1WEIGHT1 * d1DANGA1);
            row["KONGKEP1"] = double.IsNaN(d1KONGKEP1) || d1KONGKEP1 == 0 ? "" : d1KONGKEP1.ToString("n0");

            row["WEIGHT2"] = d1WEIGHT2 == 0 ? "" : d1WEIGHT2.ToString("n0");
            row["DANGA2"] = d1DANGA2 == 0 ? "" : d1DANGA2.ToString("n1");
            double d1KONGKEP2 = Math.Truncate(d1WEIGHT2 * d1DANGA2);
            row["KONGKEP2"] = double.IsNaN(d1KONGKEP2) || d1KONGKEP2 == 0 ? "" : d1KONGKEP2.ToString("n0");

            //고철A 고철B 재고 합해서 관리로 변경 2022-02-18
            d1WEIGHT3 = d1WEIGHT3 + d1WEIGHT4;
            d1KONGKEP3 = Math.Truncate(d1WEIGHT3 * d1DANGA3);

            row["WEIGHT3"] = d1WEIGHT3 == 0 ? "" : d1WEIGHT3.ToString("n0");
            row["DANGA3"] = double.IsNaN(d1DANGA3) || d1DANGA3 == 0 ? "" : d1DANGA3.ToString("n1");
            row["KONGKEP3"] = double.IsNaN(d1KONGKEP3) || d1KONGKEP3 == 0 ? "" : d1KONGKEP3.ToString("n0");

            #region 따로 관리ver
            //row["WEIGHT3"] = d1WEIGHT3 == 0 ? "" : d1WEIGHT3.ToString("n0");
            //row["DANGA3"] = d1DANGA3 == 0 ? "" : d1DANGA3.ToString("n1");
            //double d1KONGKEP3 = Math.Truncate(d1WEIGHT3 * d1DANGA3);
            //row["KONGKEP3"] = d1KONGKEP3 == 0 ? "" : d1KONGKEP3.ToString("n0");

            //row["WEIGHT4"] = d1WEIGHT4 == 0 ? "" : d1WEIGHT4.ToString("n0");
            //row["DANGA4"] = d1DANGA4 == 0 ? "" : d1DANGA4.ToString("n1");
            //double d1KONGKEP4 = Math.Truncate(d1WEIGHT4 * d1DANGA4);
            //row["KONGKEP4"] = d1KONGKEP4 == 0 ? "" : d1KONGKEP4.ToString("n0");
            #endregion

            //row["WEIGHT5"] = d1WEIGHT5 == 0 ? "" : d1WEIGHT5.ToString("n0");
            //row["DANGA5"] = d1DANGA5 == 0 ? "" : d1DANGA5.ToString("n1");
            //double d1KONGKEP5 = Math.Truncate(d1WEIGHT5 * d1DANGA5);
            //row["KONGKEP5"] = d1KONGKEP5 == 0 ? "" : d1KONGKEP5.ToString("n0");

            double d1KONGKEP6 = d1KONGKEP1 + d1KONGKEP3; //+d1KONGKEP2; //+ d1KONGKEP4; //+ d1KONGKEP5;
            double d1DANGA6 = Math.Truncate(d1KONGKEP6 / d1WEIGHT6 * 10) / 10;

            row["WEIGHT6"] = d1WEIGHT6 == 0 ? "" : d1WEIGHT6.ToString("n0");
            row["DANGA6"] = double.IsNaN(d1DANGA6) ? "" : d1DANGA6.ToString("n1");
            row["KONGKEP6"] = double.IsNaN(d1KONGKEP6) || d1KONGKEP6 == 0 ? "" : d1KONGKEP6.ToString("n0");

            dt.Rows.Add(row);
            BGridYMRetr.DataSource = dt;
        }
        #endregion

        #region 영업부-실적
        //주간
        private void SetRetr6()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);
            string sYM = sTo.Substring(0, 7).Replace("-", "");

            int jucha = ComnEtcFunc.GetWeekOfMonth(DateTime.Parse(sTo), null);

            StringBuilder strSql = new StringBuilder();

            //영업팀 목록
            strSql.Clear();
            strSql.AppendLine(" SELECT A1.NAME      ");
            strSql.AppendLine("   FROM SALEMAEIP A1 ");
            strSql.AppendLine(" WHERE YYMM = '" + sYM + "'");
            strSql.AppendLine("  GROUP BY A1.NAME   ");

            DataTable dtSawon = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            ArrayList nameList = new ArrayList();

            for (int i = 0; i < dtSawon.Rows.Count; i++)
            {
                string name = dtSawon.Rows[i]["NAME"].ToString();

                if (name.Equals("기타"))
                    continue;

                nameList.Add(name);
            }

            nameList.Add("합계");

            ComnGridFunc.BeginInitialize(BGridRetr1);

            BGridViewRetr1.Bands.Clear();

            ComnGridFunc.AddView(BGridRetr1, BGridViewRetr1, true);

            ComnGridFunc.BeginInitialize(BGridViewRetr1);

            GridBand gridBand1 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();

            BandedGridColumn title1 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            BandedGridColumn title2 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

            gridBand1.AppearanceHeader.Options.UseTextOptions = true;
            gridBand1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridBand1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gridBand1.Caption = jucha + " 주차";
            gridBand1.Columns.Add(title1);
            gridBand1.Columns.Add(title2);
            gridBand1.Width = 150;
            gridBand1.OptionsBand.AllowMove = false;

            title1.AppearanceCell.Options.UseTextOptions = true;
            title1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            title1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            title1.Caption = "TITLE1";
            title1.FieldName = "TITLE1";
            title1.OptionsColumn.AllowEdit = false;
            title1.Visible = true;

            ComnGridFunc.AddBand(BGridViewRetr1, gridBand1);

            ComnGridFunc.AddGridColumn
            (
                BGridViewRetr1,
                gridBand1,
                title1
            );

            title2.AppearanceCell.Options.UseTextOptions = true;
            title2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            title2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            title2.Caption = "TITLE2";
            title2.FieldName = "TITLE2";
            title2.OptionsColumn.AllowEdit = false;
            title2.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            title2.Visible = true;

            ComnGridFunc.AddGridColumn
            (
                BGridViewRetr1,
                gridBand1,
                title2
            );

            GridBand gridBand2 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();

            gridBand2.AppearanceHeader.Options.UseTextOptions = true;
            gridBand2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridBand2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gridBand2.Caption = "영업부 - 주간 실적표";

            ComnGridFunc.AddBand(BGridViewRetr1, gridBand2);

            GridBand[] bList = new GridBand[dtSawon.Rows.Count];
            if(dtSawon.Rows.Count > 0)
            {
                for (int i = 0; i < nameList.Count; i++)
                {
                    string name = nameList[i].ToString();

                    GridBand gridBandSub = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
                    BandedGridColumn usr = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

                    gridBandSub.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandSub.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandSub.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandSub.Caption = name;
                    gridBandSub.Columns.Add(usr);
                    gridBandSub.Name = name + i;
                    gridBandSub.Width = 75;

                    usr.AppearanceCell.Options.UseTextOptions = true;
                    usr.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    usr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    usr.Caption = name;
                    usr.FieldName = name;
                    usr.Name = name + "_sub" + i;
                    usr.OptionsColumn.AllowEdit = false;
                    usr.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                    usr.Visible = true;

                    ComnGridFunc.AddChildGridBand(gridBand2, gridBandSub);

                    ComnGridFunc.AddGridColumn
                    (
                        BGridViewRetr1,
                        gridBandSub,
                        usr
                    );

                    bList.SetValue(gridBandSub, i);
                }
            }

            gridBand2.Width = 450;
            gridBand2.OptionsBand.AllowMove = false;

            GridBand gridBand3 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            BandedGridColumn gita = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

            gridBand3.AppearanceHeader.Options.UseTextOptions = true;
            gridBand3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridBand3.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gridBand3.Caption = "기타";
            gridBand3.Columns.Add(gita);
            gridBand3.Width = 75;
            gridBand3.OptionsBand.AllowMove = false;

            ComnGridFunc.AddBand(BGridViewRetr1, gridBand3);

            gita.AppearanceCell.Options.UseTextOptions = true;
            gita.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gita.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gita.Caption = "GITA";
            gita.FieldName = "GITA";
            gita.OptionsColumn.AllowEdit = false;
            gita.Visible = true;

            ComnGridFunc.AddGridColumn
            (
                BGridViewRetr1,
                gridBand3,
                gita
            );

            ComnGridFunc.EndInitialize(BGridViewRetr1);

            ComnGridFunc.EndInitialize(BGridRetr1);

            //주간 실적 데이터
            strSql.Clear();
            strSql.AppendLine("   DECLARE @MCDANGA NUMERIC;");
            strSql.AppendLine("       SET ANSI_WARNINGS OFF ");
            strSql.AppendLine("       SET ARITHIGNORE ON    ");
            strSql.AppendLine("       SET ARITHABORT OFF;   ");
            strSql.AppendLine("    --경량야드 매출평균단가                                                                                                               ");
            strSql.AppendLine("    SELECT @MCDANGA = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))                                                ");
            strSql.AppendLine("      FROM INLIST A, JAJAE B                                                                                                              ");
            strSql.AppendLine("     WHERE A.J_SERIAL = B.J_SERIAL                                                                                                        ");
            strSql.AppendLine("       AND A.KERATYPE = '매출'                                                                                                            ");
            strSql.AppendLine("       AND A.J_DATE BETWEEN '" + sFrom + "' AND '" + sTo + "'                                                                                 ");
            strSql.AppendLine("       AND A.J_LOTNO <> '4'                                                                                                               ");
            strSql.AppendLine("       AND B.DAEGUBUN = '고철B'                                                                                                           ");
            strSql.AppendLine("       AND B.GUBUN1 <> '인센티브';                                                                                                        ");
            strSql.AppendLine("                                                                                                                                          ");
            strSql.AppendLine("    WITH TEMP1 AS(                                                                                                                        ");
            strSql.AppendLine("        --중량야드 고철A 실적                                                                                                             ");
            strSql.AppendLine("        SELECT B.CHRG_ID                                                                                                                  ");
            strSql.AppendLine("                , D.EMP_NM                                                                                                                ");
            strSql.AppendLine("                , SUM(A.DANJUNG) / 1000 AS DANJUNG                                                                                        ");
            strSql.AppendLine("            FROM INLIST A                                                                                                                 ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B                                                                                                     ");
            strSql.AppendLine("              ON A.J_ID1 = B.DEALER_CD                                                                                                    ");
            strSql.AppendLine("            LEFT JOIN JAJAE C                                                                                                             ");
            strSql.AppendLine("              ON A.J_SERIAL = C.J_SERIAL                                                                                                  ");
            strSql.AppendLine("            LEFT JOIN HR_EMP_BASIS D                                                                                                      ");
            strSql.AppendLine("              ON B.CHRG_ID = D.EMP_ID                                                                                                     ");
            strSql.AppendLine("           WHERE A.KERATYPE = '매입'                                                                                                      ");
            strSql.AppendLine("             AND A.J_DATE BETWEEN '" + sFrom + "' AND '" + sTo + "'                                                                       ");
            strSql.AppendLine("             AND A.J_LOTNO <> '4'                                                                                                         ");
            strSql.AppendLine("             AND C.DAEGUBUN = '고철A'                                                                                                     ");
            strSql.AppendLine("             AND C.GUBUN1 <> '인센티브'                                                                                                   ");
            strSql.AppendLine("           GROUP BY B.CHRG_ID, D.EMP_NM                                                                                                   ");
            strSql.AppendLine("    ), TEMP2 AS(                                                                                                                          ");
            strSql.AppendLine("        --경량야드 고철B 실적                                                                                                             ");
            strSql.AppendLine("        SELECT B.CHRG_ID                                                                                                                  ");
            strSql.AppendLine("             , D.EMP_NM                                                                                                                   ");
            strSql.AppendLine("             , SUM(A.DANJUNG) / 1000 AS DANJUNG                                                                                           ");
            strSql.AppendLine("          FROM INLIST A                                                                                                                   ");
            strSql.AppendLine("          LEFT JOIN ACC_DEALER_CD B                                                                                                       ");
            strSql.AppendLine("            ON A.J_ID1 = B.DEALER_CD                                                                                                      ");
            strSql.AppendLine("          LEFT JOIN JAJAE C                                                                                                               ");
            strSql.AppendLine("            ON A.J_SERIAL = C.J_SERIAL                                                                                                    ");
            strSql.AppendLine("          LEFT JOIN HR_EMP_BASIS D                                                                                                        ");
            strSql.AppendLine("            ON B.CHRG_ID = D.EMP_ID                                                                                                       ");
            strSql.AppendLine("         WHERE A.KERATYPE = '매입'                                                                                                        ");
            strSql.AppendLine("           AND A.J_DATE BETWEEN '" + sFrom + "' AND '" + sTo + "'                                                                             ");
            strSql.AppendLine("           AND A.J_LOTNO <> '4'                                                                                                           ");
            strSql.AppendLine("           AND C.DAEGUBUN = '고철B'                                                                                                       ");
            strSql.AppendLine("           AND C.GUBUN1 <> '인센티브'                                                                                                     ");
            strSql.AppendLine("         GROUP BY B.CHRG_ID, D.EMP_NM                                                                                                     ");
            strSql.AppendLine("    ), TEMP3 AS(                                                                                                                          ");
            strSql.AppendLine("        --경량야드 담당자별 매입단가/ 달성마진                                                                                            ");
            strSql.AppendLine("        SELECT B.CHRG_ID                                                                                                                  ");
            strSql.AppendLine("             , D.EMP_NM                                                                                                                   ");
            strSql.AppendLine("             , @MCDANGA - ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) AS I_MS                                                  ");
            strSql.AppendLine("          FROM INLIST A                                                                                                                   ");
            strSql.AppendLine("          LEFT JOIN ACC_DEALER_CD B                                                                                                       ");
            strSql.AppendLine("            ON A.J_ID1 = B.DEALER_CD                                                                                                      ");
            strSql.AppendLine("          LEFT JOIN JAJAE C                                                                                                               ");
            strSql.AppendLine("            ON A.J_SERIAL = C.J_SERIAL                                                                                                    ");
            strSql.AppendLine("          LEFT JOIN hr_emp_basis D                                                                                                        ");
            strSql.AppendLine("            ON B.CHRG_ID = D.EMP_ID                                                                                                       ");
            strSql.AppendLine("         WHERE A.KERATYPE = '매입'                                                                                                        ");
            strSql.AppendLine("           AND A.J_DATE BETWEEN '" + sFrom + "' AND '" + sTo + "'                                                                             ");
            strSql.AppendLine("           AND A.J_LOTNO <> '4'                                                                                                           ");
            strSql.AppendLine("           AND C.DAEGUBUN = '고철B'                                                                                                       ");
            strSql.AppendLine("         GROUP BY B.CHRG_ID, D.EMP_NM                                                                                                     ");
            strSql.AppendLine("    ),TEMP4 AS(                                                                                                                           ");
            strSql.AppendLine("         --폐압 집계 내역                                                                                                                 ");
            strSql.AppendLine("         SELECT B.CHRG_ID                                                                                                                 ");
            strSql.AppendLine("              , D.EMP_NM                                                                                                                  ");
            strSql.AppendLine("              , SUM(A.DANJUNG) / 1000 AS DANJUNG                                                                                          ");
            strSql.AppendLine("           FROM INLIST A, ACC_DEALER_CD B, JAJAE C, HR_EMP_BASIS D                                                                        ");
            strSql.AppendLine("          WHERE A.J_ID1 = B.DEALER_CD                                                                                                     ");
            strSql.AppendLine("            AND A.J_SERIAL = C.J_SERIAL                                                                                                   ");
            strSql.AppendLine("            AND A.KERATYPE = '매입'                                                                                                       ");
            strSql.AppendLine("            AND A.J_DATE BETWEEN '" + sFrom + "' AND '" + sTo + "'                                                                            ");
            strSql.AppendLine("            AND A.J_LOTNO <> '4'                                                                                                          ");
            strSql.AppendLine("            AND C.DAEGUBUN = '슈레더'                                                                                                     ");
            strSql.AppendLine("            AND C.GUBUN1 <> '인센티브'                                                                                                    ");
            strSql.AppendLine("            AND B.CHRG_ID = D.EMP_ID                                                                                                      ");
            strSql.AppendLine("          GROUP BY B.CHRG_ID, D.EMP_NM                                                                                                    ");
            strSql.AppendLine("    ), TEMP5 AS(                                                                                                                          ");
            strSql.AppendLine("        --직송스크랩 집계내역                                                                                                             ");
            strSql.AppendLine("        SELECT B.CHRG_ID                                                                                                                  ");
            strSql.AppendLine("             , D.EMP_NM                                                                                                                   ");
            strSql.AppendLine("             , SUM(A.DANJUNG) / 1000 AS DANJUNG                                                                                           ");
            strSql.AppendLine("          FROM INLIST A, ACC_DEALER_CD B, JAJAE C, HR_EMP_BASIS D                                                                         ");
            strSql.AppendLine("         WHERE A.J_ID1 = B.DEALER_CD                                                                                                      ");
            strSql.AppendLine("           AND A.J_SERIAL = C.J_SERIAL                                                                                                    ");
            strSql.AppendLine("           AND A.KERATYPE = '매입'                                                                                                        ");
            strSql.AppendLine("           AND A.J_DATE BETWEEN '" + sFrom + "' AND '" + sTo + "'                                                                             ");
            strSql.AppendLine("           AND A.J_LOTNO = '4'                                                                                                            ");
            strSql.AppendLine("           AND C.GUBUN1 <> '인센티브'                                                                                                     ");
            strSql.AppendLine("           AND B.CHRG_ID = D.EMP_ID                                                                                                       ");
            strSql.AppendLine("         GROUP BY B.CHRG_ID, D.EMP_NM                                                                                                     ");
            strSql.AppendLine("    ), TEMP6 AS( ");
            strSql.AppendLine("        SELECT A1.NAME                                                                                                                        ");
            strSql.AppendLine("             , ISNULL(A1.I_WEIGHT, 0) AS I_WEIGHT --중량야드                ");
            strSql.AppendLine("             , ISNULL(A1.I_LIGHT, 0) AS I_LIGHT --경량야드                  ");
            strSql.AppendLine("             , ISNULL(A1.I_MAJIN, 0) AS I_MAJIN --마진                      ");
            strSql.AppendLine("             , ISNULL(A1.I_CHAPI, 0) AS I_CHAPI --폐압모재                  ");
            strSql.AppendLine("             , ISNULL(A1.I_YK, 0) AS I_YK --직납                            ");
            strSql.AppendLine("             , ROUND(ISNULL(B1.DANJUNG, 0), 0) AS I_WS --중량야드실적           ");
            strSql.AppendLine("             , ROUND(ISNULL(B2.DANJUNG, 0), 0) AS I_LS --경량야드실적           ");
            strSql.AppendLine("             , ROUND(ISNULL(B3.I_MS, 0), 2) AS I_MS --달성마진                  ");
            strSql.AppendLine("             , ROUND(ISNULL(B4.DANJUNG, 0), 0) AS I_CS --폐압실적               ");
            strSql.AppendLine("             , ROUND(ISNULL(B5.DANJUNG, 0), 0) AS I_YS --직납실적               ");
            strSql.AppendLine("             , ROUND(B1.DANJUNG / A1.I_WEIGHT, 2) * 100 AS I_WP --중량야드달성율");
            strSql.AppendLine("             , ROUND(B2.DANJUNG / A1.I_LIGHT, 2) * 100 AS I_LP --경량야드달성율 ");
            strSql.AppendLine("             , ROUND(B3.I_MS / A1.I_MAJIN, 2) * 100 AS I_MP --마진달성율        ");
            strSql.AppendLine("             , ROUND(B4.DANJUNG / A1.I_CHAPI, 2) * 100 AS I_CP --폐압달성율     ");
            strSql.AppendLine("             , ROUND(B5.DANJUNG / A1.I_YK, 2) * 100 AS I_YP --직납달성율        ");
            strSql.AppendLine("          FROM SALEMAEIP A1                                                                                                                   ");
            strSql.AppendLine("          LEFT JOIN TEMP1 B1                                                                                                                  ");
            strSql.AppendLine("            ON A1.NAME = B1.EMP_NM                                                                                                            ");
            strSql.AppendLine("          LEFT JOIN TEMP2 B2                                                                                                                  ");
            strSql.AppendLine("            ON A1.NAME = B2.EMP_NM                                                                                                            ");
            strSql.AppendLine("          LEFT JOIN TEMP3 B3                                                                                                                  ");
            strSql.AppendLine("            ON A1.Name = B3.EMP_NM                                                                                                            ");
            strSql.AppendLine("          LEFT JOIN TEMP4 B4                                                                                                                  ");
            strSql.AppendLine("            ON A1.Name = B4.EMP_NM                                                                                                            ");
            strSql.AppendLine("          LEFT JOIN TEMP5 B5                                                                                                                  ");
            strSql.AppendLine("            ON A1.Name = B5.EMP_NM                                                                                                            ");
            strSql.AppendLine("         WHERE YYMM = '" + sYM + "'                                                                                                                ");
            strSql.AppendLine("           AND JUCHA = " + jucha + "                                                                                                                 ");
            strSql.AppendLine(" )                                           ");
            strSql.AppendLine("                                             ");
            strSql.AppendLine(" SELECT*                                     ");
            strSql.AppendLine("   FROM TEMP6                                ");
            strSql.AppendLine("  UNION ALL                                  ");
            strSql.AppendLine(" SELECT '합계'                               ");
            strSql.AppendLine("      , SUM(I_WEIGHT) AS I_WMT               ");
            strSql.AppendLine("      , SUM(I_LIGHT) AS I_LMT                ");
            strSql.AppendLine("      , SUM(I_MAJIN) AS I_MMT                ");
            strSql.AppendLine("      , SUM(I_CHAPI) AS I_CMT                ");
            strSql.AppendLine("      , SUM(I_YK) AS I_YMT                   ");
            strSql.AppendLine("      , SUM(I_WS) AS I_WST --중량야드실적    ");
            strSql.AppendLine("      , SUM(I_LS) AS I_LST --경량야드실적    ");
            strSql.AppendLine("      , SUM(I_MS) AS I_MST --달성마진        ");
            strSql.AppendLine("      , SUM(I_CS) AS I_CST --폐압실적        ");
            strSql.AppendLine("      , SUM(I_YS) AS I_YST --직납실적        ");
            strSql.AppendLine("      , ROUND(SUM(I_WS) / SUM(I_WEIGHT), 2) * 100 AS I_WPT --중량야드달성율");
            strSql.AppendLine("      , ROUND(SUM(I_LS) / SUM(I_LIGHT), 2) * 100 AS I_LPT --경량야드달성율 ");
            strSql.AppendLine("      , ROUND(SUM(I_MS) / SUM(I_MAJIN), 2) * 100 AS I_MPT --마진달성율     ");
            strSql.AppendLine("      , ROUND(SUM(I_CS) / SUM(I_CHAPI), 2) * 100 AS I_CPT --폐압달성율     ");
            strSql.AppendLine("      , ROUND(SUM(I_YS) / SUM(I_YK), 2) * 100 AS I_YPT --직납달성율        ");
            strSql.AppendLine("   FROM TEMP6                                ");
            strSql.AppendLine("  WHERE NAME != '기타'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count > 1)
            {
                DataTable temp = new DataTable();

                temp.Columns.Add("TITLE1");
                temp.Columns.Add("TITLE2");

                for (int i = 0; i < nameList.Count; i++)
                {
                    string name = nameList[i].ToString();

                    temp.Columns.Add(name);
                }

                temp.Columns.Add("GITA");

                //중량야드
                DataRow row1_1 = temp.NewRow();
                DataRow row1_2 = temp.NewRow();
                DataRow row1_3 = temp.NewRow();

                row1_1["TITLE1"] = "중량야드";
                row1_1["TITLE2"] = "중량목표";

                row1_2["TITLE1"] = "중량야드";
                row1_2["TITLE2"] = "야드실적";

                row1_3["TITLE1"] = "중량야드";
                row1_3["TITLE2"] = "달성율";

                //경량야드
                DataRow row2_1 = temp.NewRow();
                DataRow row2_2 = temp.NewRow();
                DataRow row2_3 = temp.NewRow();

                row2_1["TITLE1"] = "경량야드";
                row2_1["TITLE2"] = "경량목표";

                row2_2["TITLE1"] = "경량야드";
                row2_2["TITLE2"] = "야드실적";

                row2_3["TITLE1"] = "경량야드";
                row2_3["TITLE2"] = "달성율";

                //마진율야드
                DataRow row3_1 = temp.NewRow();
                DataRow row3_2 = temp.NewRow();
                DataRow row3_3 = temp.NewRow();

                row3_1["TITLE1"] = "마진율야드";
                row3_1["TITLE2"] = "마진목표";

                row3_2["TITLE1"] = "마진율야드";
                row3_2["TITLE2"] = "달성마진";

                row3_3["TITLE1"] = "마진율야드";
                row3_3["TITLE2"] = "달성율";

                //폐압모재
                DataRow row4_1 = temp.NewRow();
                DataRow row4_2 = temp.NewRow();
                DataRow row4_3 = temp.NewRow();

                row4_1["TITLE1"] = "폐압모재";
                row4_1["TITLE2"] = "폐압목표";

                row4_2["TITLE1"] = "폐압모재";
                row4_2["TITLE2"] = "폐압실적";

                row4_3["TITLE1"] = "폐압모재";
                row4_3["TITLE2"] = "달성율";

                //전체
                DataRow row5_1 = temp.NewRow();
                DataRow row5_2 = temp.NewRow();
                DataRow row5_3 = temp.NewRow();

                row5_1["TITLE1"] = "전체";
                row5_1["TITLE2"] = "전체목표";

                row5_2["TITLE1"] = "전체";
                row5_2["TITLE2"] = "전체실적";

                row5_3["TITLE1"] = "전체";
                row5_3["TITLE2"] = "달성율";

                //스크랩직납
                DataRow row6_1 = temp.NewRow();
                DataRow row6_2 = temp.NewRow();
                DataRow row6_3 = temp.NewRow();

                row6_1["TITLE1"] = "스크랩직납";
                row6_1["TITLE2"] = "직납목표";

                row6_2["TITLE1"] = "스크랩직납";
                row6_2["TITLE2"] = "직납실적";

                row6_3["TITLE1"] = "스크랩직납";
                row6_3["TITLE2"] = "달성율";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string name = dt.Rows[i]["NAME"]?.ToString();

                    string I_WEIGHT = dt.Rows[i]["I_WEIGHT"]?.ToString(); //중량야드목표
                    string I_WS = dt.Rows[i]["I_WS"]?.ToString(); //중량야드실적
                    string I_WP = dt.Rows[i]["I_WP"]?.ToString(); //중량야드달성율

                    string I_LIGHT = dt.Rows[i]["I_LIGHT"]?.ToString(); //경량야드목표
                    string I_LS = dt.Rows[i]["I_LS"]?.ToString(); //경량야드실적
                    string I_LP = dt.Rows[i]["I_LP"]?.ToString(); //경량야드달성율

                    string I_MAJIN = dt.Rows[i]["I_MAJIN"]?.ToString(); //마진목표
                    string I_MS = dt.Rows[i]["I_MS"]?.ToString(); //달성마진
                    string I_MP = dt.Rows[i]["I_MP"]?.ToString(); //마진달성율

                    string I_CHAPI = dt.Rows[i]["I_CHAPI"]?.ToString(); //폐압모재목표
                    string I_CS = dt.Rows[i]["I_CS"]?.ToString(); //폐압실적
                    string I_CP = dt.Rows[i]["I_CP"]?.ToString(); //폐압달성율

                    double dI_WEIGHT = 0;
                    double dI_LIGHT = 0;
                    double dI_MAJIN = 0;
                    double dI_CHAPI = 0;

                    double.TryParse(I_WEIGHT, out dI_WEIGHT);
                    double.TryParse(I_LIGHT, out dI_LIGHT);
                    double.TryParse(I_MAJIN, out dI_MAJIN);
                    double.TryParse(I_CHAPI, out dI_CHAPI);

                    double dI_WS = 0;
                    double dI_LS = 0;
                    double dI_MS = 0;
                    double dI_CS = 0;

                    double.TryParse(I_WS, out dI_WS);
                    double.TryParse(I_LS, out dI_LS);
                    double.TryParse(I_MS, out dI_MS);
                    double.TryParse(I_CS, out dI_CS);

                    double totMok = dI_WEIGHT + dI_LIGHT + dI_CHAPI; //전체목표
                    double totSil = dI_WS + dI_LS + dI_CS;//전체실적
                    double totPersent = totSil / totMok; //달성율

                    string I_YK = dt.Rows[i]["I_YK"]?.ToString(); //직납목표
                    string I_YS = dt.Rows[i]["I_YS"]?.ToString(); //직납실적
                    string I_YP = dt.Rows[i]["I_YP"]?.ToString(); //직납달성율


                    if (name.Equals("합계"))
                    {
                        string gitaTot = row5_2["GITA"]?.ToString();
                        double dGitaTot = 0;

                        double.TryParse(gitaTot, out dGitaTot);

                        double gitaSil = dGitaTot / totSil;

                        row5_3["GITA"] = gitaSil.ToString("P0");
                    }

                    if (name.Equals("기타"))
                    {
                        row1_1["GITA"] = chgFormat(I_WS, "double", "#,#");
                        row1_2["GITA"] = chgFormat(I_WS, "double", "#,#");
                        row1_3["GITA"] = chgFormat(I_WS, "double", "#,#");

                        row2_1["GITA"] = chgFormat(I_LS, "double", "#,#");
                        row2_2["GITA"] = chgFormat(I_LS, "double", "#,#");
                        row2_3["GITA"] = chgFormat(I_LS, "double", "#,#");

                        //row3_1["GITA"] = chgFormat(I_MS, "double", "#,#");
                        //row3_2["GITA"] = chgFormat(I_MS, "double", "#,#");
                        //row3_3["GITA"] = chgFormat(I_MS, "double", "#,#");

                        row4_1["GITA"] = chgFormat(I_CS, "double", "#,#");
                        row4_2["GITA"] = chgFormat(I_CS, "double", "#,#");
                        row4_3["GITA"] = chgFormat(I_CS, "double", "#,#");

                        row5_1["GITA"] = totSil.ToString("#,#");
                        row5_2["GITA"] = totSil.ToString("#,#");

                        row6_1["GITA"] = chgFormat(I_YS, "double", "#,#");
                        row6_2["GITA"] = chgFormat(I_YS, "double", "#,#");
                        row6_3["GITA"] = chgFormat(I_YS, "double", "#,#");
                    }
                    else
                    {
                        row1_1[name] = chgFormat(I_WEIGHT, "double", "#,#");
                        row1_2[name] = chgFormat(I_WS, "double", "#,#");
                        row1_3[name] = chgFormat(I_WP, "double", "n0")+"%";

                        row2_1[name] = chgFormat(I_LIGHT, "double", "#,#");
                        row2_2[name] = chgFormat(I_LS, "double", "#,#");
                        row2_3[name] = chgFormat(I_LP, "double", "n0")+"%";

                        row3_1[name] = chgFormat(I_MAJIN, "double", "#,#");
                        row3_2[name] = chgFormat(I_MS, "double", "#,#");
                        row3_3[name] = chgFormat(I_MP, "double", "n0") + "%";

                        row4_1[name] = chgFormat(I_CHAPI, "double", "#,#");
                        row4_2[name] = chgFormat(I_CS, "double", "#,#");
                        row4_3[name] = chgFormat(I_CP, "double", "n0") + "%";

                        row5_1[name] = totMok.ToString("#,#");
                        row5_2[name] = totSil.ToString("#,#");
                        row5_3[name] = totPersent.ToString("P0");

                        row6_1[name] = chgFormat(I_YK, "double", "#,#");
                        row6_2[name] = chgFormat(I_YS, "double", "#,#");
                        row6_3[name] = chgFormat(I_YP, "double", "n0") + "%";
                    }

                }
                temp.Rows.Add(row1_1);
                temp.Rows.Add(row1_2);
                temp.Rows.Add(row1_3);

                temp.Rows.Add(row2_1);
                temp.Rows.Add(row2_2);
                temp.Rows.Add(row2_3);

                temp.Rows.Add(row3_1);
                temp.Rows.Add(row3_2);
                temp.Rows.Add(row3_3);

                temp.Rows.Add(row4_1);
                temp.Rows.Add(row4_2);
                temp.Rows.Add(row4_3);

                temp.Rows.Add(row5_1);
                temp.Rows.Add(row5_2);
                temp.Rows.Add(row5_3);

                temp.Rows.Add(row6_1);
                temp.Rows.Add(row6_2);
                temp.Rows.Add(row6_3);

                BGridRetr1.DataSource = temp;
            }
        }

        //월간
        private void SetRetr7()
        {
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);
            string sYM = sTo.Substring(0, 7).Replace("-", "");

            StringBuilder strSql = new StringBuilder();

            //영업팀 목록
            strSql.Clear();
            strSql.AppendLine(" SELECT A1.NAME      ");
            strSql.AppendLine("   FROM SALEMAEIP A1 ");
            strSql.AppendLine(" WHERE YYMM = '" + sYM + "'");
            strSql.AppendLine("  GROUP BY A1.NAME   ");

            DataTable dtSawon = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            ArrayList nameList = new ArrayList();

            for (int i = 0; i < dtSawon.Rows.Count; i++)
            {
                string name = dtSawon.Rows[i]["NAME"].ToString();

                if (name.Equals("기타"))
                    continue;

                nameList.Add(name);
            }

            nameList.Add("합계");
            BGridViewRetr2.Bands.Clear();

            ComnGridFunc.BeginInitialize(BGridRetr2);

            BGridViewRetr2.Bands.Clear();

            ComnGridFunc.AddView(BGridRetr2, BGridViewRetr2, true);

            ComnGridFunc.BeginInitialize(BGridViewRetr2);

            GridBand gridBand1 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();

            BandedGridColumn title1 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            BandedGridColumn title2 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

            gridBand1.AppearanceHeader.Options.UseTextOptions = true;
            gridBand1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridBand1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gridBand1.Caption = sTo.Substring(5,2) + "월";
            gridBand1.Columns.Add(title1);
            gridBand1.Columns.Add(title2);
            gridBand1.Width = 150;
            gridBand1.OptionsBand.AllowMove = false;

            title1.AppearanceCell.Options.UseTextOptions = true;
            title1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            title1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            title1.Caption = "TITLE1";
            title1.FieldName = "TITLE1";
            title1.OptionsColumn.AllowEdit = false;
            title1.Visible = true;

            ComnGridFunc.AddBand(BGridViewRetr2, gridBand1);

            ComnGridFunc.AddGridColumn
            (
                BGridViewRetr2,
                gridBand1,
                title1
            );

            title2.AppearanceCell.Options.UseTextOptions = true;
            title2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            title2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            title2.Caption = "TITLE2";
            title2.FieldName = "TITLE2";
            title2.OptionsColumn.AllowEdit = false;
            title2.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            title2.Visible = true;

            ComnGridFunc.AddGridColumn
            (
                BGridViewRetr2,
                gridBand1,
                title2
            );

            GridBand gridBand2 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();

            gridBand2.AppearanceHeader.Options.UseTextOptions = true;
            gridBand2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridBand2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gridBand2.Caption = "영업부 - 주간 실적표";

            ComnGridFunc.AddBand(BGridViewRetr2, gridBand2);

            GridBand[] bList = new GridBand[dtSawon.Rows.Count];
            if(dtSawon.Rows.Count > 0)
            {
                for (int i = 0; i < nameList.Count; i++)
                {
                    string name = nameList[i].ToString();

                    GridBand gridBandSub = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
                    BandedGridColumn usr = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

                    gridBandSub.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandSub.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandSub.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandSub.Caption = name;
                    gridBandSub.Columns.Add(usr);
                    gridBandSub.Name = name + i;
                    gridBandSub.Width = 75;

                    usr.AppearanceCell.Options.UseTextOptions = true;
                    usr.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    usr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    usr.Caption = name;
                    usr.FieldName = name;
                    usr.Name = name + "_sub" + i;
                    usr.OptionsColumn.AllowEdit = false;
                    usr.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                    usr.Visible = true;

                    ComnGridFunc.AddChildGridBand(gridBand2, gridBandSub);

                    ComnGridFunc.AddGridColumn
                    (
                        BGridViewRetr2,
                        gridBandSub,
                        usr
                    );

                    bList.SetValue(gridBandSub, i);
                }
            }
            
            gridBand2.Width = 450;
            gridBand2.OptionsBand.AllowMove = false;

            GridBand gridBand3 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            BandedGridColumn gita = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

            gridBand3.AppearanceHeader.Options.UseTextOptions = true;
            gridBand3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridBand3.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gridBand3.Caption = "기타";
            gridBand3.Columns.Add(gita);
            gridBand3.Width = 75;
            gridBand3.OptionsBand.AllowMove = false;

            ComnGridFunc.AddBand(BGridViewRetr2, gridBand3);

            gita.AppearanceCell.Options.UseTextOptions = true;
            gita.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gita.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            gita.Caption = "GITA";
            gita.FieldName = "GITA";
            gita.OptionsColumn.AllowEdit = false;
            gita.Visible = true;

            ComnGridFunc.AddGridColumn
            (
                BGridViewRetr2,
                gridBand3,
                gita
            );

            ComnGridFunc.EndInitialize(BGridViewRetr2);

            ComnGridFunc.EndInitialize(BGridRetr2);

            //월간 실적 데이터
            strSql.Clear();
            strSql.AppendLine(" DECLARE @MCDANGA NUMERIC;                                                               ");
            strSql.AppendLine("                                                                                         ");
            strSql.AppendLine(" --경량야드 매출평균단가                                                                 ");
            strSql.AppendLine(" SELECT @MCDANGA = (SUM(A.KONGKEP) - SUM(A.CKONGKEP)) / (SUM(A.DANJUNG) + SUM(A.HALIN))  ");
            strSql.AppendLine("   FROM INLIST A, JAJAE B                                                                ");
            strSql.AppendLine("  WHERE A.J_SERIAL = B.J_SERIAL                                                          ");
            strSql.AppendLine("    AND A.KERATYPE = '매출'                                                              ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN '" + sTo.Substring(0, 7) + "-01' AND '" + sTo + "'                                   ");
            strSql.AppendLine("    AND A.J_LOTNO <> '4'                                                                 ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철B'                                                             ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브';                                                          ");
            strSql.AppendLine("                                                                                         ");
            strSql.AppendLine(" SET ANSI_WARNINGS OFF                                                                   ");
            strSql.AppendLine(" SET ARITHIGNORE ON                                                                      ");
            strSql.AppendLine(" SET ARITHABORT OFF;                                                                     ");
            strSql.AppendLine("                                                                                         ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                                          ");
            strSql.AppendLine("     --중량야드 고철A 실적                                                               ");
            strSql.AppendLine("     SELECT B.CHRG_ID                                                                    ");
            strSql.AppendLine("          , D.EMP_NM                                                                     ");
            strSql.AppendLine("          , SUM(A.DANJUNG) / 1000 AS DANJUNG                                             ");
            strSql.AppendLine("       FROM INLIST A                                                                     ");
            strSql.AppendLine("       LEFT JOIN ACC_DEALER_CD B                                                         ");
            strSql.AppendLine("         ON A.J_ID1 = B.DEALER_CD                                                        ");
            strSql.AppendLine("       LEFT JOIN JAJAE C                                                                 ");
            strSql.AppendLine("         ON A.J_SERIAL = C.J_SERIAL                                                      ");
            strSql.AppendLine("       LEFT JOIN HR_EMP_BASIS D                                                          ");
            strSql.AppendLine("         ON B.CHRG_ID = D.EMP_ID                                                         ");
            strSql.AppendLine("      WHERE A.KERATYPE = '매입'                                                          ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '" + sTo.Substring(0, 7) + "-01' AND '" + sTo + "'                               ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                             ");
            strSql.AppendLine("        AND C.DAEGUBUN = '고철A'                                                         ");
            strSql.AppendLine("        AND C.GUBUN1 <> '인센티브'                                                       ");
            strSql.AppendLine("      GROUP BY B.CHRG_ID, D.EMP_NM                                                       ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                            ");
            strSql.AppendLine("     --경량야드 고철B 실적                                                               ");
            strSql.AppendLine("     SELECT B.CHRG_ID                                                                    ");
            strSql.AppendLine("          , D.EMP_NM                                                                  ");
            strSql.AppendLine("          , SUM(A.DANJUNG) / 1000 AS DANJUNG                                          ");
            strSql.AppendLine("      FROM INLIST A                                                                   ");
            strSql.AppendLine("      LEFT JOIN ACC_DEALER_CD B                                                       ");
            strSql.AppendLine("        ON A.J_ID1 = B.DEALER_CD                                                      ");
            strSql.AppendLine("      LEFT JOIN JAJAE C                                                               ");
            strSql.AppendLine("        ON A.J_SERIAL = C.J_SERIAL                                                    ");
            strSql.AppendLine("      LEFT JOIN HR_EMP_BASIS D                                                        ");
            strSql.AppendLine("        ON B.CHRG_ID = D.EMP_ID                                                       ");
            strSql.AppendLine("     WHERE A.KERATYPE = '매입'                                                        ");
            strSql.AppendLine("       AND A.J_DATE BETWEEN '" + sTo.Substring(0, 7) + "-01' AND '" + sTo + "'                             ");
            strSql.AppendLine("       AND A.J_LOTNO <> '4'                                                           ");
            strSql.AppendLine("       AND C.DAEGUBUN = '고철B'                                                       ");
            strSql.AppendLine("       AND C.GUBUN1 <> '인센티브'                                                     ");
            strSql.AppendLine("     GROUP BY B.CHRG_ID, D.EMP_NM                                                     ");
            strSql.AppendLine(" ), TEMP3 AS(                                                                            ");
            strSql.AppendLine("     --경량야드 담당자별 매입단가/ 달성마진                                              ");
            strSql.AppendLine("     SELECT B.CHRG_ID                                                                    ");
            strSql.AppendLine("          , D.EMP_NM                                                                     ");
            strSql.AppendLine("          , @MCDANGA - ((SUM(A.IKONGKEP) + SUM(A.CKONGKEP)) / SUM(A.DANJUNG)) AS I_MS    ");
            strSql.AppendLine("       FROM INLIST A                                                                     ");
            strSql.AppendLine("       LEFT JOIN ACC_DEALER_CD B                                                         ");
            strSql.AppendLine("         ON A.J_ID1 = B.DEALER_CD                                                        ");
            strSql.AppendLine("       LEFT JOIN JAJAE C                                                                 ");
            strSql.AppendLine("         ON A.J_SERIAL = C.J_SERIAL                                                      ");
            strSql.AppendLine("       LEFT JOIN hr_emp_basis D                                                          ");
            strSql.AppendLine("         ON B.CHRG_ID = D.EMP_ID                                                         ");
            strSql.AppendLine("      WHERE A.KERATYPE = '매입'                                                          ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '" + sTo.Substring(0, 7) + "-01' AND '" + sTo + "'                               ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                             ");
            strSql.AppendLine("        AND C.DAEGUBUN = '고철B'                                                         ");
            strSql.AppendLine("      GROUP BY B.CHRG_ID, D.EMP_NM                                                       ");
            strSql.AppendLine(" ),TEMP4 AS(                                                                             ");
            strSql.AppendLine("      --폐압 집계 내역                                                                   ");
            strSql.AppendLine("      SELECT B.CHRG_ID                                                                   ");
            strSql.AppendLine("           , D.EMP_NM                                                                    ");
            strSql.AppendLine("           , SUM(A.DANJUNG) / 1000 AS DANJUNG                                            ");
            strSql.AppendLine("        FROM INLIST A, ACC_DEALER_CD B, JAJAE C, HR_EMP_BASIS D                          ");
            strSql.AppendLine("       WHERE A.J_ID1 = B.DEALER_CD                                                       ");
            strSql.AppendLine("         AND A.J_SERIAL = C.J_SERIAL                                                     ");
            strSql.AppendLine("         AND A.KERATYPE = '매입'                                                         ");
            strSql.AppendLine("         AND A.J_DATE BETWEEN '" + sTo.Substring(0, 7) + "-01' AND '" + sTo + "'                              ");
            strSql.AppendLine("         AND A.J_LOTNO <> '4'                                                            ");
            strSql.AppendLine("         AND C.DAEGUBUN = '슈레더'                                                       ");
            strSql.AppendLine("         AND C.GUBUN1 <> '인센티브'                                                      ");
            strSql.AppendLine("         AND B.CHRG_ID = D.EMP_ID                                                        ");
            strSql.AppendLine("       GROUP BY B.CHRG_ID, D.EMP_NM                                                      ");
            strSql.AppendLine(" ), TEMP5 AS(                                                                            ");
            strSql.AppendLine("     --직송스크랩 집계내역                                                               ");
            strSql.AppendLine("     SELECT B.CHRG_ID                                                                    ");
            strSql.AppendLine("          , D.EMP_NM                                                                     ");
            strSql.AppendLine("          , SUM(A.DANJUNG) / 1000 AS DANJUNG                                             ");
            strSql.AppendLine("       FROM INLIST A, ACC_DEALER_CD B, JAJAE C, HR_EMP_BASIS D                           ");
            strSql.AppendLine("      WHERE A.J_ID1 = B.DEALER_CD                                                        ");
            strSql.AppendLine("        AND A.J_SERIAL = C.J_SERIAL                                                      ");
            strSql.AppendLine("        AND A.KERATYPE = '매입'                                                          ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '" + sTo.Substring(0, 7) + "-01' AND '" + sTo + "'                               ");
            strSql.AppendLine("        AND A.J_LOTNO = '4'                                                              ");
            strSql.AppendLine("        AND C.GUBUN1 <> '인센티브'                                                       ");
            strSql.AppendLine("        AND B.CHRG_ID = D.EMP_ID                                                         ");
            strSql.AppendLine("      GROUP BY B.CHRG_ID, D.EMP_NM                                                       ");
            strSql.AppendLine(" ), TEMP6 AS(                                                                            ");
            strSql.AppendLine("     SELECT A1.NAME                                                                      ");
            strSql.AppendLine("          , ISNULL(SUM(A1.I_WEIGHT),0) AS I_WEIGHT --중량야드                            ");
            strSql.AppendLine("          , ISNULL(SUM(A1.I_LIGHT), 0) AS I_LIGHT --경량야드                             ");
            strSql.AppendLine("          , ISNULL(SUM(A1.I_MAJIN), 0) AS I_MAJIN --마진                                 ");
            strSql.AppendLine("          , ISNULL(SUM(A1.I_CHAPI), 0) AS I_CHAPI --폐압모재                             ");
            strSql.AppendLine("          , ISNULL(SUM(A1.I_YK), 0) AS I_YK --직납                                       ");
            strSql.AppendLine("       FROM SALEMAEIP A1                                                                 ");
            strSql.AppendLine("      WHERE YYMM = '" + sYM + "'                                                              ");
            strSql.AppendLine("      GROUP BY A1.NAME                                                                   ");
            strSql.AppendLine(" ), TEMP7 AS(                                                                            ");
            strSql.AppendLine("     SELECT A1.NAME                                                                      ");
            strSql.AppendLine("          , ISNULL(A1.I_WEIGHT,0) AS I_WEIGHT --중량야드                                 ");
            strSql.AppendLine("          , ISNULL(A1.I_LIGHT, 0) AS I_LIGHT --경량야드                                  ");
            strSql.AppendLine("          , ISNULL(A1.I_MAJIN, 0) AS I_MAJIN --마진                                      ");
            strSql.AppendLine("          , ISNULL(A1.I_CHAPI, 0) AS I_CHAPI --폐압모재                                  ");
            strSql.AppendLine("          , ISNULL(A1.I_YK, 0) AS I_YK --직납                                            ");
            strSql.AppendLine("          , ROUND(ISNULL(B1.DANJUNG, 0), 0) AS I_WS --중량야드실적           ");
            strSql.AppendLine("          , ROUND(ISNULL(B2.DANJUNG, 0), 0) AS I_LS --경량야드실적           ");
            strSql.AppendLine("          , ROUND(ISNULL(B3.I_MS, 0), 2) AS I_MS --달성마진                  ");
            strSql.AppendLine("          , ROUND(ISNULL(B4.DANJUNG, 0), 0) AS I_CS --폐압실적               ");
            strSql.AppendLine("          , ROUND(ISNULL(B5.DANJUNG, 0), 0) AS I_YS --직납실적               ");
            strSql.AppendLine("          , ROUND(B1.DANJUNG / A1.I_WEIGHT, 2) * 100 AS I_WP --중량야드달성율");
            strSql.AppendLine("          , ROUND(B2.DANJUNG / A1.I_LIGHT, 2) * 100 AS I_LP --경량야드달성율 ");
            strSql.AppendLine("          , ROUND(B3.I_MS / A1.I_MAJIN, 2) * 100 AS I_MP --마진달성율        ");
            strSql.AppendLine("          , ROUND(B4.DANJUNG / A1.I_CHAPI, 2) * 100 AS I_CP --폐압달성율     ");
            strSql.AppendLine("          , ROUND(B5.DANJUNG / A1.I_YK, 2) * 100 AS I_YP --직납달성율        ");
            strSql.AppendLine("       FROM TEMP6 A1                                                                     ");
            strSql.AppendLine("       LEFT JOIN TEMP1 B1                                                                ");
            strSql.AppendLine("         ON A1.NAME = B1.EMP_NM                                                          ");
            strSql.AppendLine("       LEFT JOIN TEMP2 B2                                                                ");
            strSql.AppendLine("         ON A1.NAME = B2.EMP_NM                                                          ");
            strSql.AppendLine("       LEFT JOIN TEMP3 B3                                                                ");
            strSql.AppendLine("         ON A1.Name = B3.EMP_NM                                                          ");
            strSql.AppendLine("       LEFT JOIN TEMP4 B4                                                                ");
            strSql.AppendLine("         ON A1.Name = B4.EMP_NM                                                          ");
            strSql.AppendLine("       LEFT JOIN TEMP5 B5                                                                ");
            strSql.AppendLine("         ON A1.Name = B5.EMP_NM                                                          ");
            strSql.AppendLine(" )                                                                                       ");
            strSql.AppendLine("                                                                                         ");
            strSql.AppendLine(" SELECT *                                                                                ");
            strSql.AppendLine("   FROM TEMP7                                                                            ");
            strSql.AppendLine("  UNION ALL                                                                              ");
            strSql.AppendLine(" SELECT '합계'                                                                           ");
            strSql.AppendLine("      , SUM(I_WEIGHT) AS I_WMT                                                           ");
            strSql.AppendLine("      , SUM(I_LIGHT) AS I_LMT                                                            ");
            strSql.AppendLine("      , SUM(I_MAJIN) AS I_MMT                                                            ");
            strSql.AppendLine("      , SUM(I_CHAPI) AS I_CMT                                                            ");
            strSql.AppendLine("      , SUM(I_YK) AS I_YMT                                                               ");
            strSql.AppendLine("      , SUM(I_WS) AS I_WST --중량야드실적                                                ");
            strSql.AppendLine("      , SUM(I_LS) AS I_LST --경량야드실적                                                ");
            strSql.AppendLine("      , SUM(I_MS) AS I_MST --달성마진                                                    ");
            strSql.AppendLine("      , SUM(I_CS) AS I_CST --폐압실적                                                    ");
            strSql.AppendLine("      , SUM(I_YS) AS I_YST --직납실적                                                    ");
            strSql.AppendLine("      , ROUND(SUM(I_WS) / SUM(I_WEIGHT), 2) * 100 AS I_WPT --중량야드달성율");
            strSql.AppendLine("      , ROUND(SUM(I_LS) / SUM(I_LIGHT), 2) * 100 AS I_LPT --경량야드달성율 ");
            strSql.AppendLine("      , ROUND(SUM(I_MS) / SUM(I_MAJIN), 2) * 100 AS I_MPT --마진달성율     ");
            strSql.AppendLine("      , ROUND(SUM(I_CS) / SUM(I_CHAPI), 2) * 100 AS I_CPT --폐압달성율     ");
            strSql.AppendLine("      , ROUND(SUM(I_YS) / SUM(I_YK), 2) * 100 AS I_YPT --직납달성율        ");
            strSql.AppendLine("   FROM TEMP7                                                                            ");
            strSql.AppendLine("  WHERE NAME != '기타'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count > 1)
            {
                DataTable temp = new DataTable();

                temp.Columns.Add("TITLE1");
                temp.Columns.Add("TITLE2");

                for (int i = 0; i < nameList.Count; i++)
                {
                    string name = nameList[i].ToString();

                    temp.Columns.Add(name);
                }

                temp.Columns.Add("GITA");

                //중량야드
                DataRow row1_1 = temp.NewRow();
                DataRow row1_2 = temp.NewRow();
                DataRow row1_3 = temp.NewRow();

                row1_1["TITLE1"] = "중량야드";
                row1_1["TITLE2"] = "중량목표";

                row1_2["TITLE1"] = "중량야드";
                row1_2["TITLE2"] = "야드실적";

                row1_3["TITLE1"] = "중량야드";
                row1_3["TITLE2"] = "달성율";

                //경량야드
                DataRow row2_1 = temp.NewRow();
                DataRow row2_2 = temp.NewRow();
                DataRow row2_3 = temp.NewRow();

                row2_1["TITLE1"] = "경량야드";
                row2_1["TITLE2"] = "경량목표";

                row2_2["TITLE1"] = "경량야드";
                row2_2["TITLE2"] = "야드실적";

                row2_3["TITLE1"] = "경량야드";
                row2_3["TITLE2"] = "달성율";

                //마진율야드
                DataRow row3_1 = temp.NewRow();
                DataRow row3_2 = temp.NewRow();
                DataRow row3_3 = temp.NewRow();

                row3_1["TITLE1"] = "마진율야드";
                row3_1["TITLE2"] = "마진목표";

                row3_2["TITLE1"] = "마진율야드";
                row3_2["TITLE2"] = "달성마진";

                row3_3["TITLE1"] = "마진율야드";
                row3_3["TITLE2"] = "달성율";

                //폐압모재
                DataRow row4_1 = temp.NewRow();
                DataRow row4_2 = temp.NewRow();
                DataRow row4_3 = temp.NewRow();

                row4_1["TITLE1"] = "폐압모재";
                row4_1["TITLE2"] = "폐압목표";

                row4_2["TITLE1"] = "폐압모재";
                row4_2["TITLE2"] = "폐압실적";

                row4_3["TITLE1"] = "폐압모재";
                row4_3["TITLE2"] = "달성율";

                //전체
                DataRow row5_1 = temp.NewRow();
                DataRow row5_2 = temp.NewRow();
                DataRow row5_3 = temp.NewRow();

                row5_1["TITLE1"] = "전체";
                row5_1["TITLE2"] = "전체목표";

                row5_2["TITLE1"] = "전체";
                row5_2["TITLE2"] = "전체실적";

                row5_3["TITLE1"] = "전체";
                row5_3["TITLE2"] = "달성율";

                //스크랩직납
                DataRow row6_1 = temp.NewRow();
                DataRow row6_2 = temp.NewRow();
                DataRow row6_3 = temp.NewRow();

                row6_1["TITLE1"] = "스크랩직납";
                row6_1["TITLE2"] = "직납목표";

                row6_2["TITLE1"] = "스크랩직납";
                row6_2["TITLE2"] = "직납실적";

                row6_3["TITLE1"] = "스크랩직납";
                row6_3["TITLE2"] = "달성율";


                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string name = dt.Rows[i]["NAME"]?.ToString();

                    string I_WEIGHT = dt.Rows[i]["I_WEIGHT"]?.ToString(); //중량야드목표
                    string I_WS = dt.Rows[i]["I_WS"]?.ToString(); //중량야드실적
                    string I_WP = dt.Rows[i]["I_WP"]?.ToString(); //중량야드달성율

                    string I_LIGHT = dt.Rows[i]["I_LIGHT"]?.ToString(); //경량야드목표
                    string I_LS = dt.Rows[i]["I_LS"]?.ToString(); //경량야드실적
                    string I_LP = dt.Rows[i]["I_LP"]?.ToString(); //경량야드달성율

                    string I_MAJIN = dt.Rows[i]["I_MAJIN"]?.ToString(); //마진목표
                    string I_MS = dt.Rows[i]["I_MS"]?.ToString(); //달성마진
                    string I_MP = dt.Rows[i]["I_MP"]?.ToString(); //마진달성율

                    string I_CHAPI = dt.Rows[i]["I_CHAPI"]?.ToString(); //폐압모재목표
                    string I_CS = dt.Rows[i]["I_CS"]?.ToString(); //폐압실적
                    string I_CP = dt.Rows[i]["I_CP"]?.ToString(); //폐압달성율

                    double dI_WEIGHT = 0;
                    double dI_LIGHT = 0;
                    double dI_MAJIN = 0;
                    double dI_CHAPI = 0;

                    double.TryParse(I_WEIGHT, out dI_WEIGHT);
                    double.TryParse(I_LIGHT, out dI_LIGHT);
                    double.TryParse(I_MAJIN, out dI_MAJIN);
                    double.TryParse(I_CHAPI, out dI_CHAPI);

                    double dI_WS = 0;
                    double dI_LS = 0;
                    double dI_MS = 0;
                    double dI_CS = 0;

                    double.TryParse(I_WS, out dI_WS);
                    double.TryParse(I_LS, out dI_LS);
                    double.TryParse(I_MS, out dI_MS);
                    double.TryParse(I_CS, out dI_CS);

                    double totMok = dI_WEIGHT + dI_LIGHT + dI_CHAPI; //전체목표
                    double totSil = dI_WS + dI_LS + dI_CS;//전체실적
                    double totPersent = totSil / totMok; //달성율

                    string I_YK = dt.Rows[i]["I_YK"]?.ToString(); //직납목표
                    string I_YS = dt.Rows[i]["I_YS"]?.ToString(); //직납실적
                    string I_YP = dt.Rows[i]["I_YP"]?.ToString(); //직납달성율

                    if (name.Equals("합계"))
                    {
                        string gitaTot = row5_2["GITA"]?.ToString();
                        double dGitaTot = 0;

                        double.TryParse(gitaTot, out dGitaTot);

                        row5_3["GITA"] = ((totSil + dGitaTot) / totMok).ToString("P0");
                        row5_1["GITA"] = (dGitaTot+ totSil).ToString("n0");
                        row5_2["GITA"] = (dGitaTot + totSil).ToString("n0");
                    }

                    if (name.Equals("기타"))
                    {
                        row1_1["GITA"] = chgFormat(I_WS, "double", "#,#");
                        row1_2["GITA"] = chgFormat(I_WS, "double", "#,#");
                        row1_3["GITA"] = chgFormat(I_WS, "double", "#,#");

                        row2_1["GITA"] = chgFormat(I_LS, "double", "#,#");
                        row2_2["GITA"] = chgFormat(I_LS, "double", "#,#");
                        row2_3["GITA"] = chgFormat(I_LS, "double", "#,#");

                        //row3_1["GITA"] = chgFormat(I_MS, "double", "#,#");
                        //row3_2["GITA"] = chgFormat(I_MS, "double", "#,#");
                        //row3_3["GITA"] = chgFormat(I_MS, "double", "#,#");

                        row4_1["GITA"] = chgFormat(I_CS, "double", "#,#");
                        row4_2["GITA"] = chgFormat(I_CS, "double", "#,#");
                        row4_3["GITA"] = chgFormat(I_CS, "double", "#,#");

                        row5_1["GITA"] = totSil.ToString("#,#");
                        row5_2["GITA"] = totSil.ToString("#,#");
                        row5_3["GITA"] = "0%";

                        row6_1["GITA"] = chgFormat(I_YS, "double", "#,#");
                        row6_2["GITA"] = chgFormat(I_YS, "double", "#,#");
                        row6_3["GITA"] = chgFormat(I_YS, "double", "#,#");
                    }
                    else
                    {
                        row1_1[name] = chgFormat(I_WEIGHT, "double", "#,#");
                        row1_2[name] = chgFormat(I_WS, "double", "#,#");
                        row1_3[name] = chgFormat(I_WP, "double", "n0") + "%";

                        row2_1[name] = chgFormat(I_LIGHT, "double", "#,#");
                        row2_2[name] = chgFormat(I_LS, "double", "#,#");
                        row2_3[name] = chgFormat(I_LP, "double", "n0") + "%";

                        row3_1[name] = chgFormat(I_MAJIN, "double", "#,#");
                        row3_2[name] = chgFormat(I_MS, "double", "#,#");
                        row3_3[name] = chgFormat(I_MP, "double", "n0") + "%";

                        row4_1[name] = chgFormat(I_CHAPI, "double", "#,#");
                        row4_2[name] = chgFormat(I_CS, "double", "#,#");
                        row4_3[name] = chgFormat(I_CP, "double", "n0") + "%";

                        row5_1[name] = totMok.ToString("#,#");
                        row5_2[name] = totSil.ToString("#,#");
                        row5_3[name] = totPersent.ToString("P0");

                        row6_1[name] = chgFormat(I_YK, "double", "#,#");
                        row6_2[name] = chgFormat(I_YS, "double", "#,#");
                        row6_3[name] = chgFormat(I_YP, "double", "n0") + "%";
                    }

                }
                temp.Rows.Add(row1_1);
                temp.Rows.Add(row1_2);
                temp.Rows.Add(row1_3);

                temp.Rows.Add(row2_1);
                temp.Rows.Add(row2_2);
                temp.Rows.Add(row2_3);

                temp.Rows.Add(row3_1);
                temp.Rows.Add(row3_2);
                temp.Rows.Add(row3_3);

                temp.Rows.Add(row4_1);
                temp.Rows.Add(row4_2);
                temp.Rows.Add(row4_3);

                temp.Rows.Add(row5_1);
                temp.Rows.Add(row5_2);
                temp.Rows.Add(row5_3);

                temp.Rows.Add(row6_1);
                temp.Rows.Add(row6_2);
                temp.Rows.Add(row6_3);

                BGridRetr2.DataSource = temp;
            }
        }

        //매출실적
        private void GetRetr8()
        {
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);
            string sFrom = sTo.Substring(0, 7)+"-01";
            string sYM = sTo.Substring(0, 7).Replace("-", "");

            GridColRetr8Title.Caption = sTo.Substring(5, 2)+"월 매출목표";

            DataTable dt = new DataTable();

            dt.Columns.Add("TITLE");
            dt.Columns.Add("W1");
            dt.Columns.Add("W2");
            dt.Columns.Add("W3");
            dt.Columns.Add("W4");
            dt.Columns.Add("W5");
            dt.Columns.Add("W6");
            dt.Columns.Add("TOT");

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" WITH TEMP1 AS(                                        ");
            strSql.AppendLine("     SELECT JUCHA                                      ");
            strSql.AppendLine("          , O_GS + O_WEIGHT + O_SD AS MCHMOK--매출목표 ");
            strSql.AppendLine("       FROM SALEMAECHUL                                ");
            strSql.AppendLine("      WHERE YYMM = '" + sYM + "'                       ");
            strSql.AppendLine(" )                                                     ");
            strSql.AppendLine(" SELECT*                                               ");
            strSql.AppendLine("   FROM TEMP1                                          ");
            strSql.AppendLine("  UNION ALL                                            ");
            strSql.AppendLine(" SELECT 111                                            ");
            strSql.AppendLine("      , SUM(MCHMOK)                                    ");
            strSql.AppendLine("   FROM TEMP1                                          ");

            DataTable dtMok = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //매출목표
            DataRow row1 = dt.NewRow();

            row1["TITLE"] = "매출 목표";

            if (dtMok != null && dtMok.Rows.Count > 0)
            {
                for (int i=0;i< dtMok.Rows.Count; i++)
                {
                    string jucha = dtMok.Rows[i]["JUCHA"]?.ToString();
                    string sWeight = dtMok.Rows[i]["MCHMOK"]?.ToString();

                    if (jucha.Equals("1"))
                    {
                        row1["W1"] = chgFormat(sWeight, "double", "n0");
                    }
                    else if (jucha.Equals("2"))
                    {
                        row1["W2"] = chgFormat(sWeight, "double", "n0");
                    }
                    else if (jucha.Equals("3"))
                    {
                        row1["W3"] = chgFormat(sWeight, "double", "n0");
                    }
                    else if (jucha.Equals("4"))
                    {
                        row1["W4"] = chgFormat(sWeight, "double", "n0");
                    }
                    else if (jucha.Equals("5"))
                    {
                        row1["W5"] = chgFormat(sWeight, "double", "n0");
                    }
                    else if (jucha.Equals("6"))
                    {
                        row1["W6"] = chgFormat(sWeight, "double", "n0");
                    }
                    else if (jucha.Equals("111"))
                    {
                        row1["TOT"] = chgFormat(sWeight, "double", "n0");
                    }
                }
            }
            dt.Rows.Add(row1);

            //매출실적(토요일~금요일)
            //일:0, 월:1 ~
            DateTime fstDt = DateTime.Parse(sFrom);
            DateTime lstDt = fstDt.AddMonths(1).AddDays(-1);

            DateTime sDate = fstDt;
            DateTime eDate;

            if (fstDt.DayOfWeek == DayOfWeek.Friday)
            {
                eDate = fstDt;
            }
            else
            {
                eDate = GetFriDay(sDate);
            }

            DataRow row2 = dt.NewRow();
            row2["TITLE"] = "매출 실적";

            bool chk = false;
            double totsil = 0;

            while (true)
            {
                strSql.Clear();
                strSql.AppendLine(" SET ANSI_WARNINGS OFF");
                strSql.AppendLine(" SET ARITHIGNORE ON   ");
                strSql.AppendLine(" SET ARITHABORT OFF;  ");
                strSql.AppendLine(" SELECT ROUND(SUM(DANJUNG)/ 1000,0) AS MCHSIL       ");
                strSql.AppendLine("   FROM INLIST                                      ");
                strSql.AppendLine("  WHERE KERATYPE = '매출'                           ");
                strSql.AppendLine("    AND J_DATE BETWEEN '"+ sDate.ToString("yyyy-MM-dd") + "' AND '"+ eDate.ToString("yyyy-MM-dd") + "'");
                strSql.AppendLine("    AND J_LOTNO<> '4'                               ");

                DataTable dtJu = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                string sWeight = dtJu.Rows[0]["MCHSIL"]?.ToString();
                double.TryParse(sWeight, out double dWeight);

                totsil = totsil + dWeight;

                int jucha = ComnEtcFunc.GetWeekOfMonth(eDate, null);

                if(jucha == 1)
                {
                    row2["W1"] = chgFormat(sWeight, "double", "n0");
                }
                else if(jucha == 2)
                {
                    row2["W2"] = chgFormat(sWeight, "double", "n0");
                }
                else if (jucha == 3)
                {
                    row2["W3"] = chgFormat(sWeight, "double", "n0");
                }
                else if (jucha == 4)
                {
                    row2["W4"] = chgFormat(sWeight, "double", "n0");
                }
                else if (jucha == 5)
                {
                    row2["W5"] = chgFormat(sWeight, "double", "n0");
                }
                else if (jucha == 6)
                {
                    row2["W6"] = chgFormat(sWeight, "double", "n0");
                }

                sDate = eDate.AddDays(1);
                eDate = sDate.AddDays(6);

                if (chk)
                {
                    break;
                }

                if (eDate > lstDt)
                {
                    eDate = lstDt;
                    chk = true; 
                }
            }

            row2["TOT"] = totsil.ToString("n0");
            dt.Rows.Add(row2);

            DataRow row3 = dt.NewRow();
            row3["TITLE"] = "달성율";

            if(dt.Rows.Count > 1)
            {
                for(int i = 0; i < 6; i++)
                {
                    double dMok = 0;
                    double dSil = 0;

                    string sMok = dt.Rows[0]["W" + (i + 1)]?.ToString();
                    string sSil = dt.Rows[1]["W" + (i + 1)]?.ToString();

                    double.TryParse(sMok, out dMok);
                    double.TryParse(sSil, out dSil);

                    if(dMok == 0 || dSil == 0)
                    {
                        row3["W" + (i + 1)] = "0%";
                    }
                    else
                    {
                        row3["W" + (i + 1)] = chgFormat((dSil / dMok).ToString(), "double", "P0");
                    }
                }

                string sTotMok = dt.Rows[0]["TOT"]?.ToString();
                string sTotSil = dt.Rows[1]["TOT"]?.ToString();

                double dTotMok = 0;
                double dTotSil = 0;

                double.TryParse(sTotMok, out dTotMok);
                double.TryParse(sTotSil, out dTotSil);

                if(dTotMok == 0 || dTotSil == 0)
                {
                    row3["TOT"] = "0%";
                }
                else
                {
                    row3["TOT"] = chgFormat((dTotSil / dTotMok).ToString(), "double", "P0");
                }
            }

            dt.Rows.Add(row3);

            GridRetr8.DataSource = dt;
        }

        private DateTime GetFriDay(DateTime sDate)
        {
            DateTime result;

            int fDay = Convert.ToInt32(sDate.DayOfWeek);
            int goalDay = Convert.ToInt32(DayOfWeek.Friday);

            DateTime fridayDate = sDate.AddDays(goalDay - fDay);

            if (sDate > fridayDate)
            {
                DateTime tempDt = sDate.AddDays(6);

                int tDay = Convert.ToInt32(tempDt.DayOfWeek);

                DateTime fridayDate2 = tempDt.AddDays(goalDay - tDay);

                result = fridayDate2;
            }
            else
            {
                result = fridayDate;
            }

            return result;
        }
        #endregion

        #region 매출매입현황
        //매출
        private DataTable GetRetr9()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" SELECT B.GUBUN1--등급                                                         ");
            strSql.AppendLine("      , COUNT(*) AS CNT--횟수                                                  ");
            strSql.AppendLine("      , SUM(CONVERT(NUMERIC, A.J_Custom)) AS SANCNT --상차량                   ");
            strSql.AppendLine("      , SUM(A.DANJUNG) AS DANJUNG --납품량                                     ");
            strSql.AppendLine("      , SUM(A.DANJUNG) / COUNT(*) AS AVGDJ--평균량                             ");
            strSql.AppendLine("      , SUM(A.HALIN) AS HALIN --감량                                           ");
            strSql.AppendLine("      , SUM(A.DANJUNG) - SUM(CONVERT(NUMERIC, A.J_Custom)) AS NAPJAN --납품잔량");
            strSql.AppendLine("  FROM INLIST A, JAJAE B                                                       ");
            strSql.AppendLine(" WHERE A.J_SERIAL = B.J_SERIAL                                                 ");
            strSql.AppendLine("   AND A.KERATYPE = '매출'                                                     ");
            strSql.AppendLine("   AND A.J_DATE BETWEEN '"+ sFrom + "' AND '"+ sTo + "'                        ");
            strSql.AppendLine("   AND A.J_LOTNO <> '4'                                                        ");
            strSql.AppendLine(" GROUP BY B.GUBUN1                                                             ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //매입
        private DataTable GetRetr10()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                                       ");
            strSql.AppendLine("     --등급별 스크랩 매입 집계 내역 테이블단가: SUM(인수량 * 기준단가) / SUM(인수량)  ");
            strSql.AppendLine("     --매입단가: SUM(도착도금액) / SUM(인수량)                                        ");
            strSql.AppendLine("     SELECT 1 AS NUM, B.GUBUN1--등급                                                            ");
            strSql.AppendLine("          , SUM(A.DANJUNG) AS DANJUNG--매입량                                         ");
            strSql.AppendLine("          , SUM(A.DANJUNG * A.MIDANGA) / SUM(A.DANJUNG) AS TDANGA--테이블단가         ");
            strSql.AppendLine("          , SUM(IKONGKEP + CKONGKEP) / SUM(A.DANJUNG) AS AVGDAN--평균매입단가         ");
            strSql.AppendLine("          , SUM(A.HALIN) AS HALIN--감량                                               ");
            strSql.AppendLine("          , '' AS BANCNT                                                            ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                                         ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                                   ");
            strSql.AppendLine("        AND A.KERATYPE = '매입'                                                       ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                            ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                          ");
            strSql.AppendLine("        AND B.DAEGUBUN <> '슈레더'                                                    ");
            strSql.AppendLine("        AND B.DAEGUBUN <> '인센티브'                                                  ");
            strSql.AppendLine("      GROUP BY B.GUBUN1                                                               ");
            strSql.AppendLine(" ), TEMP2 AS(                                                                         ");
            strSql.AppendLine("     --등급별 슈레더 매입 집계 내역 테이블단가: SUM(인수량 * 기준단가) / SUM(인수량)  ");
            strSql.AppendLine("     --매입단가: SUM(도착도금액) / SUM(인수량)                                        ");
            strSql.AppendLine("     SELECT 3 AS NUM, B.GUBUN1--등급                                                            ");
	        strSql.AppendLine("          , SUM(A.DANJUNG) AS DANJUNG --매입량                                        ");
	        strSql.AppendLine("          , SUM(A.DANJUNG * A.MIDANGA) / SUM(A.DANJUNG) AS TDANGA--테이블단가         ");
	        strSql.AppendLine("          , SUM(IKONGKEP + CKONGKEP) / SUM(A.DANJUNG) AS AVGDAN--평균매입단가         ");
	        strSql.AppendLine("          , SUM(A.HALIN) AS HALIN --감량                                              ");
	        strSql.AppendLine("          , '' AS BANCNT                                                            ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                                         ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                                   ");
            strSql.AppendLine("        AND A.KERATYPE = '매입'                                                       ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+sFrom+"' AND '"+sTo+"'                            ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                                          ");
            strSql.AppendLine("        AND B.DAEGUBUN = '슈레더'                                                     ");
            strSql.AppendLine("      GROUP BY B.GUBUN1                                                               ");
            strSql.AppendLine(" ), TEMP3 AS(                                                                         ");
            strSql.AppendLine("     SELECT *                                                                         ");
            strSql.AppendLine("       FROM TEMP1                                                                     ");
            strSql.AppendLine("      UNION ALL                                                                       ");
            strSql.AppendLine("     SELECT 2 ,'스크랩 합계'                                                             ");
	        strSql.AppendLine("          , SUM(DANJUNG)                                                              ");
	        strSql.AppendLine("          , NULL                                                                      ");
	        strSql.AppendLine("          , NULL                                                                      ");
	        strSql.AppendLine("          , SUM(HALIN)                                                                ");
	        strSql.AppendLine("          , CONVERT(VARCHAR,ROUND((SUM(HALIN)/SUM(DANJUNG))*100,2))+'%'                                                 ");
            strSql.AppendLine("       FROM TEMP1                                                                     ");
            strSql.AppendLine("      UNION ALL                                                                       ");
            strSql.AppendLine("     SELECT *                                                                         ");
            strSql.AppendLine("       FROM TEMP2                                                                     ");
            strSql.AppendLine("      UNION ALL                                                                       ");
            strSql.AppendLine("     SELECT 4, '폐압 합계'                                                               ");
	        strSql.AppendLine("          , SUM(DANJUNG)                                                              ");
	        strSql.AppendLine("          , NULL                                                                      ");
	        strSql.AppendLine("          , NULL                                                                      ");
	        strSql.AppendLine("          , SUM(HALIN)                                                                ");
	        strSql.AppendLine("          , CONVERT(VARCHAR,ROUND((SUM(HALIN)/SUM(DANJUNG))*100,2))+'%'                                                ");
            strSql.AppendLine("       FROM TEMP2                                                                     ");
            strSql.AppendLine(" )                                                                                    ");
            strSql.AppendLine(" SELECT *                                                                             ");
            strSql.AppendLine("   FROM TEMP3                                                                         ");
            strSql.AppendLine("  UNION ALL                                                                           ");
            strSql.AppendLine(" SELECT 5, '합계'                                                                        ");
            strSql.AppendLine("      , SUM(DANJUNG)                                                                  ");
            strSql.AppendLine("      , NULL                                                                          ");
            strSql.AppendLine("      , NULL                                                                          ");
            strSql.AppendLine("      , SUM(HALIN)                                                                    ");
            strSql.AppendLine("      , CONVERT(VARCHAR,ROUND((SUM(HALIN)/SUM(DANJUNG))*100,2))+'%'                                                     ");
            strSql.AppendLine("   FROM TEMP3                                                                         ");
            strSql.AppendLine("  WHERE GUBUN1 LIKE '%합계%'                                                          ");
                                                                                                                     
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        //매출차트
        private void SetPieChartData(ChartControl chart, DataTable dt, string title)
        {
            chart.Series.Clear();

            Dictionary<string, double> totalInfo = new Dictionary<string, double>();

            foreach (DataRow row in dt.Rows)
            {
                string GUBUN = row["GUBUN"].ToString();
                double DANJUNG = double.Parse(row["DANJUNG"].ToString());

                totalInfo.Add(GUBUN, DANJUNG);

            }

            Series series = new Series("Total", ViewType.Pie3D);
            chart.Series.Add(series);

            foreach (KeyValuePair<string, double> info in totalInfo)
            {
                string GUBUN = info.Key;
                double DANJUNG = info.Value;

                SeriesPoint point = new SeriesPoint(GUBUN, DANJUNG);
                series.Points.Add(point);
            }

            //textformat
            series.Label.TextPattern = "{A} : {V:n0} ({VP:0%})";
            series.LegendTextPattern = "{A}";

            //title
            chart.Titles.Clear();

            ChartTitle chartTitle = new ChartTitle();

            chartTitle.Text = title;
            chartTitle.Font = new Font("Tahoma", 18, FontStyle.Bold);
            chartTitle.Alignment = StringAlignment.Center;
            chart.Titles.Add(chartTitle);
        }

        private void SetChart1(DataTable dt)
        {
            DataTable chartDt = new DataTable();

            chartDt.Columns.Add("GUBUN");
            chartDt.Columns.Add("DANJUNG");

            if(dt != null)
            {
                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    string sGubun = dt.Rows[i]["GUBUN1"]?.ToString();
                    string sDANJUNG = dt.Rows[i]["DANJUNG"]?.ToString();

                    double dDANJUNG = 0;

                    double.TryParse(sDANJUNG, out dDANJUNG);

                    dDANJUNG = dDANJUNG / 1000;

                    DataRow row = chartDt.NewRow();

                    row["GUBUN"] = sGubun;
                    row["DANJUNG"] = dDANJUNG.ToString("n0");

                    chartDt.Rows.Add(row);
                }
                
                SetPieChartData(chartControl1, chartDt, "매출량");
            }
        }

        private void SetChart2()
        {
            string sFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" SET ANSI_WARNINGS OFF");
            strSql.AppendLine(" SET ARITHIGNORE ON   ");
            strSql.AppendLine(" SET ARITHABORT OFF;  ");
            strSql.AppendLine(" WITH TEMP1 AS(                                                            ");
            strSql.AppendLine("     SELECT B.MAIPGB AS GUBUN1                                             ");
            strSql.AppendLine("          , SUM(A.DANJUNG) AS DANJUNG--매입량                              ");
            strSql.AppendLine("       FROM INLIST A, JAJAE B                                              ");
            strSql.AppendLine("      WHERE A.J_SERIAL = B.J_SERIAL                                        ");
            strSql.AppendLine("        AND A.KERATYPE = '매입'                                            ");
            strSql.AppendLine("        AND A.J_DATE BETWEEN '"+ sFrom + "' AND '"+ sTo + "'                 ");
            strSql.AppendLine("        AND A.J_LOTNO <> '4'                                               ");
            strSql.AppendLine("        AND B.DAEGUBUN <> '인센티브'                                       ");
            strSql.AppendLine("      GROUP BY B.MAIPGB                                                    ");
            strSql.AppendLine(" )                                                                         ");
            strSql.AppendLine(" SELECT SUM(CASE WHEN GUBUN1 = '중량물' THEN DANJUNG END)/ 1000 AS 중량물    ");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '경량A' THEN DANJUNG END)/ 1000 AS 경량A   ");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '경량B' THEN DANJUNG END)/ 1000 AS 경량B    ");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '기계작업철' THEN DANJUNG END)/ 1000 AS 기계작업철");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '모재' THEN DANJUNG END)/ 1000 AS 모재      ");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '특급모재' THEN DANJUNG END)/ 1000 AS 특급모재 ");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '슈레더C' THEN DANJUNG END)/ 1000 AS 슈레더C    ");
            strSql.AppendLine("      , SUM(CASE WHEN GUBUN1 = '슈레더A' THEN DANJUNG END)/ 1000 AS 슈레더A    ");
            strSql.AppendLine("   FROM TEMP1                                                              ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            DataTable chartDt = new DataTable();

            chartDt.Columns.Add("GUBUN");
            chartDt.Columns.Add("DANJUNG");

            if (dt != null)
            {
                for(int i=0;i< dt.Columns.Count; i++)
                {
                    string sGubun = dt.Columns[i].ColumnName;
                    string sDANJUNG = dt.Rows[0][sGubun]?.ToString();

                    DataRow row = chartDt.NewRow();

                    row["GUBUN"] = sGubun;
                    row["DANJUNG"] = chgFormat(sDANJUNG, "double", "n0");

                    chartDt.Rows.Add(row);
                }

                SetPieChartData(chartControl2, chartDt, "매입량");
            }
        }

        #endregion

        #endregion

        #region 기타 메소드

        //string displayformat 변경
        private string chgFormat(string target, string type, string format)
        {
            string result = "";
            if (type.Equals("double"))
            {
                double.TryParse(target, out double dResult);

                result = dResult.ToString(format);
            }

            return result;
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
    }
}