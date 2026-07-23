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
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.Data;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DevExpress.XtraGrid;

/*
 * 작성일자 : 2023-12-20
 * 작성자 : 이상탁
 * ---------------------HISTORY-----------------------
 * 1. 현장매입출일보 복사 
 *
 */
namespace AccAdm
{
    public partial class AccFieldVs : DevExpress.XtraEditors.XtraForm
    {
        public AccFieldVs()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccFieldVs_Load(object sender, EventArgs e)
        {
            //


            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            SetGridLookupEdit(GLkupEditCate, "N", "1");
            GLkupEditCate.Properties.View.PopulateColumns(GLkupEditCate.Properties.DataSource);
            GLkupEditCate.Properties.View.Columns[GLkupEditCate.Properties.ValueMember].Visible = false;



            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewBuyerPurc, GridViewBuyerSales, GridViewBuyerDtSend };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        /*
        private void SetGridGroup2()
        {
            GridViewBuyerPurc.GroupSummary.Clear();
            GridGroupSummaryItem itemNM = new GridGroupSummaryItem();
            itemNM.FieldName = "DEALER_NM";
            itemNM.SummaryType = DevExpress.Data.SummaryItemType.Max;
            itemNM.DisplayFormat = "상호명 : {0}";
            GridViewBuyerPurc.GroupSummary.Add(itemNM);
        }
        */

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }

            ClearFps();

            string sYmd = DateEditFrom.EditValue.ToString();

            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);
            string sDealGB = RdGbDealGB.EditValue.ToString();

            if (sDealGB == "A")
            {
                sDealGB = "매입";
            }
            else if (sDealGB == "B")
            {
                sDealGB = "매출";
            }
            else
            {
                sDealGB = "직송";
            }

            string sBigCate = string.Empty;

            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.ValueMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0)
            {
                sBigCate = sBigCate.Substring(0, sBigCate.Length - 1);
            }
            else
            {
                sBigCate = null;
            }

            if (RdGbDealGB.SelectedIndex == 0)
            {
                PurchacePerBuyer(sYmdFrom, sYmdTo, sDealGB, sBigCate);
                
            }
            

            //XtTControl_SelectedPageChanged(null, null);

            //GridColPurcCompNm.Group();
            //SetGridGroup2();
            //GridViewBuyerPurc.ExpandAllGroups();

        }

        private void RdoDealGB_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void PurchacePerBuyer(string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            //#001
            strSql.AppendLine(" SELECT * , (Y1.iKongkep + Y1.BUGASE1 - ACCPRIC) AS DIFFUNITPRICE                                                                        ");
            strSql.AppendLine("     FROM (                                                                                                                              ");
            strSql.AppendLine("             SELECT	A.JUNPYOID, A.J_Date ,D.DEALER_NM ,A.CHRG_NM ,A.Danjung ,A.Gubun1 ,A.MiDanga ,A.Danga ,A.iKongKep ,A.SeakPoham ,    ");
            //strSql.AppendLine("                     CASE	WHEN	A.SeakPoham  = 'N'	THEN A.Bugase                                                                   ");
            strSql.AppendLine("                     CASE	WHEN	A.SeakPoham  = 'N'	THEN A.iKongKep * 0.1                                                           ");
            strSql.AppendLine("                     ELSE	0                                                                                                           ");
            strSql.AppendLine("                     END AS BUGASE1 ,                                                                                                    ");
            strSql.AppendLine("                     (SELECT sum(a2.ADAMT)  FROM actran a2 WHERE a2.REF1 = CONVERT (varchar,A.junpyoid)  GROUP by a2.REF1)   AS ACCPRIC  ");
            strSql.AppendLine("             FROM	inlist A                                                                                                            ");
            strSql.AppendLine("                     LEFT OUTER JOIN ACC_DEALER_CD D                                                                                     ");
            strSql.AppendLine("                         ON D.DEALER_CD = A.J_ID1                                                                                        ");
            strSql.AppendLine("             WHERE       A.J_DATE >= '" + sYmdFrom + "'                                                                                  ");
            strSql.AppendLine("                     AND A.J_DATE <= '" + sYmdTo + "'                                                                                    ");
            strSql.AppendLine("                     AND 	A.KeraType = '매입'                                                                                         ");
            strSql.AppendLine("                     AND 	A.J_LotNo <> '4'                                                                                            ");
            strSql.AppendLine("                     AND 	A.APRV_YN  = 'Y'                                                                                            ");
            strSql.AppendLine("         ) Y1                                                                                                                            ");
            strSql.AppendLine("ORDER BY Y1.J_DATE, Y1.DEALER_NM                                                                                                         ");
            
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerPurc;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;

        }

       
        

        

       

        private void ClearFps()
        {
            GridBuyer.DataSource = null;
            //GridGrade.DataSource = null;
        }

        public enum WEIGHT_GUBUN { Purc, Sale, None }
        public WEIGHT_GUBUN _Weight_Gubun = WEIGHT_GUBUN.None;
       
        private void RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {

        }

        private void SetGridLookupEdit(DevExpress.XtraEditors.GridLookUpEdit gLkup, string sNullYn, string sGb)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '전체' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }
            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT DAEGUBUN AS CD ");
                strSql.AppendLine("      , DAEGUBUN AS NM ");
                strSql.AppendLine("   FROM JAJAE ");
                strSql.AppendLine("  WHERE DAEGUBUN IN ('고철A', '고철B', '슈레더') ");
                strSql.AppendLine("  GROUP BY DAEGUBUN  ");

                
            }
            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gLkup.Properties.DataSource = dt;
            gLkup.Properties.DisplayMember = "NM";
            gLkup.Properties.ValueMember = "CD";
        }

        private void PerBuyerRowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;

            if (GridBuyer.MainView == GridViewBuyerPurc)
            {
                if (e.Column.FieldName == "PAYEDUNITPRICE")
                {
                    string sStdd = view.GetRowCellDisplayText(e.RowHandle, view.Columns["STDDUNITPRICE"]).ToString();
                    string dPay = view.GetRowCellDisplayText(e.RowHandle, view.Columns["PAYEDUNITPRICE"]).ToString();

                    double dStddUtPrc = String.IsNullOrEmpty(sStdd) ? 0 : Convert.ToDouble(sStdd);
                    double dPayUtPrc = String.IsNullOrEmpty(dPay) ? 0 : Convert.ToDouble(dPay);

                    if (dStddUtPrc > dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.RoyalBlue;
                    }
                    else if (dStddUtPrc < dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.IndianRed;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            if (GridBuyer.MainView == GridViewBuyerSales)
            {
                if (e.Column.FieldName == "SALEUNITPRICE")
                {
                    string sStdd = view.GetRowCellDisplayText(e.RowHandle, view.Columns["STDDUNITPRICE"]).ToString();
                    string sSale = view.GetRowCellDisplayText(e.RowHandle, view.Columns["SALEUNITPRICE"]).ToString();

                    double dStddUtPrc = String.IsNullOrEmpty(sStdd) ? 0 : Convert.ToDouble(sStdd);
                    double dPayUtPrc = String.IsNullOrEmpty(sSale) ? 0 : Convert.ToDouble(sSale);

                    if (dStddUtPrc > dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.RoyalBlue;
                    }
                    else if (dStddUtPrc < dPayUtPrc)
                    {
                        e.Appearance.BackColor = Color.IndianRed;
                    }
                    else
                    {
                        return;
                    }
                }
                if (e.Column.FieldName == "DIFFWEIGHT")
                {
                    string sDiff = view.GetRowCellDisplayText(e.RowHandle, view.Columns["DIFFWEIGHT"]).ToString();

                    double dDiffUtPrc = String.IsNullOrEmpty(sDiff) ? 0 : Convert.ToDouble(sDiff);

                    if (dDiffUtPrc > 0) e.Appearance.BackColor = Color.IndianRed;
                }
            }

            
        }
        private void PerGradeRowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;

            if (e.Column.FieldName == "DIFFWEIGHT")
            {
                string sDiff = view.GetRowCellDisplayText(e.RowHandle, view.Columns["DIFFWEIGHT"]).ToString();
                double dDiffWeight = String.IsNullOrEmpty(sDiff) ? 0 : Convert.ToDouble(sDiff);

                if (dDiffWeight != 0)
                {
                    e.Appearance.BackColor = Color.IndianRed;
                }
            }
        }


        private void GLkupEditCate_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            string sBigCate = string.Empty;

            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.DisplayMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0) e.DisplayText = sBigCate.Substring(0, sBigCate.Length - 1);
        }

        private void GLkupEditCate_CloseUp(object sender, DevExpress.XtraEditors.Controls.CloseUpEventArgs e)
        {
            BtnRetr.Focus();
        }

        private void DateEditFrom_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void DateEditTo_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                string sDealType = RdGbDealGB.EditValue?.ToString();

                if (sDealType.Equals("A"))
                {
                    sDealType = "매입";
                }
                else if (sDealType.Equals("B"))
                {
                    sDealType = "매출";
                }
                else if (sDealType.Equals("C"))
                {
                    sDealType = "직송";
                }

                //string sFileNM = "현장매입출일보 " + sYmdFrom + "~" + sYmdTo + " " + sDealType;
                string sFileNM = this.Text;
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    if (XtTControl.SelectedTabPage.Name.ToString().Equals("XtTCRetr1"))
                    {
                        XlsxExportOptions options = new XlsxExportOptions();
                        options.ExportMode = XlsxExportMode.SingleFilePageByPage;
                        //CompositeLink link = new CompositeLink();

                        GridBuyer.ExportToXlsx(FileName, options);
                        Process.Start(FileName);
                    }
                   
                }
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

        private void SaveMultiGridViewToExcel(GridView[] ArrGridViews)
        {
            //int num = 0;

            //object missingType = Type.Missing;
            //Excel.Application objApp;
            //Excel._Workbook objBook;
            //Excel.Workbooks objBooks;
            //Excel.Sheets objSheets;
            //Excel._Worksheet objSheet;
            //Excel.Range range;

            //string[] headers = new string[myDataGridViews[0].ColumnCount];
            //string[] columns = new string[myDataGridViews[0].ColumnCount];
            //for (int c = 0; c < myDataGridViews[0].ColumnCount; c++)
            //{
            //    headers[c] = myDataGridViews[0].Rows[0].Cells[c].OwningColumn.HeaderText.ToString();
            //    if (c <= 25)
            //    {
            //        num = c + 65;
            //        columns[c] = Convert.ToString((char)num);
            //    }
            //    else
            //    {
            //        columns[c] = Convert.ToString((char)(Convert.ToInt32(c / 26) - 1 + 65)) + Convert.ToString((char)(c % 26 + 65));
            //    }
            //}

            //try
            //{
            //    objApp = new Excel.Application();
            //    objBooks = objApp.Workbooks;
            //    objBook = objBooks.Add(Missing.Value);
            //    objSheets = objBook.Worksheets;
            //    objSheet = (Excel._Worksheet)objSheets.get_Item(1);

            //    if (captions)
            //    {
            //        for (int c = 0; c < myDataGridViews[0].ColumnCount; c++)
            //        {
            //            range = objSheet.get_Range(columns[c] + "1", Missing.Value);
            //            range.set_Value(Missing.Value, headers[c]);
            //        }
            //    }

            //    int iGridViewNum = 0;
            //    foreach (DataGridView myDataGridView in myDataGridViews)
            //    {
            //        columns = new string[myDataGridView.ColumnCount];
            //        for (int i = 0; i < myDataGridView.RowCount; i++)
            //        {
            //            for (int j = 0; j < myDataGridView.ColumnCount; j++)
            //            {
            //                if (j <= 25)
            //                {
            //                    num = j + 65;
            //                    columns[j] = Convert.ToString((char)num);
            //                }
            //                else
            //                {
            //                    columns[j] = Convert.ToString((char)(Convert.ToInt32(j / 26) - 1 + 65)) + Convert.ToString((char)(j % 26 + 65));
            //                }

            //                range = objSheet.get_Range(columns[j] + Convert.ToString(iGridViewNum + i + 2), Missing.Value);
            //                range.set_Value(Missing.Value, myDataGridView.Rows[i].Cells[j].Value.ToString());
            //            }
            //        }
            //        iGridViewNum++;
            //    }

            //    objApp.Visible = false;
            //    objApp.UserControl = false;
            //    objBook.SaveAs(@saveFileDialog.FileName,
            //        Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
            //        missingType, missingType, missingType, missingType,
            //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            //        missingType, missingType, missingType, missingType, missingType);

            //    objBook.Close(false, missingType, missingType);
            //    Cursor.Current = Cursors.Default;

            //    MessageBox.Show("Save Success!!!");
            //}
            //catch (Exception theException)
            //{
            //    String errorMessage;
            //    errorMessage = "Error: ";
            //    errorMessage = String.Concat(errorMessage, theException.Message);
            //    errorMessage = String.Concat(errorMessage, " Line: ");
            //    errorMessage = String.Concat(errorMessage, theException.Source);
            //    MessageBox.Show(errorMessage, "Error");
            //}
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AccFieldVs_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {

            }
            else if (e.KeyCode == Keys.F1)
            {

            }
            else if (e.KeyCode == Keys.F3)
            {

            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExport_Click(null, null);
            }
        }

        private void GridBuyer_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridGrade_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void XtTControl_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewBuyerPurc_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccFieldVs_FormClosed(object sender, FormClosedEventArgs e)
        {

        }


        #region [정렬기능(2020-06-02 정은영)]
        private ColumnSortOrder GetNextSortOrder(ColumnSortOrder order)
        {
            switch (order)
            {
                case ColumnSortOrder.None: return ColumnSortOrder.Descending;
                case ColumnSortOrder.Descending: return ColumnSortOrder.Ascending;
                case ColumnSortOrder.Ascending: return ColumnSortOrder.None;
            }

            return ColumnSortOrder.None;
        }

        private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        {

            if (sender.GetType() == GridViewBuyerPurc.GetType())
            {
                BandedGridView view = (BandedGridView)sender;
                BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

                if (hitInfo.InBandPanel)
                {
                    if (hitInfo.Band.Name.Equals("gridBand11") || hitInfo.Band.Name.Equals("gridBand30")) //상호
                    {
                        ColumnSortOrder order = view.Columns["DEALER_NM"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "상호↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "상호↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "상호";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand12") || hitInfo.Band.Name.Equals("gridBand31") || hitInfo.Band.Name.Equals("gridBand49")) //담당자
                    {
                        ColumnSortOrder order = view.Columns["INSPECTOR"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "담당자↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "담당자↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "담당자";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand18") || hitInfo.Band.Name.Equals("gridBand37") || hitInfo.Band.Name.Equals("gridBand50")) //등급
                    {
                        ColumnSortOrder order = view.Columns["GUBUN1"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "등급↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "등급↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "등급";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand52")) //매출처
                    {
                        ColumnSortOrder order = view.Columns["DEALER_NM"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "매출처↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "매출처↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "매출처";
                    }
                    else if (hitInfo.Band.Name.Equals("gridBand56")) //매입처
                    {
                        ColumnSortOrder order = view.Columns["J_BOOKING"].SortOrder;
                        order = GetNextSortOrder(order);

                        hitInfo.Band.Columns[0].SortOrder = order;

                        view.FocusedRowHandle = 0;
                        if (order == ColumnSortOrder.Descending)
                            hitInfo.Band.Caption = "매입처↓";
                        else if (order == ColumnSortOrder.Ascending)
                            hitInfo.Band.Caption = "매입처↑";
                        else if (order == ColumnSortOrder.None)
                            hitInfo.Band.Caption = "매입처";
                    }
                }

            }
            else
            {
                GridView view = (GridView)sender;
                GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

                if (hitInfo.InColumn)
                {
                    ColumnSortOrder order = hitInfo.Column.SortOrder;

                    order = GetNextSortOrder(order);

                    hitInfo.Column.SortOrder = order;
                    view.FocusedRowHandle = 0;
                }
            }
        }
        #endregion

        private void AccFieldVs_TextChanged(object sender, EventArgs e)
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

        private void GridViewBuyerPurc_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewBuyerPurc_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColPurcPayUtPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColPurcPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
           
        }

        private void GridViewBuyerSales_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColSaleUtPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColSalePrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
            else if (e.Column == GridColSaleArrvPrc)
            {
                if (string.IsNullOrEmpty(e.CellValue?.ToString()))
                    e.Appearance.BackColor = Color.LightSalmon;
                else if (Convert.ToDouble(e.CellValue) <= 0)
                    e.Appearance.BackColor = Color.LightSalmon;
            }
        }

        private void GridViewBuyerPurc_RowClick(object sender, RowClickEventArgs e)
        {

        }

        

        private void GLkupEditCate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}