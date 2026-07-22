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
    public partial class AccFieldCustom : DevExpress.XtraEditors.XtraForm
    {
        public AccFieldCustom()
        {
            InitializeComponent();
        }
        
        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccFieldCustom_Load(object sender, EventArgs e)
        {
            
            Cursor = Cursors.Default;
            DateEditFrom.EditValue = DateTime.Today.ToString("yyyy-MM-01");
            DateEditTo.EditValue = DateTime.Today;

            SetGridLookupEdit(GLkupEditCate, "N", "1");
            GLkupEditCate.Properties.View.PopulateColumns(GLkupEditCate.Properties.DataSource);
            GLkupEditCate.Properties.View.Columns[GLkupEditCate.Properties.ValueMember].Visible = false;

            

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewBuyerPurc, GridViewCompany };
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

            BtnRetr_Click(null, null);

        }

        
        private void SetGridGroup2()
        {
            GridViewBuyerPurc.GroupSummary.Clear();
            GridGroupSummaryItem itemNM = new GridGroupSummaryItem();
            itemNM.FieldName = "DEALER_NM";
            itemNM.SummaryType = DevExpress.Data.SummaryItemType.Max;
            itemNM.DisplayFormat = "상호명 : {0}";
            GridViewBuyerPurc.GroupSummary.Add(itemNM);
        }
        

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

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

            string sDept = "";
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

            PurchacePerBuyer(sDept, sYmdFrom, sYmdTo, sDealGB, sBigCate);
            //if (RdGbDealGB.SelectedIndex == 0)
            //{
            //    PurchacePerBuyer(sDept, sYmdFrom, sYmdTo, sDealGB, sBigCate);
                
            //}
            

            XtTControl_SelectedPageChanged(null, null);

            //GridColPurcCompNm.Group();

            //SetGridGroup2();
            //GridViewBuyerPurc.ExpandAllGroups();
            Cursor = Cursors.Default;

        }

        private void RdoDealGB_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void PurchacePerBuyer(string sDept, string sYmdFrom, string sYmdTo, string jType, string sBigCate)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            //#001
            strSql.AppendLine(" SELECT *                                                                 ");
            strSql.AppendLine("      , LANDEDWEIGHT - LOSS AS TOTWEIGHT                                  ");
            strSql.AppendLine("      , LANDEDWEIGHT - LOSS - ACCEPTWEIGHT AS DIFFWEIGHT                                  ");
            strSql.AppendLine("      , ARRVUNITPRICE - STDDUNITPRICE AS DIFFUNITPRICE                    ");
            strSql.AppendLine("   FROM(                                                                  ");
            strSql.AppendLine("          SELECT A.J_DATE                                                 ");
            strSql.AppendLine("               , A.JUNPYOID                                               ");
            strSql.AppendLine("               , G.J_DATE AS MESURE_DT                                    ");
            strSql.AppendLine("               , D.DEALER_NM                                              ");
            strSql.AppendLine("               , E.EMP_NM AS INSPECTOR                                    ");
            strSql.AppendLine("               , G.J_BNUM                                                 ");
            strSql.AppendLine("               , G.WEIGHT AS LANDEDWEIGHT                                 ");
            strSql.AppendLine("               , A.HALIN AS LOSS                                          ");
            strSql.AppendLine("               , A.IWEIGHT AS ACCEPTWEIGHT                                ");
            strSql.AppendLine("               , F.GUBUN1                                                 ");
            strSql.AppendLine("               , A.MIDANGA AS STDDUNITPRICE                               ");
            strSql.AppendLine("               , A.DANGA AS PAYEDUNITPRICE                                ");
            strSql.AppendLine("               , ((ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) + (ISNULL(G.ETC_COST1, 0) + ISNULL(G.ETC_COST2, 0))) / A.IWEIGHT AS ARRVUNITPRICE");
            strSql.AppendLine("               , ISNULL(A.CKONGKEP, 0) AS CARRYCOST                                                      ");
            strSql.AppendLine("               , A.IKONGKEP AS TOTALPRICE                                                                ");
            strSql.AppendLine("               , (ISNULL(A.IWEIGHT, 0) * ISNULL(A.DANGA, 0)) + ISNULL(A.CKONGKEP, 0) AS ARRVTOTALPRICE   ");
            strSql.AppendLine("               , G.J_STATE AS LOSSCAUSE                                                                  ");
            strSql.AppendLine("               , G.JUNPYOID AS IMAGE                                                                     ");
            strSql.AppendLine("               , G.GUMSUBIGO AS INSPECTNOTE                                                              ");
            strSql.AppendLine("               , ROUND((ISNULL(A.CKONGKEP, 0) / A.DANJUNG), 1, 0) AS CARRY_UNIT_PRICE                    ");
            strSql.AppendLine("               , H.DEALER_NM AS ETC_DEALER_NM1                                                           ");
            strSql.AppendLine("               , ISNULL(G.ETC_COST1, 0) AS ETC_COST1                                                     ");
            strSql.AppendLine("               , G.ETC_REMARK1                                                                           ");
            strSql.AppendLine("               , I.DEALER_NM AS ETC_DEALER_NM2                                                           ");
            strSql.AppendLine("               , ISNULL(G.ETC_COST2, 0) AS ETC_COST2                                                     ");
            strSql.AppendLine("               , ISNULL(G.ETC_COST1, 0) + ISNULL(G.ETC_COST1, 0) AS ETC_COST                             ");
            strSql.AppendLine("               , G.ETC_REMARK2                                                                           ");
            strSql.AppendLine("            FROM INLIST A                                                                                ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE C                                                                      ");
            strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL                                                                 ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D                                                              ");
            strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1                                                                   ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E                                                               ");
            strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID                                                                    ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE F                                                                      ");
            strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1                                                                     ");
            strSql.AppendLine("            LEFT OUTER JOIN MESURING G                                                                   ");
            strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID                                                              ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H                                                                    ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)                                         ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I                                                                    ");
            strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)                                         ");
            strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "'                                                          ");
            strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "'                                                            ");
            strSql.AppendLine("             AND A.KERATYPE = '" + jType + "'                                                            ");
            strSql.AppendLine("             AND G.WEIGHT_GUBUN = 1                                                                      ");
            strSql.AppendLine("             AND D.DEALER_NM = '" + sDept + "' ");
            strSql.AppendLine("             AND G.KERATYPE <> '직송'                                                                    ");
            strSql.AppendLine("             AND G.J_CHECK = '1'                                                                         ");
            if (!string.IsNullOrEmpty(sBigCate)) strSql.AppendLine("    AND F.DAEGUBUN IN (" + sBigCate + ")                            ");
            strSql.AppendLine("      ) Y1                                                                                               ");
            strSql.AppendLine("  ORDER BY Y1.J_DATE, Y1.DEALER_NM                                                                       ");

           

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridBuyer.MainView = GridViewBuyerPurc;
            GridBuyer.DataSource = dt;

            Cursor = Cursors.Default;

        }

        private void ClearFps()
        {
            GridBuyer.DataSource = null;
           
        }

        public enum WEIGHT_GUBUN { Purc, Sale, None }
        public WEIGHT_GUBUN _Weight_Gubun = WEIGHT_GUBUN.None;
        private void XtTControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            
            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);
            string sDealGB = RdGbDealGB.EditValue.ToString();

          

            if (XtTControl.SelectedTabPage.Name.Equals("XtTCCompany") )
            {
                StringBuilder strSql = new StringBuilder();
                Cursor = Cursors.WaitCursor;
                strSql.Clear();
                strSql.AppendLine(" ");


                strSql.AppendLine(" SELECT *                                                                 ");
                strSql.AppendLine("      , LANDEDWEIGHT - LOSS AS TOTWEIGHT                                  ");
                strSql.AppendLine("      , LANDEDWEIGHT - LOSS - ACCEPTWEIGHT AS DIFFWEIGHT                                  ");
                strSql.AppendLine("   FROM(                                                                  ");
                strSql.AppendLine("          SELECT D.DEALER_NM                                              ");
                strSql.AppendLine("               , SUM(G.WEIGHT) AS LANDEDWEIGHT                                 ");
                strSql.AppendLine("               , SUM(A.HALIN) AS LOSS                                          ");
                strSql.AppendLine("               , SUM(A.IWEIGHT) AS ACCEPTWEIGHT                                ");
                strSql.AppendLine("               , SUM(A.IKONGKEP) AS TOTALPRICE                                                                ");
                strSql.AppendLine("               , (ISNULL(SUM(A.IWEIGHT), 0) * ISNULL(SUM(A.DANGA), 0)) + ISNULL(SUM(A.CKONGKEP), 0) AS ARRVTOTALPRICE   ");
                strSql.AppendLine("            FROM INLIST A                                                                                ");
                strSql.AppendLine("            LEFT OUTER JOIN JAJAE C                                                                      ");
                strSql.AppendLine("              ON C.J_SERIAL = A.J_SERIAL                                                                 ");
                strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD D                                                              ");
                strSql.AppendLine("              ON D.DEALER_CD = A.J_ID1                                                                   ");
                strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS E                                                               ");
                strSql.AppendLine("              ON E.EMP_ID = D.CHRG_ID                                                                    ");
                strSql.AppendLine("            LEFT OUTER JOIN JAJAE F                                                                      ");
                strSql.AppendLine("              ON F.GUBUN1 = C.GUBUN1                                                                     ");
                strSql.AppendLine("            LEFT OUTER JOIN MESURING G                                                                   ");
                strSql.AppendLine("              ON G.IPCHULGO_MAIPID = A.J_ID                                                              ");
                strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD H                                                                    ");
                strSql.AppendLine("              ON G.ETC_DEALER_CD1 = CONVERT(VARCHAR,H.DEALER_CD)                                         ");
                strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD I                                                                    ");
                strSql.AppendLine("              ON G.ETC_DEALER_CD2 = CONVERT(VARCHAR,I.DEALER_CD)                                         ");
                strSql.AppendLine("           WHERE A.J_DATE >= '" + sYmdFrom + "'                                                          ");
                strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "'                                                            ");
                strSql.AppendLine("             AND A.KERATYPE = '매입'                                                                     ");
                strSql.AppendLine("             AND G.WEIGHT_GUBUN = 1                                                                      ");
                strSql.AppendLine("             AND G.KERATYPE <> '직송'                                                                    ");
                strSql.AppendLine("             AND G.J_CHECK = '1'                                                                         ");
                strSql.AppendLine("           GROUP BY D.DEALER_NM                                                                                  ");
                strSql.AppendLine("      ) Y1                                                                                               ");
                strSql.AppendLine("  ORDER BY  DIFFWEIGHT                                                                       ");
                

                DataTable dtCompany = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                GridCompany.MainView = GridViewCompany;
                GridCompany.DataSource = dtCompany;

               

            }
            Cursor = Cursors.Default;
        }

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

                //strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                //strSql.AppendLine("      , A.COM_NM AS NM ");
                //strSql.AppendLine("   FROM COM_BASE_CD A ");
                //strSql.AppendLine("  WHERE A.CD_GB = '0002' ");
                //strSql.AppendLine("    AND A.COM_CD NOT IN('E', 'F')");
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
                    
                    else
                    {
                        GridCompany.ExportToXls(FileName);
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

       

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AccFieldCustom_KeyDown(object sender, KeyEventArgs e)
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

        private void AccFieldCustom_FormClosed(object sender, FormClosedEventArgs e)
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

        private void AccFieldCustom_TextChanged(object sender, EventArgs e)
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


        private void GLkupEditCate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewCompany_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sDept = Convert.ToString(GridViewCompany.GetFocusedRowCellValue("DEALER_NM") ?? string.Empty);
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
                PurchacePerBuyer(sDept, sYmdFrom, sYmdTo, sDealGB, sBigCate);

            }
        }
    }
}