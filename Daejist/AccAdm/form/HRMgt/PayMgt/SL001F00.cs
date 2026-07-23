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
using System.Data.SqlClient;
using DevExpress.ClipboardSource.SpreadsheetML;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using ComLib;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;

namespace AccAdm
{
    public partial class SL001F00: DevExpress.XtraEditors.XtraForm
    {
        public SL001F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_SL001F00";
        
        private void SL001F00_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            ComnEtcFunc.gp_SetColorFocused(layoutControl5);
            string today1 = DateTime.Today.ToString();
            Dtstdt.EditValue = today1.Substring(0, 7);

            //사원명
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupEmpid, "HR_EMP_BASIS", "", "");
            //부서명
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupDept, "ACC_DEPT_CD", "DEPT_CD", "DEPT_NM");
            //직위
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupJikwi, "COM_BASE_CD", "GRADE_CD", "");

            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr };
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

        private void GetRetr2()
        {
            string dt_From = Dtstdt.EditValue?.ToString().Substring(0, 7);
            string cb_Search = Cbserch.SelectedIndex.ToString();
            string ed_Search = Txserch.EditValue?.ToString().Trim();
            string sNamek = Txserch.EditValue?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            try
            {
                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("INKNAM", sNamek);
                dicParams.Add("DATE_FROM", dt_From);
                dicParams.Add("FIND_IDX", cb_Search);
                dicParams.Add("FIND_WORD", ed_Search);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    GridRetr.DataSource = dt;

                    if (dt.Rows.Count > 0)
                    {
                        GridRetr.Focus();
                    }
                    else
                    {
                        Txserch.Focus();
                        Txserch.SelectAll();
                    }
                }

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 리스트 오류");
            }
        }

        private void SaveFirRetr()
        {
            string dt_From = Dtstdt.EditValue?.ToString().Substring(0, 7);

            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                //dicParams.Add("CMD", "SAVE");
                dicParams.Add("CMD", "SAVECALC");
                dicParams.Add("BASYM", dt_From);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sRESULT = dt.Rows[0]["RESULT"]?.ToString();

                    if(double.Parse(sRESULT) > 0)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        #region [버튼이벤트]

        //조회
        public void BtnRetr_Click(object sender, EventArgs e)
        {
            //if (ComnEtcFunc.GetAuthInfo("USE", ComnEtcFunc.EXS_ID, this.Name))
            //    return;

            //전체리스트
            GetRetr2();
        }

        //변동급여계산
        private void BtnCalc_Click(object sender, EventArgs e)
        {
            SaveFirRetr();

            BtnRetr.PerformClick();
        }

        //추가
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            //SL001F00_POP01 frm = new SL001F00_POP01();

            //frm.Owner = this;
            //frm._ModeGubun = SL001F00_POP01.ModeGubun.Add;
            //frm.Dtstdt2 = Dtstdt.EditValue?.ToString().Substring(0, 7);
            ////frm.DataRowSendEvent += new IN001F01.SendDataHandler(GetEMPID);
            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //    BtnRetr.PerformClick();
            //}

            //if (GridViewRetr.FocusedRowHandle == GridViewRetr.RowCount - 1
            //        || GridViewRetr.RowCount == 0)
            //{
            //    string sEmpid = GridViewRetr.GetFocusedRowCellValue("EMPID")?.ToString();

            //    if (GridViewRetr.FocusedRowHandle >= 0 && string.IsNullOrEmpty(sEmpid))
            //    {
            //        return;
            //    }

            //    GridViewRetr.AddNewRow();
            //    GridViewRetr.UpdateCurrentRow();

            //    GridViewRetr.FocusedColumn = GridColEmpid;
            //}
            //else
            //{
            //    GridViewRetr.Focus();
            //    GridViewRetr.ClearSelection();
            //    GridViewRetr.FocusedRowHandle = GridViewRetr.RowCount - 1;
            //    GridViewRetr.SelectRow(GridViewRetr.RowCount - 1);
            //}

            if (GridViewRetr.RowCount == 0)
            {
                GridViewRetr.AddNewRow();
                GridViewRetr.UpdateCurrentRow();

                GridViewRetr.FocusedColumn = GridColEmpid;
            }
            else
            {
                string sEmpid = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, "EMPID")?.ToString();

                if (!string.IsNullOrEmpty(sEmpid))
                {
                    GridViewRetr.AddNewRow();
                    GridViewRetr.UpdateCurrentRow();

                    GridViewRetr.FocusedColumn = GridColEmpid;
                }
            }
        }

        //엑셀
        private void BtnExcel_Click(object sender, EventArgs e)
        {
            string sPath = string.Format(@"{0}\Temp_File\", Application.StartupPath);
            string sFileName = "변동급여 관리_" + DateTime.Now.ToString("yyyyMM") + ".xls";
            ComnEtcFunc.ExportExcelFile(sPath, sFileName, GridRetr);
        }

        //닫기
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        #endregion

        #region [그리드스타일]

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }
        #endregion

        private void SL001F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }
        
        private void BtnDel_Click(object sender, EventArgs e)
        {
            DelInfo();
        }
        private void DelInfo()
        {
            DataTable dt = GridRetr.DataSource as DataTable;

            List<int> selRows = new List<int>();

            if (dt != null)
            {
                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    string sChk = dt.Rows[i]["CHK"]?.ToString();

                    if (string.IsNullOrEmpty(sChk))
                        return;

                    if (sChk.Equals("Y"))
                    {
                        selRows.Add(i);
                    }
                }
            }

            if (selRows.Count == 0)
            {
                XtraMessageBox.Show("삭제할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("{0}의 건에 대하여 삭제를 진행하시겠습니까?", selRows.Count), "삭제확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selRows.Count; i++)
                {
                    string sEMPID = string.Empty;
                    string sBASYM = string.Empty;

                    sEMPID = dt.Rows[selRows[i]]["EMPID"]?.ToString();
                    sBASYM = dt.Rows[selRows[i]]["BASYM"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine("DELETE FROM PCHGF    ");
                    strSql.AppendLine(" WHERE EMPID = '" + sEMPID + "' AND  BASYM = '" + sBASYM + "'  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = selRows[0] - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            DataTable dtRetr = GridRetr.DataSource as DataTable;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dtRetr);
            DataTable dtMerge = dsSave.Tables[0];

            if (dtMerge == null || dtMerge.Rows.Count == 0)
            {
                XtraMessageBox.Show("변경된 데이터가 없습니다.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("편집한 데이터를 모두 저장하시겠습니까?"), "저장확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sEmpid = dtMerge.Rows[i]["EMPID"]?.ToString();

                if (string.IsNullOrEmpty(sEmpid))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            DataTable dtResult = SavePCHGF(dtMerge);

            if (dtResult != null && dtResult.Rows.Count > 0)
            {
                XtraMessageBox.Show(dtResult.Rows[0]["MSG"]?.ToString());

                if (double.Parse(dtResult.Rows[0]["RESULT"]?.ToString()) > 0)
                {
                    GetRetr2();
                }
            }
        }

        private DataTable SavePCHGF(DataTable dtModi)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            string json = ComnEtcFunc.DataTableToJsonObj(dtModi);

            dicParams.Clear();
            dicParams.Add("CMD", "SAVE2");
            dicParams.Add("JSON", json);
            dicParams.Add("USER", FmMainToolBar2.drUser["USRCD"]?.ToString());

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);

            return dt;
        }

        private void RepoGridLkupEmpid_EditValueChanged(object sender, EventArgs e)
        {
            GridLookUpEditBase lookUpEdit = (GridLookUpEditBase)sender;
            string sEmpid = lookUpEdit.EditValue?.ToString();
            string sBasym = Dtstdt.EditValue?.ToString().Substring(0, 7);

            if (!string.IsNullOrEmpty(sEmpid))
            {
                DataTable dtRetr = GridRetr.DataSource as DataTable;

                if(dtRetr != null)
                {
                    for(int i=0; i < dtRetr.Rows.Count; i++)
                    {
                        string sEmpid_1 = dtRetr.Rows[i]["EMPID"]?.ToString();

                        if (sEmpid_1.Equals(sEmpid))
                        {
                            XtraMessageBox.Show("이미 같은 사원의 데이터가 존재합니다.");
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "EMPID", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "DEPT_CD", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GRADE_CD", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GIBON", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "SIGUB", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "YJSIGUB", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColBasym, null);
                            return;
                        }
                    }
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");
                dicParams.Add("EMPID", sEmpid);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null && dt.Rows.Count != 0)
                {
                    string sDept = dt.Rows[0]["DEPT_CD"]?.ToString();
                    string sJw = dt.Rows[0]["GRADE_CD"]?.ToString();
                    string sGIBON = dt.Rows[0]["GIBON"]?.ToString();
                    string sSIGUB1 = dt.Rows[0]["SIGUB1"]?.ToString();
                    string sSIGUB2 = dt.Rows[0]["SIGUB2"]?.ToString();

                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "DEPT_CD", sDept);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GRADE_CD", sJw);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GIBON", sGIBON);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "SIGUB", sSIGUB1);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "YJSIGUB", sSIGUB2);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColBasym, sBasym);
                }
                else
                {
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "DEPT_CD", null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GRADE_CD", null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GIBON", null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "SIGUB", null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "YJSIGUB", null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColBasym, null);
                }

                SetWkTime();
            }
        }

        private void SetWkTime()
        {
            string sBasym = GridViewRetr.GetFocusedRowCellValue("BASYM")?.ToString();
            string sEmpid = GridViewRetr.GetFocusedRowCellValue("EMPID")?.ToString();

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
                    //string sJCHUL = dt.Rows[0]["JCHUL"]?.ToString();
                    string sTKGUN = dt.Rows[0]["TKGUN"]?.ToString();
                    string sDANG = dt.Rows[0]["DANG"]?.ToString();
                    string sSUK = dt.Rows[0]["SUK"]?.ToString();
                    string sCHGAM = dt.Rows[0]["CHGAM"]?.ToString();

                    double dJANUP = 0;
                    double dYAGUN = 0;
                    //double dJCHUL = 0;
                    double dTKGUN = 0;
                    double dDANG = 0;
                    double dSUK = 0;
                    double dCHGAM = 0;

                    double.TryParse(sJANUP, out dJANUP);
                    double.TryParse(sYAGUN, out dYAGUN);
                    //double.TryParse(sJCHUL, out dJCHUL);
                    double.TryParse(sTKGUN, out dTKGUN);
                    double.TryParse(sDANG, out dDANG);
                    double.TryParse(sSUK, out dSUK);
                    double.TryParse(sCHGAM, out dCHGAM);

                    double dTotJan = (dJANUP + dYAGUN + dTKGUN + dDANG + dSUK) - dCHGAM;

                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "JANTM", dTotJan);
                }
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sEDIT = GridViewRetr.GetFocusedRowCellValue("EDIT")?.ToString();

            if (string.IsNullOrEmpty(sEDIT))
            {
                GridColEmpid.OptionsColumn.AllowEdit = true;
                GridColEmpid.OptionsColumn.AllowFocus = true;
                GridColDept.OptionsColumn.AllowEdit = true;
                GridColDept.OptionsColumn.AllowFocus = true;
                GridColJw.OptionsColumn.AllowEdit = true;
                GridColJw.OptionsColumn.AllowFocus = true;
            }
            else
            {
                GridColEmpid.OptionsColumn.AllowEdit = false;
                GridColEmpid.OptionsColumn.AllowFocus = false;
                GridColDept.OptionsColumn.AllowEdit = false;
                GridColDept.OptionsColumn.AllowFocus = false;
                GridColJw.OptionsColumn.AllowEdit = false;
                GridColJw.OptionsColumn.AllowFocus = false;
            }
        }

        private void GridViewRetr_MouseDown(object sender, MouseEventArgs e)
        {
            BandedGridView view = (BandedGridView)sender;
            BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            //전체선택
            if (hitInfo.InBandPanel)
            {
                if (hitInfo.Band == null)
                    return;

                if (hitInfo.Band == gridBand6)
                {
                    string sYn = string.Empty;
                    for (int i = 0; i < view.RowCount; i++)
                    {
                        if (i == 0)
                        {
                            sYn = view.GetRowCellValue(i, GridColChk)?.ToString();
                        }

                        if (sYn.Equals("Y"))
                        {
                            view.SetRowCellValue(i, GridColChk, "N");
                        }
                        else
                        {
                            view.SetRowCellValue(i, GridColChk, "Y");
                        }
                    }

                    if (sYn.Equals("Y"))
                    {
                        view.ClearSelection();
                    }
                    else
                    {
                        view.SelectAll();
                    }
                }
            }
            else if(hitInfo.InColumnPanel)
            {
                if (hitInfo.Column == null)
                    return;

                if(hitInfo.Column == GridColChk)
                {
                    string sYn = string.Empty;
                    for (int i = 0; i < view.RowCount; i++)
                    {
                        if (i == 0)
                        {
                            sYn = view.GetRowCellValue(i, GridColChk)?.ToString();
                        }

                        if (sYn.Equals("Y"))
                        {
                            view.SetRowCellValue(i, GridColChk, "N");
                        }
                        else
                        {
                            view.SetRowCellValue(i, GridColChk, "Y");
                        }
                    }

                    if (sYn.Equals("Y"))
                    {
                        view.ClearSelection();
                    }
                    else
                    {
                        view.SelectAll();
                    }
                }
            }
        }

        private void RepoChkEdit_CheckedChanged(object sender, EventArgs e)
        {
            int iIdx = GridViewRetr.FocusedRowHandle;
            string sChk = GridViewRetr.GetRowCellValue(iIdx, "CHK")?.ToString();

            if (string.IsNullOrEmpty(sChk))
                return;

            GridViewRetr.Focus();

            if (sChk.Equals("Y"))
            {
                GridViewRetr.SetRowCellValue(iIdx, "CHK", "N");
                GridViewRetr.UnselectRow(iIdx);
            }
            else
            {
                GridViewRetr.SetRowCellValue(iIdx, "CHK", "Y");
                GridViewRetr.SelectRow(iIdx);
            }
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if(e.Column.FieldName.Equals("CSUDN1") || e.Column.FieldName.Equals("CSUDN2") || e.Column.FieldName.Equals("CSUDN3") || e.Column.FieldName.Equals("CSUDN4")
                 || e.Column.FieldName.Equals("CSUDN5") || e.Column.FieldName.Equals("CSUDN6") || e.Column.FieldName.Equals("CSUDN7"))
            {
                string sCSUDN1 = GridViewRetr.GetFocusedRowCellValue("CSUDN1")?.ToString();
                string sCSUDN2 = GridViewRetr.GetFocusedRowCellValue("CSUDN2")?.ToString();
                string sCSUDN3 = GridViewRetr.GetFocusedRowCellValue("CSUDN3")?.ToString();
                string sCSUDN4 = GridViewRetr.GetFocusedRowCellValue("CSUDN4")?.ToString();
                string sCSUDN5 = GridViewRetr.GetFocusedRowCellValue("CSUDN5")?.ToString();
                string sCSUDN6 = GridViewRetr.GetFocusedRowCellValue("CSUDN6")?.ToString();
                string sCSUDN7 = GridViewRetr.GetFocusedRowCellValue("CSUDN7")?.ToString();

                double dCSUDN1 = 0;
                double dCSUDN2 = 0;
                double dCSUDN3 = 0;
                double dCSUDN4 = 0;
                double dCSUDN5 = 0;
                double dCSUDN6 = 0;
                double dCSUDN7 = 0;

                double.TryParse(sCSUDN1, out dCSUDN1);
                double.TryParse(sCSUDN2, out dCSUDN2);
                double.TryParse(sCSUDN3, out dCSUDN3);
                double.TryParse(sCSUDN4, out dCSUDN4);
                double.TryParse(sCSUDN5, out dCSUDN5);
                double.TryParse(sCSUDN6, out dCSUDN6);
                double.TryParse(sCSUDN7, out dCSUDN7);

                double tot = dCSUDN1 + dCSUDN2 + dCSUDN3 + dCSUDN4 + dCSUDN5 + dCSUDN6 + dCSUDN7;

                GridViewRetr.SetFocusedRowCellValue("CSUDN", tot.ToString("n0"));
            }
            else if (e.Column.FieldName.Equals("CGONJ1") || e.Column.FieldName.Equals("CGONJ2") || e.Column.FieldName.Equals("CGONJ3") 
                 || e.Column.FieldName.Equals("CGONJ4") || e.Column.FieldName.Equals("CGONJ5") || e.Column.FieldName.Equals("CGONJ6"))
            {
                string sCGONJ1 = GridViewRetr.GetFocusedRowCellValue("CGONJ1")?.ToString();
                string sCGONJ2 = GridViewRetr.GetFocusedRowCellValue("CGONJ2")?.ToString();
                string sCGONJ3 = GridViewRetr.GetFocusedRowCellValue("CGONJ3")?.ToString();
                string sCGONJ4 = GridViewRetr.GetFocusedRowCellValue("CGONJ4")?.ToString();
                string sCGONJ5 = GridViewRetr.GetFocusedRowCellValue("CGONJ5")?.ToString();
                string sCGONJ6 = GridViewRetr.GetFocusedRowCellValue("CGONJ6")?.ToString();

                double dCGONJ1 = 0;
                double dCGONJ2 = 0;
                double dCGONJ3 = 0;
                double dCGONJ4 = 0;
                double dCGONJ5 = 0;
                double dCGONJ6 = 0;

                double.TryParse(sCGONJ1,out dCGONJ1);
                double.TryParse(sCGONJ2,out dCGONJ2);
                double.TryParse(sCGONJ3,out dCGONJ3);
                double.TryParse(sCGONJ4,out dCGONJ4);
                double.TryParse(sCGONJ5,out dCGONJ5);
                double.TryParse(sCGONJ6,out dCGONJ6);

                //가불금 값 변경시 잔액 비교
                if (e.Column.FieldName.Equals("CGONJ3"))
                {
                    string sEmpid = GridViewRetr.GetFocusedRowCellValue("EMPID")?.ToString();
                    string sBasym = GridViewRetr.GetFocusedRowCellValue("BASYM")?.ToString();

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Clear();
                    dicParams.Add("CMD", "CHK1");
                    dicParams.Add("EMPID", sEmpid);
                    dicParams.Add("BASYM", sBasym);

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                    if (dt != null)
                    {
                        if (double.TryParse(dt.Rows[0]["JAN"]?.ToString(), out double dJan))
                        {
                            if (dCGONJ3 > dJan)
                            {
                                XtraMessageBox.Show("가불금 잔액보다 큰 금액은 입력할 수 없습니다.");

                                dCGONJ3 = dJan;
                                GridViewRetr.SetFocusedRowCellValue("CGONJ3", dJan);
                            }
                        }
                    }
                }

                double tot = dCGONJ1 + dCGONJ2 + dCGONJ3 + dCGONJ4 + dCGONJ5 + dCGONJ6;

                GridViewRetr.SetFocusedRowCellValue("CGONJ", tot.ToString("n0"));
            }
        }

        private void SL001F00_TextChanged(object sender, EventArgs e)
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

        private void Txserch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = Dtstdt.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevMonth(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                Dtstdt.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = Dtstdt.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextMonth(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                Dtstdt.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }
    }
}