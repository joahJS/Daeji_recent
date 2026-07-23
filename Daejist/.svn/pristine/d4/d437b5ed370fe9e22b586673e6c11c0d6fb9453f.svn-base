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
using DevExpress.XtraBars;
using ComLib;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;

namespace AccAdm
{
    public partial class SL001F01 : DevExpress.XtraEditors.XtraForm
    {
        public SL001F01()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_SL001F01";
        

        private void SL001F01_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            ComnEtcFunc.gp_SetColorFocused(layoutControl5);

            string today1 = DateTime.Today.ToString();
            Dtstdt.EditValue = today1.Substring(0, 7);

            //사원명
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupEmpid, "HR_EMP_BASIS", "", "");
            //부서명
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupDept, "ACC_DEPT_CD", "DEPT_CD", "DEPT_NM");
            //직위
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupJw, "COM_BASE_CD", "GRADE_CD", "");


            //사원명
            ComnEtcFunc.SetBoundGridLookUp(repositoryItemGridLookUpEdit1, "HR_EMP_BASIS", "", "");
            //부서명
            ComnEtcFunc.SetBoundGridLookUp(repositoryItemGridLookUpEdit2, "ACC_DEPT_CD", "DEPT_CD", "DEPT_NM");
            //직위
            ComnEtcFunc.SetBoundGridLookUp(repositoryItemGridLookUpEdit3, "COM_BASE_CD", "GRADE_CD", "");

            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { BGridViewRetr };
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
                dicParams.Add("DATE_FROM", dt_From);
                dicParams.Add("FIND_IDX", cb_Search);
                dicParams.Add("FIND_WORD", ed_Search);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    GridRetr.DataSource = dt;

                    if (dt.Rows.Count > 0)
                    {
                        string sPAPLY = dt.Rows[0]["PAPLY"]?.ToString();

                        if (!string.IsNullOrEmpty(sPAPLY) && sPAPLY.Equals("Y"))
                        {
                            BtnApply.Text = "승인취소";
                        }
                        else if (!string.IsNullOrEmpty(sPAPLY))
                        {
                            BtnApply.Text = "승인";
                        }

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

        #region [버튼이벤트]
        public void BtnRetr_Click(object sender, EventArgs e)
        {
            //if (ComnEtcFunc.GetAuthInfo("USE", ComnEtcFunc.EXS_ID, this.Name))
            //    return;

            //전체리스트
            GetRetr2();

            //ComnEtcFunc.GetLog("RETR", ComnEtcFunc.EXS_ID, ComnEtcFunc.Client_IP, Name, Text);
        }


        private void BtnAdd_Click(object sender, EventArgs e)
        {
            //SL001F01_POP01 frm = new SL001F01_POP01();

            //frm.Owner = this;
            //frm._ModeGubun = SL001F01_POP01.ModeGubun.Add;
            //frm.Dtstdt2 = Dtstdt.EditValue?.ToString().Substring(0, 7);
            ////frm.DataRowSendEvent += new IN001F01.SendDataHandler(GetEMPID);
            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //    BtnRetr.PerformClick();
            //}

            if(BGridViewRetr.FocusedRowHandle >= 0)
            {
                string sEmpid = BGridViewRetr.GetFocusedRowCellValue("EMPID")?.ToString();
                string sBASYM = BGridViewRetr.GetFocusedRowCellValue("BASYM")?.ToString();
                string sPAYDT = BGridViewRetr.GetFocusedRowCellValue("PAYDT")?.ToString();

                if (string.IsNullOrEmpty(sEmpid))
                    return;

                BGridViewRetr.AddNewRow();
                BGridViewRetr.UpdateCurrentRow();
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColBasym, sBASYM);
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColPAYDT, sPAYDT);
                BGridViewRetr.FocusedColumn = GridColEmpid;
            }
            else
            {
                string sBASYM = Dtstdt.EditValue?.ToString().Substring(0, 7);

                BGridViewRetr.AddNewRow();
                BGridViewRetr.UpdateCurrentRow();
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColBasym, sBASYM);
                BGridViewRetr.FocusedColumn = GridColEmpid;
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

            List<Dictionary<string, string>> listParam = new List<Dictionary<string, string>>(); 

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sEmpid = dtMerge.Rows[i]["EMPID"]?.ToString();
                string sBasym = dtMerge.Rows[i]["BASYM"]?.ToString();
                string sPAPLY = dtMerge.Rows[i]["PAPLY"]?.ToString();

                if (string.IsNullOrEmpty(sEmpid))
                {
                    dtMerge.Rows[i].Delete();
                }

                if(!string.IsNullOrEmpty(sPAPLY) && sPAPLY.Equals("C"))
                {
                    Dictionary<string, string> param = new Dictionary<string, string>();

                    param.Add("BASYM", sBasym);
                    param.Add("EMPID", sEmpid);

                    listParam.Add(param);
                }
                else if(!string.IsNullOrEmpty(sPAPLY) && sPAPLY.Equals("Y"))
                {
                    XtraMessageBox.Show("승인된 데이터는 수정이 불가능 합니다.");
                    return;
                }
            }

            DataTable dtResult = SavePMONF(dtMerge);

            if (dtResult != null && dtResult.Rows.Count > 0)
            {
                XtraMessageBox.Show(dtResult.Rows[0]["MSG"]?.ToString());

                if (double.Parse(dtResult.Rows[0]["RESULT"]?.ToString()) > 0)
                {
                    for(int i=0;i< listParam.Count; i++)
                    {
                        string sBASYM = listParam[i]["BASYM"]?.ToString();
                        string sEMPID = listParam[i]["EMPID"]?.ToString();

                        string sLogRmk = "Table:PMONF -> BASYM:" + sBASYM + ",EMPID:" + sEMPID;
                        saveLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, (i + 1).ToString(), "S", this.Name, sLogRmk, ClsFunc.GetLocalIP());
                    }

                    GetRetr2();
                }
            }
        }

        private void saveLog(string sDateTime, string sUsrCd, string sLogSeq, string sEditKind, string sPgmId, string sRemark, string sIp)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("_OCCUR_DATE", sDateTime);
            dicParams.Add("_USRCD"     , sUsrCd);
            dicParams.Add("_LOG_SEQ"   , sLogSeq);
            dicParams.Add("_EDIT_KIND" , sEditKind);
            dicParams.Add("_PGM_ID"    , sPgmId);
            dicParams.Add("_ACS_IP"    , sIp);
            dicParams.Add("_EDIT_RMK", sRemark);

            DBConn.GetDataTable(DBConn.dbCon, "DP_IST_LOG", dicParams);
        }

        private DataTable SavePMONF(DataTable dtModi)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            string json = ComnEtcFunc.DataTableToJsonObj(dtModi);

            dicParams.Clear();
            dicParams.Add("CMD", "SAVE1_1");
            dicParams.Add("JSON", json);
            dicParams.Add("USER", FmMainToolBar2.drUser["USRCD"]?.ToString());

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);

            return dt;
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            DelInfo();
        }

        private void DelInfo()
        {
            //int[] selRows = BGridViewRetr.GetSelectedRows();

            DataTable dt = GridRetr.DataSource as DataTable;

            List<int> selRows = new List<int>();

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sChk = dt.Rows[i]["CHK"]?.ToString();
                    string sPAPLY = dt.Rows[i]["PAPLY"]?.ToString();

                    if (sPAPLY.Equals("Y"))
                    {
                        XtraMessageBox.Show("승인된 데이터는 삭제가 불가능 합니다.");
                        return;
                    }

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
                    string sEMPID = dt.Rows[selRows[i]]["EMPID"]?.ToString();
                    string sBASYM = dt.Rows[selRows[i]]["BASYM"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine("DELETE FROM PMONF    ");
                    strSql.AppendLine(" WHERE EMPID = '" + sEMPID + "' AND  BASYM = '" + sBASYM + "'  ");
                    strSql.AppendLine(" DECLARE @CNT NUMERIC = 0              ");
                    strSql.AppendLine("  SELECT @CNT = COUNT(*)               ");
                    strSql.AppendLine("    FROM PMONF                         ");
                    strSql.AppendLine("   WHERE BASYM = '" + sBASYM + "'             ");
                    strSql.AppendLine("                                       ");
                    strSql.AppendLine("      IF @CNT = 0                      ");
                    strSql.AppendLine("         BEGIN                         ");
                    strSql.AppendLine("               DELETE FROM PMONF_1     ");
                    strSql.AppendLine("                WHERE BASYM = '" + sBASYM + "'");
                    strSql.AppendLine("           END                         ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    string sPAPLY = dt.Rows[selRows[i]]["PAPLY"]?.ToString();

                    if (sPAPLY.Equals("C"))//승인취소건의 경우 삭제 로그 추가
                    {
                        string sLogRmk = "Table:PMONF -> BASYM:" + sBASYM + ",EMPID:" + sEMPID;
                        ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, (i+1).ToString(), "D", this.Name, sLogRmk, cmd);
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();
                BGridViewRetr.FocusedRowHandle = selRows[0] - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            string sBASYM = Dtstdt.EditValue?.ToString();

            DateTime dt = DateTime.Parse(sBASYM);

            DataTable dataTable = GridRetr.DataSource as DataTable;

            DataTable dtCopy = dataTable.Copy();
            dtCopy.Columns.Remove("CHK");
            dtCopy.Columns.Remove("PAPLY");

            gridControl1.DataSource = dtCopy;

            string sPath = string.Format(@"{0}\Temp_File\", Application.StartupPath);
            string sFileName = dt.ToString("yyyy년도MM월_") + "급여대장.xls";

            ComnEtcFunc.ExportExcelFile(sPath, sFileName, gridControl1);
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            string sPAPLY = string.Empty;

            if (BtnApply.Text.Equals("승인"))
            {
                sPAPLY = "Y";
            }
            else
            {
                sPAPLY = "C";
            }

            string sBASYM = Dtstdt.EditValue?.ToString().Substring(0, 7);

            if (string.IsNullOrEmpty(sBASYM))
                return;

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "APLY");
            dicParams.Add("BASYM", sBASYM);
            dicParams.Add("PAPLY", sPAPLY);
            dicParams.Add("USER", FmMainToolBar2.drUser["USRCD"]?.ToString());

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);

            if(dt != null && dt.Rows.Count > 0)
            {
                XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());
                if(int.Parse(dt.Rows[0]["RESULT"]?.ToString()) > 0)
                {
                    BtnRetr.PerformClick();
                }
            }
        }

        private void BtnCal_Click(object sender, EventArgs e)
        {
            SL001F01_POP02 frm = new SL001F01_POP02();

            frm.Owner = this;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }

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

        private void SL001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F2)
                BtnCal.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void GridViewRetr2_RowClick_1(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                SL001F01_POP01 frm = new SL001F01_POP01();

                frm.Owner = this;
                frm._ModeGubun = SL001F01_POP01.ModeGubun.Modi;
                frm._ModRow = BGridViewRetr.GetFocusedDataRow();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr.PerformClick();
                }
            }
        }

        private void BGridViewRetr_MouseDown(object sender, MouseEventArgs e)
        {
            BandedGridView view = (BandedGridView)sender;
            BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            //전체선택
            if (hitInfo.InBandPanel)
            {
                if (hitInfo.Band == null)
                    return;

                if (hitInfo.Band == gridBand22)
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
            else if (hitInfo.InColumnPanel)
            {
                if (hitInfo.Column == null)
                    return;

                if (hitInfo.Column == GridColChk)
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

        private void RepoMsCopy_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            string sEmpid = BGridViewRetr.GetFocusedRowCellValue("EMPID")?.ToString();
            string sBasym = BGridViewRetr.GetFocusedRowCellValue("BASYM")?.ToString();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST5");
            dicParams.Add("BASYM", sBasym);
            dicParams.Add("EMPID", sEmpid);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            if(dt != null && dt.Rows.Count > 0)
            {
                MemoryStream ms = new MemoryStream();

                RptPaySlip report = new RptPaySlip(dt);
                report.ExportToImage(ms);

                Image img = Image.FromStream(ms);

                ms.Close();

                Clipboard.SetImage(img);
            }
            else
            {
                XtraMessageBox.Show("급여 데이터가 없습니다.");
            }
        }

        private void RepoChkEdit_EditValueChanged(object sender, EventArgs e)
        {
            CheckEdit checkEdit = (CheckEdit) sender;

            string sVal = checkEdit.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
                return;

            if (sVal.Equals("Y"))
            {
                //BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColChk, "N");
                BGridViewRetr.SelectRow(BGridViewRetr.FocusedRowHandle);
                
            }
            else
            {
                //BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColChk, "Y");
                BGridViewRetr.UnselectRow(BGridViewRetr.FocusedRowHandle);
            }
        }

        private void BGridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sEDIT = BGridViewRetr.GetFocusedRowCellValue("EDIT")?.ToString();
            string sPAYDT = BGridViewRetr.GetFocusedRowCellValue("PAYDT")?.ToString();

            if (string.IsNullOrEmpty(sEDIT))
            {
                GridColEmpid.OptionsColumn.AllowEdit = true;
                GridColEmpid.OptionsColumn.AllowFocus = true;
                GridColDept.OptionsColumn.AllowEdit = true;
                GridColDept.OptionsColumn.AllowFocus = true;
                GridColJw.OptionsColumn.AllowEdit = true;
                GridColJw.OptionsColumn.AllowFocus = true;

                if (string.IsNullOrEmpty(sPAYDT))
                {
                    GridColPAYDT.OptionsColumn.AllowEdit = true;
                }
            }
            else
            {
                GridColEmpid.OptionsColumn.AllowEdit = false;
                GridColEmpid.OptionsColumn.AllowFocus = false;
                GridColDept.OptionsColumn.AllowEdit = false;
                GridColDept.OptionsColumn.AllowFocus = false;
                GridColJw.OptionsColumn.AllowEdit = false;
                GridColJw.OptionsColumn.AllowFocus = false;
                GridColPAYDT.OptionsColumn.AllowEdit = false;
            }
        }

        private void RepoGridLkupEmpid_EditValueChanged(object sender, EventArgs e)
        {
            GridLookUpEditBase lookUpEdit = (GridLookUpEditBase)sender;
            string sEmpid = lookUpEdit.EditValue?.ToString();
            string sBasym = Dtstdt.EditValue?.ToString().Substring(0, 7);
            string sPaydt = BGridViewRetr.GetRowCellValue(BGridViewRetr.FocusedRowHandle - 1, "PAYDT")?.ToString();

            if (!string.IsNullOrEmpty(sEmpid))
            {
                DataTable dtRetr = GridRetr.DataSource as DataTable;

                if (dtRetr != null)
                {
                    for (int i = 0; i < dtRetr.Rows.Count; i++)
                    {
                        string sEmpid_1 = dtRetr.Rows[i]["EMPID"]?.ToString();

                        if (sEmpid_1.Equals(sEmpid))
                        {
                            XtraMessageBox.Show("이미 같은 사원의 데이터가 존재합니다.");
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "EMPID", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "ENTDT", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PDEPT", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PJKWK", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "YAKAMT", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PSIGUB1", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PSIGUB2", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGJSD1", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PWKTM", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PYJTM", null);

                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PBKCOD", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PBKNUM", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PBKOWN", null);
                            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColBasym, null);
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
                    string sENTDT = dt.Rows[0]["ENTDT"]?.ToString();
                    string sDept = dt.Rows[0]["DEPT_CD"]?.ToString();
                    string sJw = dt.Rows[0]["GRADE_CD"]?.ToString();
                    string sYAKAMT = dt.Rows[0]["YAKAMT"]?.ToString();
                    string sSIGUB1 = dt.Rows[0]["SIGUB1"]?.ToString();
                    string sSIGUB2 = dt.Rows[0]["SIGUB2"]?.ToString();
                    string sGIBON = dt.Rows[0]["GIBON"]?.ToString();
                    string sBASTIME1 = dt.Rows[0]["BASTIME1"]?.ToString();
                    string sBASTIME2 = dt.Rows[0]["BASTIME2"]?.ToString();

                    string sACNT_HDR = dt.Rows[0]["ACNT_HDR"]?.ToString();
                    string sDEAL_BANK_CD = dt.Rows[0]["DEAL_BANK_CD"]?.ToString();
                    string sPMNT_ACNT_FST = dt.Rows[0]["PMNT_ACNT_FST"]?.ToString();

                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "ENTDT", sENTDT);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PDEPT", sDept);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PJKWK", sJw);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "YAKAMT", sYAKAMT);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PSIGUB1", sSIGUB1);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PSIGUB2", sSIGUB2);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGJSD1", sGIBON);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PWKTM", sBASTIME1);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PYJTM", sBASTIME2);

                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PBKCOD", sDEAL_BANK_CD);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PBKNUM", sPMNT_ACNT_FST);
                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PBKOWN", sACNT_HDR);

                    BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, GridColBasym, sBasym);
                }
            }
        }

        private void BGridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            string sFileName = e.Column.FieldName;

            if(sFileName.Equals("PGJSD1") || sFileName.Equals("PGJSD2") || sFileName.Equals("PGJSD3") 
                || sFileName.Equals("PGJSD4") || sFileName.Equals("PGJSD5") || sFileName.Equals("PGJSD6") || sFileName.Equals("PGJSD7"))
            {
                //고정수당
                string sPGJSD1 = BGridViewRetr.GetFocusedRowCellValue("PGJSD1")?.ToString();
                string sPGJSD2 = BGridViewRetr.GetFocusedRowCellValue("PGJSD2")?.ToString();
                string sPGJSD3 = BGridViewRetr.GetFocusedRowCellValue("PGJSD3")?.ToString();
                string sPGJSD4 = BGridViewRetr.GetFocusedRowCellValue("PGJSD4")?.ToString();
                string sPGJSD5 = BGridViewRetr.GetFocusedRowCellValue("PGJSD5")?.ToString();
                string sPGJSD6 = BGridViewRetr.GetFocusedRowCellValue("PGJSD6")?.ToString();
                string sPGJSD7 = BGridViewRetr.GetFocusedRowCellValue("PGJSD7")?.ToString();

                string sPSOGE2 = BGridViewRetr.GetFocusedRowCellValue("PSOGE2")?.ToString();

                double dPGJSD1 = 0;
                double dPGJSD2 = 0;
                double dPGJSD3 = 0;
                double dPGJSD4 = 0;
                double dPGJSD5 = 0;
                double dPGJSD6 = 0;
                double dPGJSD7 = 0;

                double dPSOGE2 = 0;

                double.TryParse(sPGJSD1, out dPGJSD1);
                double.TryParse(sPGJSD2, out dPGJSD2);
                double.TryParse(sPGJSD3, out dPGJSD3);
                double.TryParse(sPGJSD4, out dPGJSD4);
                double.TryParse(sPGJSD5, out dPGJSD5);
                double.TryParse(sPGJSD6, out dPGJSD6);
                double.TryParse(sPGJSD7, out dPGJSD7);

                double.TryParse(sPSOGE2, out dPSOGE2);

                double tot = dPGJSD1 + dPGJSD2 + dPGJSD3 + dPGJSD4 + dPGJSD5 + dPGJSD6 + dPGJSD7;
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGSOGE", tot);
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGROSS", tot+ dPSOGE2);
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGROSS2", tot + dPSOGE2);
                setPCHAIN();
            }
            else if (sFileName.Equals("PCHSD1") || sFileName.Equals("PCHSD2") || sFileName.Equals("PCHSD3") || sFileName.Equals("PCHSD4")
                 || sFileName.Equals("PCHSD5") || sFileName.Equals("PCHSD6") || sFileName.Equals("PCHSD7") || sFileName.Equals("PCHSD8"))
            {
                //변동수당
                string sPCHSD1 = BGridViewRetr.GetFocusedRowCellValue("PCHSD1")?.ToString();
                string sPCHSD2 = BGridViewRetr.GetFocusedRowCellValue("PCHSD2")?.ToString();
                string sPCHSD3 = BGridViewRetr.GetFocusedRowCellValue("PCHSD3")?.ToString();
                string sPCHSD4 = BGridViewRetr.GetFocusedRowCellValue("PCHSD4")?.ToString();
                string sPCHSD5 = BGridViewRetr.GetFocusedRowCellValue("PCHSD5")?.ToString();
                string sPCHSD6 = BGridViewRetr.GetFocusedRowCellValue("PCHSD6")?.ToString();
                string sPCHSD7 = BGridViewRetr.GetFocusedRowCellValue("PCHSD7")?.ToString();
                string sPCHSD8 = BGridViewRetr.GetFocusedRowCellValue("PCHSD8")?.ToString();

                string sPGSOGE = BGridViewRetr.GetFocusedRowCellValue("PGSOGE")?.ToString();

                double dPCHSD1 = 0;
                double dPCHSD2 = 0;
                double dPCHSD3 = 0;
                double dPCHSD4 = 0;
                double dPCHSD5 = 0;
                double dPCHSD6 = 0;
                double dPCHSD7 = 0;
                double dPCHSD8 = 0;

                double dPGSOGE = 0;

                double.TryParse(sPCHSD1, out dPCHSD1);
                double.TryParse(sPCHSD2, out dPCHSD2);
                double.TryParse(sPCHSD3, out dPCHSD3);
                double.TryParse(sPCHSD4, out dPCHSD4);
                double.TryParse(sPCHSD5, out dPCHSD5);
                double.TryParse(sPCHSD6, out dPCHSD6);
                double.TryParse(sPCHSD7, out dPCHSD7);
                double.TryParse(sPCHSD8, out dPCHSD8);

                double.TryParse(sPGSOGE, out dPGSOGE);

                double tot = dPCHSD1 + dPCHSD2 + dPCHSD3 + dPCHSD4 + dPCHSD5 + dPCHSD6 + dPCHSD7 + dPCHSD8;

                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PSOGE2", tot);
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGROSS", tot + dPGSOGE);
                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PGROSS2", tot + dPGSOGE);
                setPCHAIN();
            }
            else if (sFileName.Equals("PGJGJ1") || sFileName.Equals("PGJGJ2") || sFileName.Equals("PGJGJ3") || sFileName.Equals("PGJGJ4")
                 || sFileName.Equals("PGJGJ5") || sFileName.Equals("PGJGJ6") || sFileName.Equals("PGJGJ7")
                 || sFileName.Equals("PCHGJ1") || sFileName.Equals("PCHGJ2") || sFileName.Equals("PCHGJ3") || sFileName.Equals("PCHGJ4")
                 || sFileName.Equals("PCHGJ5") || sFileName.Equals("PCHGJ6") || sFileName.Equals("PCHGJ7"))
            {
                //고정공제
                string sPGJGJ1 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ1")?.ToString();
                string sPGJGJ2 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ2")?.ToString();
                string sPGJGJ3 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ3")?.ToString();
                string sPGJGJ4 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ4")?.ToString();
                string sPGJGJ5 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ5")?.ToString();
                string sPGJGJ6 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ6")?.ToString();
                string sPGJGJ7 = BGridViewRetr.GetFocusedRowCellValue("PGJGJ7")?.ToString();
                //변동공제
                string sPCHGJ1 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ1")?.ToString();
                string sPCHGJ2 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ2")?.ToString();
                string sPCHGJ3 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ3")?.ToString();
                string sPCHGJ4 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ4")?.ToString();
                string sPCHGJ5 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ5")?.ToString();
                string sPCHGJ6 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ6")?.ToString();
                string sPCHGJ7 = BGridViewRetr.GetFocusedRowCellValue("PCHGJ7")?.ToString();

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

                double.TryParse(sPGJGJ1, out dPGJGJ1);
                double.TryParse(sPGJGJ2, out dPGJGJ2);
                double.TryParse(sPGJGJ3, out dPGJGJ3);
                double.TryParse(sPGJGJ4, out dPGJGJ4);
                double.TryParse(sPGJGJ5, out dPGJGJ5);
                double.TryParse(sPGJGJ6, out dPGJGJ6);
                double.TryParse(sPGJGJ7, out dPGJGJ7);
                double.TryParse(sPCHGJ1, out dPCHGJ1);
                double.TryParse(sPCHGJ2, out dPCHGJ2);
                double.TryParse(sPCHGJ3, out dPCHGJ3);
                double.TryParse(sPCHGJ4, out dPCHGJ4);
                double.TryParse(sPCHGJ5, out dPCHGJ5);
                double.TryParse(sPCHGJ6, out dPCHGJ6);
                double.TryParse(sPCHGJ7, out dPCHGJ7);

                //가불금 값 변경시 잔액 비교
                if (sFileName.Equals("PCHGJ3"))
                {
                    string sEmpid = BGridViewRetr.GetFocusedRowCellValue("EMPID")?.ToString();
                    string sBasym = BGridViewRetr.GetFocusedRowCellValue("BASYM")?.ToString();

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Clear();
                    dicParams.Add("CMD", "CHK1");
                    dicParams.Add("EMPID", sEmpid);
                    dicParams.Add("BASYM", sBasym);

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                    if(dt!= null)
                    {
                        if(double.TryParse(dt.Rows[0]["JAN"]?.ToString(), out double dJan))
                        {
                            if(dPCHGJ3 > dJan)
                            {
                                XtraMessageBox.Show("가불금 잔액보다 큰 금액은 입력할 수 없습니다.");

                                dPCHGJ3 = dJan;
                                BGridViewRetr.SetFocusedRowCellValue("PCHGJ3", dJan);
                            }
                        }
                    }
                }


                double tot = dPGJGJ1 + dPGJGJ2 + dPGJGJ3 + dPGJGJ4 + dPGJGJ5 + dPGJGJ6 + dPGJGJ7
                            + dPCHGJ1 + dPCHGJ2 + dPCHGJ3 + dPCHGJ4 + dPCHGJ5 + dPCHGJ6 + dPCHGJ7;

                BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PTOTGO", tot);

                setPCHAIN();
            }
        }

        private void setPCHAIN()
        {
            string sPGROSS = BGridViewRetr.GetFocusedRowCellValue("PGROSS")?.ToString();
            string sPTOTGO = BGridViewRetr.GetFocusedRowCellValue("PTOTGO")?.ToString();

            double dPGROSS = 0;
            double dPTOTGO = 0;

            double.TryParse(sPGROSS, out dPGROSS);
            double.TryParse(sPTOTGO, out dPTOTGO);

            double dPCHAIN = dPGROSS - dPTOTGO;

            BGridViewRetr.SetRowCellValue(BGridViewRetr.FocusedRowHandle, "PCHAIN", dPCHAIN);
        }

        private void SL001F01_TextChanged(object sender, EventArgs e)
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