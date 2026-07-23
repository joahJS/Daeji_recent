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
using System.Data.SqlClient;
using DevExpress.XtraGrid.Columns;

namespace AccAdm
{
    public partial class SL002F00 : DevExpress.XtraEditors.XtraForm
    {
        public SL002F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_SL002F00";

        private void SL002F00_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);

            DateTime today = DateTime.Today;
            DateFrom.EditValue = today.ToString("yyyy") + "-01-01";
            DateTo.EditValue = today;

            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewInsa, GridViewRetr, GridViewRetr2 };
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
            SelectInsa();
            SelectDtl();
            SelectJan();
        }

        private void SelectInsa()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "INSAINFO");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    GridInsa.DataSource = dt;

                    if (dt.Rows.Count > 0)
                    {
                        GridInsa.Focus();
                    }
                    else
                    {
                        DateFrom.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewInsa_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SelectDtl();
            SelectJan();
        }

        private void SelectDtl()
        {
            try
            {
                string sDateFrom = DateFrom.EditValue?.ToString().Substring(0, 10);
                string sDateTo = DateTo.EditValue?.ToString().Substring(0, 10);
                string sEmpid = GridViewInsa.GetFocusedRowCellValue("EMP_ID")?.ToString();

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("EMPID", sEmpid);
                dicParams.Add("DATEFROM", sDateFrom);
                dicParams.Add("DATETO", sDateTo);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    GridRetr.DataSource = dt;

                    if (dt.Rows.Count > 0)
                    {
                        GridRetr.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SelectJan()
        {
            try
            {
                string sDateFrom = DateFrom.EditValue?.ToString().Substring(0, 7);
                string sDateTo = DateTo.EditValue?.ToString().Substring(0, 7);
                string sEmpid = GridViewInsa.GetFocusedRowCellValue("EMP_ID")?.ToString();

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST2");
                dicParams.Add("EMPID", sEmpid);
                dicParams.Add("DATEFROM", sDateFrom);
                dicParams.Add("DATETO", sDateTo);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    GridRetr2.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string sEmpid = GridViewInsa.GetFocusedRowCellValue("EMP_ID")?.ToString();
            string sBasdt = GridViewRetr.GetFocusedRowCellValue("BASDT")?.ToString();

            if (GridViewRetr.FocusedRowHandle >= 0 && string.IsNullOrEmpty(sBasdt))
            {
                return;
            }

            GridViewRetr.AddNewRow();
            GridViewRetr.UpdateCurrentRow();

            GridViewRetr.SetFocusedRowCellValue("EMPID", sEmpid);
            GridViewRetr.FocusedColumn = gridColumn14;
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

            for(int i= dtMerge.Rows.Count-1; i >= 0; i--)
            {
                string sBasdt = dtMerge.Rows[i]["BASDT"]?.ToString();

                if (string.IsNullOrEmpty(sBasdt))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            DataTable dtResult = SaveGABUL(dtMerge);

            if(dtResult != null && dtResult.Rows.Count > 0)
            {
                XtraMessageBox.Show(dtResult.Rows[0]["MSG"]?.ToString());

                if (double.Parse(dtResult.Rows[0]["RESULT"]?.ToString()) > 0)
                {
                    SelectDtl();
                }
            }
        }

        private DataTable SaveGABUL(DataTable dtModi)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            string json = ComnEtcFunc.DataTableToJsonObj(dtModi);

            dicParams.Clear();
            dicParams.Add("CMD", "SAVE");
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
            int[] selRows = GridViewRetr.GetSelectedRows();

            if (selRows.Length == 0)
            {
                XtraMessageBox.Show("삭제할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("선택한 항목에 대해 삭제를 진행하시겠습니까?", selRows.Length), "삭제확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            //if (XtraMessageBox.Show(string.Format("{0}의 건에 대하여 삭제를 진행하시겠습니까?", selRows.Length), "삭제확인", MessageBoxButtons.YesNo) == DialogResult.No)
            //    return;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sEMPID = string.Empty;
                string sBASDT = string.Empty;
                for (int i = 0; i < selRows.Length; i++)
                {
                    sEMPID = GridViewRetr.GetRowCellValue(selRows[i], "EMPID")?.ToString();
                    sBASDT = GridViewRetr.GetRowCellValue(selRows[i], "BASDT")?.ToString();

                    strSql.Clear();
                    strSql.AppendLine("DELETE FROM GABUL    ");
                    strSql.AppendLine(" WHERE EMPID = '" + sEMPID + "' AND  BASDT = '" + sBASDT + "'  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                SelectDtl();
                GridViewRetr.FocusedRowHandle = selRows[0] - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void RepoDateEdit_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime result))
            {
                DataTable dt = GridRetr.DataSource as DataTable;

                if (dt != null)
                {
                    string sEmpid = GridViewInsa.GetFocusedRowCellValue("EMP_ID")?.ToString();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sBasdt = dt.Rows[i]["BASDT"]?.ToString();

                        Dictionary<string, string> dicParams = new Dictionary<string, string>();

                        dicParams.Clear();
                        dicParams.Add("CMD", "CHK1");
                        dicParams.Add("BASDT", sBasdt);
                        dicParams.Add("EMPID", sEmpid);

                        DataTable dtCHK = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                        if (dtCHK != null && dtCHK.Rows.Count > 0)
                        {
                            string sCnt = dtCHK.Rows[0]["CNT"]?.ToString();

                            if (double.TryParse(sCnt, out double dCnt))
                            {
                                if(dCnt > 0)
                                {
                                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridViewRetr.FocusedColumn, null);
                                    XtraMessageBox.Show("해당일자 데이터가 이미 존재합니다.");
                                    return;
                                }
                            }
                        }
                    }
                }

                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridViewRetr.FocusedColumn, result.ToString("yyyy-MM-dd"));
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            int idx = xtraTabControl1.SelectedTabPageIndex;

            if(idx == 0)
            {
                layoutControlItem8.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem12.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else if(idx == 1)
            {
                layoutControlItem8.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem12.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sBASYM = GridViewRetr2.GetRowCellValue(e.RowHandle, "BASYM")?.ToString();
            if (!string.IsNullOrEmpty(sBASYM) && (sBASYM.Equals("[이월]")))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void SL002F00_TextChanged(object sender, EventArgs e)
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

        private void SL002F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
        }
    }
}