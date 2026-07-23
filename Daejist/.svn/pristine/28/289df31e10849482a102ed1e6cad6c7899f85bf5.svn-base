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
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;

namespace AccAdm
{
    public partial class AllReprotPopUpSetting : DevExpress.XtraEditors.XtraForm
    {
        private string PROCEDURE_ID = "DP_AllReport";

        public AllReprotPopUpSetting()
        {
            InitializeComponent();
        }


        private void AllReprotPopUpSetting_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            DateTime today = DateTime.Now.Date;
            Dtyyyydt.EditValue = today.AddYears(0);
            BtnRetr.PerformClick();
        }


        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            Update();//기초재고 변경
        }

        private void Update()
        {
            try
            {
                DataTable dtKPI = (DataTable)GridRetr.DataSource;

                DataSet dsSave = ComnGridFunc.GET_DATASET_FOR_MERGE(dtKPI);
                DataTable dtMerge = dsSave.Tables[0];

                if (dtMerge == null || dtMerge.Rows.Count == 0)
                {
                    XtraMessageBox.Show("변경된 데이터가 없습니다.");
                    return;
                }

                GridColumn[] Gc = new GridColumn[] { gridColumn2 };// 추가로 여러개 넣어야함
                string errMsg = ComnGridFunc.GridViewEmptyValueCheck(bandedGridView1, Gc);

                
                if (!string.IsNullOrEmpty(errMsg))
                {
                    XtraMessageBox.Show(errMsg, "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridRetr.Focus();
                    return;
                }

                string JSON1 = ComnEtcFunc.DataTableToJsonObj(dtMerge);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "SettingUpdate");
                dicParams.Add("JSON1", JSON1);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        XtraMessageBox.Show(dt.Rows[0]["MSG"].ToString(), "저장");
                        if (dt.Rows[0]["RESULT"].ToString().Equals("1"))
                        {
                            DialogResult = DialogResult.OK;
                            //Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void AllReprotPopUpSetting_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                BtnSave.PerformClick();
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            GetRetr1();
        }
        private void GetRetr1()
        {
            string stFrom = Dtyyyydt.EditValue?.ToString().Substring(0, 4);
            //int iFrom = Convert.ToInt32(stFrom) - 1;
            //string xFrom = Convert.ToString(iFrom)?.ToString();
            string sxFrom = string.Concat(stFrom, "-01-01");

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "SettingLIST1");
            dicParams.Add("SETTINGDATE", sxFrom);

            DataTable dtC = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            
            
            if (dtC.Rows.Count == 0)
            {
                GridRetr.DataSource = null;
                //XtraMessageBox.Show("'" + stFrom + "'년도 기초재고 마감이 되어있지 않습니다. 마감 후 조회해주세요");
                return;
            }
            else
            {
                GridRetr.DataSource = dtC;
            }


        }

        private void bandedGridView1_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
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

        private void repositoryItemTextEdit1_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            double value = 0;
            double.TryParse(txt.EditValue?.ToString(), out value);

            if (value > 999999999999)
            {
                value = 999999999999;
            }

            bandedGridView1.SetFocusedRowCellValue(bandedGridView1.FocusedColumn, value);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, bandedGridView1.FocusedColumn, txt.EditValue);

            double SURDW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDW"));
            double SURDC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDC"));
            double GOCULW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULW"));
            double GOCULC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULC"));
            double TOTALW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALW"));
            double TOTALC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALC"));
            double SURDD = Math.Abs(SURDC / SURDW);

            TOTALW = SURDW + GOCULW;
            TOTALC = SURDC + GOCULC;
            double TOTALD = Math.Abs(TOTALC / TOTALW);

            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn3, SURDD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn9, TOTALD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn10, TOTALC);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn8, TOTALW);

            bandedGridView1.UpdateTotalSummary();
        }

        private void repositoryItemTextEdit4_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt1 = (TextEdit)sender;
            double value = 0;
            double.TryParse(txt1.EditValue?.ToString(), out value);

            if (value > 999999999999)
            {
                value = 999999999999;
            }

            bandedGridView1.SetFocusedRowCellValue(bandedGridView1.FocusedColumn, value);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, bandedGridView1.FocusedColumn, txt1.EditValue);

            double SURDW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDW"));
            double SURDC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDC"));
            double GOCULW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULW"));
            double GOCULC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULC"));
            double TOTALW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALW"));
            double TOTALC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALC"));
            double SURDD = Math.Abs(SURDC / SURDW);

            TOTALW = SURDW + GOCULW;
            TOTALC = SURDC + GOCULC;
            double TOTALD = Math.Abs(TOTALC / TOTALW);

            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn3, SURDD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn9, TOTALD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn10, TOTALC);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn8, TOTALW);

            bandedGridView1.UpdateTotalSummary();
        }

        private void repositoryItemTextEdit5_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            double value = 0;
            double.TryParse(txt.EditValue?.ToString(), out value);

            if (value > 999999999999)
            {
                value = 999999999999;
            }

            bandedGridView1.SetFocusedRowCellValue(bandedGridView1.FocusedColumn, value);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, bandedGridView1.FocusedColumn, txt.EditValue);

            double SURDW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDW"));
            double SURDC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDC"));
            double GOCULW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULW"));
            double GOCULC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULC"));
            double TOTALW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALW"));
            double TOTALC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALC"));
            double GOCULD = Math.Abs(GOCULC / GOCULW);

            TOTALW = SURDW + GOCULW;
            TOTALC = SURDC + GOCULC;
            double TOTALD = Math.Abs(TOTALC / TOTALW);

            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn6, GOCULD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn9, TOTALD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn10, TOTALC);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn8, TOTALW);

            bandedGridView1.UpdateTotalSummary();
        }

        private void repositoryItemTextEdit6_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt1 = (TextEdit)sender;
            double value = 0;
            double.TryParse(txt1.EditValue?.ToString(), out value);

            if (value > 999999999999)
            {
                value = 999999999999;
            }

            bandedGridView1.SetFocusedRowCellValue(bandedGridView1.FocusedColumn, value);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, bandedGridView1.FocusedColumn, txt1.EditValue);

            double SURDW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDW"));
            double SURDC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "SURDC"));
            double GOCULW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULW"));
            double GOCULC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "GOCULC"));
            double TOTALW = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALW"));
            double TOTALC = Convert.ToDouble(bandedGridView1.GetRowCellValue(bandedGridView1.FocusedRowHandle, "TOTALC"));
            double GOCULD = Math.Abs(GOCULC / GOCULW);

            TOTALW = SURDW + GOCULW;
            TOTALC = SURDC + GOCULC;
            double TOTALD = Math.Abs(TOTALC / TOTALW);

            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn6, GOCULD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn9, TOTALD);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn10, TOTALC);
            bandedGridView1.SetRowCellValue(bandedGridView1.FocusedRowHandle, gridColumn8, TOTALW);

            bandedGridView1.UpdateTotalSummary();
        }
    }
}