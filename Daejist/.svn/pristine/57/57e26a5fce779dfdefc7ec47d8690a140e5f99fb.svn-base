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
/*
* 
* 수정일자: 2022-12-19
* 수정자  : 정은영
* ID      : #0001
* 내용    : (현업요청)
*          1. 자금일보에서 결재 등록
*/
namespace AccAdm
{
    public partial class AC20001F00 : DevExpress.XtraEditors.XtraForm
    {
        public AC20001F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_AC20001F00";

        private void AC20001F00_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            ComnEtcFunc.SetDateFromFromValue(DateFrom, DateTo);
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            try
            {
                string sFdate = DateFrom.EditValue?.ToString().Substring(0, 10).Replace("-", "");
                string sTdate = DateTo.EditValue?.ToString().Substring(0, 10).Replace("-", "");

                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("FDATE", sFdate);
                dicParams.Add("TDATE", sTdate);
                dicParams.Add("CASH", ComnEtcFunc.CashCode);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr.DataSource = dt;

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        //bool bChk = false;
                        //for(int i = 0; i < dt.Rows.Count; i++)
                        //{
                        //    string sIDX = dt.Rows[i]["IDX"]?.ToString();

                        //    if(!string.IsNullOrEmpty(sIDX) && (sIDX.Equals("1") || sIDX.Equals("2") || sIDX.Equals("5") || sIDX.Equals("6")))
                        //    {
                        //        bChk = true;
                        //        break;
                        //    }
                        //}

                        //if (bChk)
                        //{
                        //    GridRetr.DataSource = dt;
                        //    GridRetr.Focus();
                        //}
                        GridRetr.Focus();
                    }
                    else
                    {
                        DateFrom.Focus();
                        DateFrom.SelectAll();
                    }
                }

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 리스트 오류");
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sFdate = DateFrom.EditValue?.ToString().Substring(0, 10).Replace("-", "");
            string sTdate = DateTo.EditValue?.ToString().Substring(0, 10).Replace("-", "");

            string sPath = string.Format(@"{0}\Temp_File\", Application.StartupPath);
            string sFileName = "자금일보_" + sFdate + "_" + sTdate + ".xls";
            ComnEtcFunc.ExportExcelFile(sPath, sFileName, GridRetr);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sACNAM = GridViewRetr.GetRowCellValue(e.RowHandle, "ACNAM")?.ToString();

            if (!string.IsNullOrEmpty(sACNAM) && sACNAM.Equals("소계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }
            else if (!string.IsNullOrEmpty(sACNAM) && sACNAM.Equals("합계"))
            {
                e.Appearance.BackColor = Color.LightYellow;
            }
            else if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void AC20001F00_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode  == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if(e.KeyCode == Keys.F8)
            {
                BtnExcel.PerformClick();
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave.PerformClick();
            }
        }

        //#0001
        private void BtnDoctAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            RptApplSystemP1 fm = new RptApplSystemP1();

            fm.Owner = this;
            fm.AddModifyGb = "ADD";
            if (fm.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void AC20001F00_TextChanged(object sender, EventArgs e)
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

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            int[] selectedRows = GridViewRetr.GetSelectedRows();

            SetClose(selectedRows);
        }
        private void SetClose(int[] selectedRows)
        {
            if (selectedRows.Length == 0)
            {
                XtraMessageBox.Show("변경하고자 하는 행을 선택해주세요.");
                return;
            }
            for (int i = 0; i < selectedRows.Length; i++)
            {
                string sSEQNO = GridViewRetr.GetRowCellValue(selectedRows[i], "SEQNO")?.ToString();
                if (string.IsNullOrEmpty(sSEQNO))
                {
                    XtraMessageBox.Show("선택된 행을 다시 확인해주세요(ex.현금/소계/합계 저장불가)");
                    return;
                }
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selectedRows.Length; i++)
                {
                    Cursor = Cursors.WaitCursor;

                    DataRow row = GridViewRetr.GetDataRow(selectedRows[i]);
                   // int iRowIdx = selectedRows[i];
                    string sATEXT = GridViewRetr.GetRowCellValue(selectedRows[i], "ATEXT")?.ToString();
                    string sACNAMS = GridViewRetr.GetRowCellValue(selectedRows[i], "ACNAMS")?.ToString();
                    string sSEQNO = GridViewRetr.GetRowCellValue(selectedRows[i], "SEQNO")?.ToString();
                    string sTDATE = GridViewRetr.GetRowCellValue(selectedRows[i], "TDATE")?.ToString().Substring(0,10);
                    
                    strSql.Clear();
                    strSql.AppendLine("IF EXISTS(SELECT * FROM actrans WHERE TDATE = '" + sTDATE + "' AND SEQNO ='" + sSEQNO + "' AND ATEXT ='" + sATEXT + "'   ) ");
                    strSql.AppendLine("   BEGIN                                            ");
                    strSql.AppendLine("       UPDATE actrans                              ");
                    strSql.AppendLine("           SET ACNAMS = '" + sACNAMS + "' ");
                    strSql.AppendLine("         WHERE TDATE = '" + sTDATE + "' AND SEQNO ='" + sSEQNO + "' AND ATEXT ='" + sATEXT + "' ");
                    strSql.AppendLine("   END                        ");
                    strSql.AppendLine("ELSE                          ");
                    strSql.AppendLine("    BEGIN                     ");
                    strSql.AppendLine("        INSERT INTO actrans ");
                    strSql.AppendLine("                ( TDATE, SEQNO, ACNAMS,ATEXT ");
                    strSql.AppendLine("                ) ");
                    strSql.AppendLine("    VALUES(  ");
                    strSql.AppendLine("                 '" + sTDATE + "' ");
                    strSql.AppendLine("                , '" + sSEQNO + "' ");
                    strSql.AppendLine("                , '" + sACNAMS + "' ");
                    strSql.AppendLine("                , '" + sATEXT + "' ");
                    strSql.AppendLine("                ) ");
                    strSql.AppendLine("    END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");
                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column == gridColumn6)
            {
                try
                {
                    string sXQTY_F = GridViewRetr.GetRowCellValue(e.RowHandle, gridColumn6)?.ToString();
                    string sXXX = GridViewRetr.GetRowCellValue(e.RowHandle, gridColumn11)?.ToString();
                    
                    if (sXQTY_F.Equals(sXXX))
                    {
                        GridViewRetr.UnselectRow(e.RowHandle);
                    }
                    else
                    {
                        GridViewRetr.SelectRow(e.RowHandle);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}