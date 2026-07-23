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
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class ChagmReasonMgt : DevExpress.XtraEditors.XtraForm
    {
        public ChagmReasonMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void ChagmReasonMgt_Load(object sender, EventArgs e)
        {
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            arrGrdView = new GridView[] { GridViewChaResn };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("");
            strSql.AppendLine(" SELECT REASON    ");
            strSql.AppendLine("      , REASON_NO ");
            strSql.AppendLine("	 FROM CHAGAMREASON");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count < 1)
            {
                MessageBox.Show("조회결과가 없습니다");
            }
            else
            {
                GridChaResn.DataSource = dt;
            }

        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sText = GridViewChaResn.GetRowCellValue(GridViewChaResn.RowCount - 1, "REASON")?.ToString();

            if (string.IsNullOrEmpty(sText))
                return;

            GridViewChaResn.AddNewRow();
            GridViewChaResn.Focus();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            DataTable dt = (DataTable)GridChaResn.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            string sReason = string.Empty;
            int sReasonNo = 0;

            StringBuilder strSql = new StringBuilder();
            if (dtMerge.Rows.Count > 0)
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        if (dtMerge.Rows[j]["REASON"] is DBNull)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("반품사유를 입력하세요");
                            return;
                        }

                        if (dtMerge.Rows[j]["REASON_NO"] is DBNull)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" SELECT MAX(A.REASON_NO) AS MAX_VALUE ");
                            strSql.AppendLine("   FROM CHAGAMREASON A ");

                            DataTable dtMaxValue = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            string sMaxValue = dtMaxValue.Rows[0]["MAX_VALUE"]?.ToString();
                            sReasonNo = string.IsNullOrEmpty(sMaxValue) ? 1 : Convert.ToInt32(sMaxValue) + 1;
                        }
                        else
                        {
                            sReasonNo = Convert.ToInt32(dtMerge.Rows[j]["REASON_NO"]);
                        }

                        sReason = Convert.ToString(dtMerge.Rows[j]["REASON"]);

                        strSql.Clear();

                        strSql.AppendLine(" MERGE INTO CHAGAMREASON AS a    ");
                        strSql.AppendLine("  USING(SELECT                   ");
                        strSql.AppendLine("                                 ");
                        strSql.AppendLine("     REASON_NO = '" + sReasonNo + "'              ");
                        strSql.AppendLine("     , REASON = '" + sReason + "'               ");
                        strSql.AppendLine("     ) AS b                      ");
                        strSql.AppendLine("  ON a.REASON_NO = b.REASON_NO   ");
                        strSql.AppendLine("  WHEN MATCHED THEN UPDATE SET   ");
                        strSql.AppendLine("     REASON_NO = '" + sReasonNo + "'              ");
                        strSql.AppendLine("     , REASON = '" + sReason + "'               ");
                        strSql.AppendLine(" WHEN NOT MATCHED THEN INSERT(   ");
                        strSql.AppendLine("     REASON_NO                   ");
                        strSql.AppendLine("     , REASON)                   ");
                        strSql.AppendLine("  VALUES(                        ");
                        strSql.AppendLine("     '" + sReasonNo + "'                          ");
                        strSql.AppendLine("     , '" + sReason + "'                        ");
                        strSql.AppendLine(" );                              ");

                        /*
                        strSql.AppendLine("");
                        strSql.AppendLine("INSERT INTO CHAGAMREASON ");
                        strSql.AppendLine("           (REASON_NO  ");
                        strSql.AppendLine("           ,REASON     ");
                        strSql.AppendLine("           )     ");
                        strSql.AppendLine("     VALUES (");
                        strSql.AppendLine("            " + sReasonNo + "");
                        strSql.AppendLine(" 		  ,'" + sReason + "'     ");
                        strSql.AppendLine("            )                    ");
                        strSql.AppendLine("         ON DUPLICATE KEY UPDATE              ");
                        strSql.AppendLine(" 	       REASON            = '" + sReason + "'   ");
                        */

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    BtnRetr_Click(null, null);
                    GridViewChaResn.FocusedRowHandle = GridViewChaResn.LocateByDisplayText(0, GridColReason, dtMerge.Rows[0]["REASON"]?.ToString());

                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                XtraMessageBox.Show("변경된 데이터가 없습니다.");
            }

        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if ((MessageBox.Show(this, "삭제 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
            {
                return;
            }

            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sReason_no = string.Empty;
            sReason_no = GridViewChaResn.GetFocusedRowCellValue("REASON_NO").ToString();

            StringBuilder strSql = new StringBuilder();

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            try
            {
                strSql.Clear();
                strSql.AppendLine("DELETE FROM CHAGAMREASON ");
                strSql.AppendLine(" WHERE REASON_NO =" + sReason_no + "");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewChaResn.FocusedRowHandle;
                BtnRetr_Click(null, null);
                GridViewChaResn.FocusedRowHandle = idx - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ChagmReasonMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void ChagmReasonMgt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void GridViewChaResn_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewChaResn.UpdateCurrentRow();
        }

        private void GridViewChaResn_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }
    }
}