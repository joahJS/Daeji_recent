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

namespace AccAdm
{
    public partial class PopUpWithHoldInfo : DevExpress.XtraEditors.XtraForm
    {
        public PopUpWithHoldInfo()
        {
            InitializeComponent();
        }

        public string sIncmrIdtNo { get;  set; }
        public string sIncmrNm { get; set; }
        public DataTable dtResult { get; set; }
        public string sIncmrCd { get; set; }

        private void PopUpWithHoldInfo_Load(object sender, EventArgs e)
        {
            CboWithHodGb.SelectedIndex = 0;
            BtneIncmrCd.EditValue = sIncmrIdtNo;
            TxtIncmrNm.EditValue = sIncmrNm;

            DateEditTo.EditValue = System.DateTime.Now;
            DateEditFrom.EditValue = System.DateTime.Now.AddMonths(-3);

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sWithGb = Convert.ToString(CboWithHodGb.SelectedIndex + 1);
            string sFrYmd = Convert.ToString(DateEditFrom.EditValue).Replace("-", "").Substring(0, 8);
            string sToYmd = Convert.ToString(DateEditTo.EditValue).Replace("-", "").Substring(0, 8);
            //string sInCmrCd = Convert.ToString(BtneIncmrCd.EditValue).Replace("_", "").Trim();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT WITHTAX_GB ");
            strSql.AppendLine("      , PMNT_YMD ");
            strSql.AppendLine("      , (SELECT X.INCMR_IDT_NO FROM ACC_INCM_CD X WHERE X.INCMR_CD = A.INCMR_CD) AS INCMR_IDT_NO ");
            strSql.AppendLine("      , (SELECT X.INCMR_NM FROM ACC_INCM_CD X WHERE X.INCMR_CD = A.INCMR_CD) AS INCMR_NM ");
            strSql.AppendLine("      , EARN_GB      ");
            strSql.AppendLine("      , DUTY_YM ");
            strSql.AppendLine("      , DUTY_DCNT ");
            strSql.AppendLine("      , PMNT_TOT ");
            strSql.AppendLine("      , DDCT_AMT ");
            strSql.AppendLine("      , EARN_AMT ");
            strSql.AppendLine("      , NONTAX_EARN_AMT ");
            strSql.AppendLine("      , EARNTAX ");
            strSql.AppendLine("      , RSDNTAX ");
            strSql.AppendLine("      , SPCTAX ");
            strSql.AppendLine("      , WITHTAX_TOT ");
            strSql.AppendLine("      , NOTE ");
            strSql.AppendLine("      , INCMR_CD ");
            strSql.AppendLine("      , EVDN_KEY ");
            strSql.AppendLine("      , WITHTAX_SEQ ");
            strSql.AppendLine("   FROM ACC_WITHHOLD_TAX A ");
            strSql.AppendLine("  WHERE WITHTAX_GB = '" + sWithGb + "' ");
            strSql.AppendLine("    AND SLIP_ISSUE_YN = 'N' ");
            strSql.AppendLine("    AND PMNT_YMD >= '" + sFrYmd + "' ");
            strSql.AppendLine("    AND PMNT_YMD <= '" + sToYmd + "' ");
            strSql.AppendLine("    AND INCMR_CD = '" + sIncmrCd + "' ");
            strSql.AppendLine("    AND DEL_YN = 'N' ");

            strSql.AppendLine("  ORDER BY PMNT_YMD, WITHTAX_SEQ ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("조회된 결과가 없습니다.\r\n조회조건을 다시 확인하세요!");
                return;
            }

            GridWith.DataSource = dt;
        }

        private void BtnConFirm_Click(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            dtResult = null;
            this.DialogResult = DialogResult.Cancel;
        }

        private void CloseOk()
        {
            int[] iSelRowIdx = GridViewWith.GetSelectedRows();

            for(int i = 0; i < GridViewWith.SelectedRowsCount; i++)
            {
                DataRow dr = dtResult.NewRow();
                dtResult.Rows.Add(GridViewWith.GetRow(iSelRowIdx[i]));
            }

            this.DialogResult = DialogResult.OK;
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
                MessageBox.Show("조회 종료일자가 시작일자 보다 이전 일 수 없습니다.\r\n종료일자를 시작일자로 변환합니다.");
                DateEditTo.EditValue = DateEditFrom.EditValue;
                return;
            }
        }

        private void TxtIncmrNm_Leave(object sender, EventArgs e)
        {
            BtneIncmrCd.EditValue = "";
            BtneIncmrCd_ButtonClick(null, null);
        }

        DataRow drIncmrInfo;

        private void BtneIncmrCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataTable dtIncmrInfo = ComLib.ClsFunc.GetIncmRByIdtNo(Convert.ToString(BtneIncmrCd.EditValue).Replace("_", ""));

            if (dtIncmrInfo.Rows.Count == 1)
            {
                drIncmrInfo = dtIncmrInfo.Rows[0];
                TxtIncmrNm.EditValue = Convert.ToString(drIncmrInfo["INCMR_NM"]);
                sIncmrCd = Convert.ToString(drIncmrInfo["INCMR_CD"]);

                BtnRetr.PerformClick();
            }
            else
            {
                PopUpIncmrCd frm = new PopUpIncmrCd();
                frm.sIncmrNmByVal = TxtIncmrNm.Text.Trim();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    drIncmrInfo = frm.drResult;
                    BtneIncmrCd.EditValue = Convert.ToString(drIncmrInfo["INCMR_IDT_NO"]);
                    TxtIncmrNm.EditValue = Convert.ToString(drIncmrInfo["INCMR_NM"]);
                    sIncmrCd = Convert.ToString(drIncmrInfo["INCMR_CD"]);

                    BtnRetr.PerformClick();
                }
                else
                {
                    TxtIncmrNm.Focus();
                }
            }
        }

        private void CboWithHodGb_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void PopUpWithHoldInfo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnConFirm_Click(null, null);
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
                
            }
        }

        private void GridWith_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewWith_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}