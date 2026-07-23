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
    public partial class PopUpVatInfo : DevExpress.XtraEditors.XtraForm
    {
        public PopUpVatInfo()
        {
            InitializeComponent();
        }

        public int iTaxGb { get; set; }
        public string sIdtNo { get;  set; }
        public string sDealerNm { get; set; }
        public DataTable dtResult { get; set; }
        public string sDealerCd { get; set; }

        private void PopUpVatInfo_Load(object sender, EventArgs e)
        {   
            DateEditTo.EditValue = System.DateTime.Now;
            DateEditFrom.EditValue = DateEditTo.EditValue.ToString().Substring(0, 4) + "-" + "01-01";

            DataTable dtPurcSaleGb = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupPurcSaleGb, dtPurcSaleGb, GridTax, GridColPrSlGb, "CD", "NM", "");

            CboTaxGb.SelectedIndex = iTaxGb;
            BtneDealerCd.EditValue = sIdtNo;
            TxtDealerNm.EditValue = sDealerNm;

            BtnRetr_Click(null, null);
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '전체' AS NM ");
                strSql.AppendLine("      , '0'    AS SEQ ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'PURCH_SALE_GB'");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY CD");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sBillGb = Convert.ToString(CboTaxGb.SelectedIndex + 1);
            string sFrYmd = Convert.ToString(DateEditFrom.EditValue).Replace("-", "").Substring(0, 8);
            string sToYmd = Convert.ToString(DateEditTo.EditValue).Replace("-", "").Substring(0, 8);
            string sIdtNo = Convert.ToString(BtneDealerCd.EditValue).Replace("_", "").Trim();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT PURCH_SALE_GB ");
            strSql.AppendLine("      , BILL_ISSUE_YMD ");
            strSql.AppendLine("      , IDT_NO ");
            strSql.AppendLine("      , DEALER_NM ");
            strSql.AppendLine("      , TOT_AMT ");            
            strSql.AppendLine("      , SUPL_AMT ");
            strSql.AppendLine("      , VAT_AMT ");            
            strSql.AppendLine("      , NOTE ");
            strSql.AppendLine("      , REP_NM ");
            strSql.AppendLine("      , BIZ_NM ");
            strSql.AppendLine("      , TYPE_NM ");
            strSql.AppendLine("      , BILL_KEY ");
            strSql.AppendLine("      , DEALER_CD ");
            strSql.AppendLine("      , 0 AS CHK ");
            strSql.AppendLine("   FROM ACC_TAX_MGT ");
            strSql.AppendLine("  WHERE BILL_GB = '" + sBillGb + "' ");
            strSql.AppendLine("    AND BILL_ISSUE_YMD >= '" + sFrYmd + "' ");
            strSql.AppendLine("    AND BILL_ISSUE_YMD <= '" + sToYmd + "' ");
            //strSql.AppendLine("    AND SLIP_ISSUE_YN = 'N' ");

            if (sIdtNo.Length != 0)
            {
                strSql.AppendLine("    AND IDT_NO = '" + sIdtNo + "' ");
            }
            
            strSql.AppendLine("  ORDER BY BILL_ISSUE_YMD, BILL_SEQ ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("조회된 결과가 없습니다.\r\n조회조건을 다시 확인하세요!");
            }

            GridTax.DataSource = dt;
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
            int[] iSelRowIdx = GridViewTax.GetSelectedRows();
            DataTable dt = (DataTable)GridTax.DataSource;
            dtResult = dt.Clone();

            for (int i = 0; i < GridViewTax.SelectedRowsCount; i++)
            {
                //DataRow dr = dtResult.NewRow();
                DataRow dr = dt.Rows[iSelRowIdx[i]];
                dtResult.ImportRow(dr);
            }

            this.DialogResult = DialogResult.OK;
        }

        private void CboTaxGb_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        public DataRow drDealerInfo;

        private void BtneDealerCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataTable dtDealerInfo = ComLib.ClsFunc.GetAccDefaultInfo(Convert.ToString(BtneDealerCd.EditValue).Replace("_", ""));

            if (dtDealerInfo.Rows.Count == 1)
            {
                drDealerInfo = dtDealerInfo.Rows[0];
                BtneDealerCd.EditValue = Convert.ToString(drDealerInfo["IDT_NO"]);
                TxtDealerNm.EditValue = Convert.ToString(drDealerInfo["DEALER_NM"]);
                sDealerCd = Convert.ToString(drDealerInfo["DEALER_CD"]);

                BtnRetr.PerformClick();
            }
            else
            {
                PopUpDealerCd frm = new PopUpDealerCd();
                frm.sIdtNoByVal = Convert.ToString(BtneDealerCd.EditValue).Replace("_", "");
                frm.sDealerNmByVal = TxtDealerNm.Text.Trim();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    drDealerInfo = frm.drResult;
                    BtneDealerCd.EditValue = Convert.ToString(drDealerInfo["IDT_NO"]);
                    TxtDealerNm.EditValue = Convert.ToString(drDealerInfo["DEALER_NM"]);
                    sDealerCd = Convert.ToString(drDealerInfo["DEALER_CD"]);

                    BtnRetr.PerformClick();
                }
                else
                {
                    TxtDealerNm.Focus();
                }
            }
        }

        private void TxtDealerNm_Leave(object sender, EventArgs e)
        {
            BtneDealerCd.EditValue = "";
            BtneDealerCd_ButtonClick(null, null);
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

        private void PopUpVatInfo_KeyDown(object sender, KeyEventArgs e)
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

        private void GridTax_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewTax_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}