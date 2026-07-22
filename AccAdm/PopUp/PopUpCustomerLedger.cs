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
    public partial class PopUpCustomerLedger : DevExpress.XtraEditors.XtraForm
    {
        public PopUpCustomerLedger()
        {
            InitializeComponent();
        }

        public string sAccCd { get;  set; }
        public string sAccNm { get; set; }
        public string sDealerCd { get; set; }
        public string sDealerNm { get; set; }
        public string sIdtNo { get; set; }
        public string sFrYmd { get; set; }
        public string sToYmd { get; set; }


        private void PopUpCustomerLedger_Load(object sender, EventArgs e)
        {
            TxtAccNm.Text = sAccNm;
            TxtDealerNm.Text = sDealerNm;
            TxtIdtNo.Text = sIdtNo;
            DtpFrdate.EditValue = Convert.ToDateTime(sFrYmd);
            DtpToDate.EditValue = Convert.ToDateTime(sToYmd);

            BtnRetr.PerformClick();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            sFrYmd = Convert.ToString(DtpFrdate.EditValue).Replace("-", "").Substring(0,8);
            sToYmd = Convert.ToString(DtpToDate.EditValue).Replace("-", "").Substring(0, 8);
            string sYyDate = sFrYmd.Substring(0, 4) + "00";
            string sYyStrtDate = sToYmd.Substring(0, 4) + "0101";

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" WITH OVER_CARRY AS ");
            strSql.AppendLine(" (SELECT A.DEALER_CD ");
            strSql.AppendLine("       , SUM(A.DEBT_AMT) AS DEBT_AMT ");
            strSql.AppendLine("       , SUM(A.CRDT_AMT) AS CRDT_AMT ");
            strSql.AppendLine("   FROM ACC_OVER_CARRIED A ");
            strSql.AppendLine("  WHERE A.ACC_CD = '" + sAccCd + "'  ");
            strSql.AppendLine("    AND A.DEALER_CD = " + sDealerCd + " ");
            strSql.AppendLine("    AND A.ACCNT_YM < '" + sYyDate + "' ");
            strSql.AppendLine("  GROUP BY A.DEALER_CD ");
            strSql.AppendLine(" ) ");
            strSql.AppendLine(" , BFR_AMT AS ");
            strSql.AppendLine(" (SELECT DEALER_CD ");
            strSql.AppendLine("      , SUM(DEBT_AMT) AS DEBT_AMT ");
            strSql.AppendLine("      , SUM(CRDT_AMT) AS CRDT_AMT ");
            strSql.AppendLine("   FROM ACC_SLIP_DET A ");
            strSql.AppendLine("  WHERE A.SLIP_YMD >= '" + sYyStrtDate + "' ");
            strSql.AppendLine("    AND A.SLIP_YMD < '" + sFrYmd + "'   ");
            strSql.AppendLine("    AND A.ACC_CD = '" + sAccCd + "' ");
            strSql.AppendLine("    AND A.DEALER_CD =  " + sDealerCd + " ");
            strSql.AppendLine("  GROUP BY DEALER_CD   ");
            strSql.AppendLine(" ) ");
            strSql.AppendLine(" , BFR_INFO AS ( ");
            strSql.AppendLine(" SELECT A.DEALER_CD  ");
            strSql.AppendLine("      , A.DEALER_NM ");
            strSql.AppendLine("      , A.IDT_NO ");
            strSql.AppendLine("      , IFNULL(B.DEBT_AMT, 0) + IFNULL(C.DEBT_AMT, 0) AS CHA ");
            strSql.AppendLine("      , IFNULL(B.CRDT_AMT, 0) + IFNULL(C.CRDT_AMT, 0) AS DAE ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A  ");
            strSql.AppendLine("   LEFT JOIN OVER_CARRY B ");
            strSql.AppendLine("     ON A.DEALER_CD = B.DEALER_CD ");
            strSql.AppendLine("   LEFT JOIN BFR_AMT C  ");
            strSql.AppendLine("     ON A.DEALER_CD = C.DEALER_CD ");
            strSql.AppendLine("  WHERE A.DEALER_CD = " + sDealerCd + "     ");
            strSql.AppendLine(" ) ");
            strSql.AppendLine(" , THIS_AMT AS ");
            strSql.AppendLine(" (SELECT SLIP_YMD ");
            strSql.AppendLine("      , SLIP_NO  ");
            strSql.AppendLine("      , DEBT_AMT   ");
            strSql.AppendLine("      , CRDT_AMT ");
            strSql.AppendLine("      , B.DEALER_NM ");
            strSql.AppendLine("      , ACC_RMK ");
            strSql.AppendLine("      , MGMT_CD1 ");
            strSql.AppendLine("      , (SELECT MGMT_NM FROM ACC_MGMT_CD WHERE MGMT_DTL_CD = MGMT_CD1) AS MGMT_NM ");
            strSql.AppendLine("   FROM ACC_SLIP_DET A  ");
            strSql.AppendLine("  INNER JOIN ACC_DEALER_CD B ");
            strSql.AppendLine("     ON A.DEALER_CD = B.DEALER_CD  ");
            strSql.AppendLine("  WHERE A.SLIP_YMD >= '" + sFrYmd + "' ");
            strSql.AppendLine("    AND A.SLIP_YMD <= '" + sToYmd + "' ");
            strSql.AppendLine("    AND A.ACC_CD = '" + sAccCd + "' ");
            strSql.AppendLine("    AND A.DEALER_CD =  " + sDealerCd + " ");
            strSql.AppendLine(" ) ");
            strSql.AppendLine(" , RESULT AS ( ");
            strSql.AppendLine(" SELECT '00000000' AS SLIP_YMD ");
            strSql.AppendLine("      , '00000000' AS SLIP_NO ");
            strSql.AppendLine("      , A.DEALER_NM ");
            strSql.AppendLine("      , '전기이월' AS ACC_RMK ");
            strSql.AppendLine("      , 0 AS DEBT_AMT  ");
            strSql.AppendLine("      , 0 AS CRDT_AMT ");
            strSql.AppendLine("      , CASE WHEN B.DBCR_GB = 'D' THEN A.CHA - A.DAE ELSE A.DAE - A.CHA END AS RMD_AMT ");
            strSql.AppendLine("   FROM BFR_INFO A  ");
            strSql.AppendLine("  INNER JOIN ACC_ACC_CD B ");
            strSql.AppendLine("     ON B.ACC_CD = '" + sAccCd + "' ");
            strSql.AppendLine("  UNION ALL ");
            strSql.AppendLine(" SELECT A.SLIP_YMD ");
            strSql.AppendLine("      , A.SLIP_NO ");
            strSql.AppendLine("      , A.DEALER_NM ");
            strSql.AppendLine("      , A.ACC_RMK ");
            strSql.AppendLine("      , A.DEBT_AMT ");
            strSql.AppendLine("      , A.CRDT_AMT ");
            strSql.AppendLine("      , CASE WHEN B.DBCR_GB = 'D' THEN A.DEBT_AMT - A.CRDT_AMT ELSE A.CRDT_AMT - A.DEBT_AMT END AS RMD_AMT ");
            strSql.AppendLine("   FROM THIS_AMT A  ");
            strSql.AppendLine("  INNER JOIN ACC_ACC_CD B ");
            strSql.AppendLine("     ON B.ACC_CD = '" + sAccCd + "'  ");
            strSql.AppendLine(" ) ");
            strSql.AppendLine(" SELECT SLIP_YMD ");
            strSql.AppendLine("      , SLIP_NO ");
            strSql.AppendLine("      , DEALER_NM ");
            strSql.AppendLine("      , ACC_RMK ");
            strSql.AppendLine("      , DEBT_AMT ");
            strSql.AppendLine("      , CRDT_AMT ");
            strSql.AppendLine("      , SUM(RMD_AMT) OVER (ORDER BY SLIP_YMD, SLIP_NO) AS CUM_RMD_AMT ");
            strSql.AppendLine("   FROM RESULT  ");
            strSql.AppendLine("  ORDER BY SLIP_YMD, SLIP_NO ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridLedger.DataSource = dt;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridLedger.ExportToXls(FileName + ".xls");
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

        private void PopUpCustomerLedger_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Escape)
            //{
            //    this.Dispose();
            //}
            //else if (e.KeyCode == Keys.F1)
            //{
                
            //}
            //else if (e.KeyCode == Keys.F3)
            //{
            //    BtnConfirm_Click(null, null);
            //}
            //else if (e.KeyCode == Keys.F4)
            //{

            //}
            //else if (e.KeyCode == Keys.F5)
            //{
            //    BtnRetr_Click(null, null);
            //}
            //else if (e.KeyCode == Keys.F8)
            //{
            //    Cursor = Cursors.WaitCursor;
            //    BtnReport_Click(null, null);
            //    Cursor = Cursors.Default;
            //}
        }

        private void GridLedger_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewLedger_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}