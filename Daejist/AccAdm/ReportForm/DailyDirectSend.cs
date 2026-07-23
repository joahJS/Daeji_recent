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
    public partial class DailyDirectSend : DevExpress.XtraEditors.XtraForm
    {
        public DailyDirectSend()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        private void DirectSendDailyReport_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "DailyDirectSend");

            Cursor = Cursors.Default;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;

            string sYmdFrom = DateEditFrom.EditValue.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue.ToString().Replace("-", "").Substring(0, 8);

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT A.J_DATE ");
            strSql.AppendLine(" 	 , C.GUBUN1 ");
            strSql.AppendLine(" 	 , B.J_BNUM ");
            strSql.AppendLine(" 	 , D.DEALER_NM ");
            strSql.AppendLine("      , A.OWEIGHT");
            strSql.AppendLine("	     , A.DANGA AS SALEUNITPRICE ");
            strSql.AppendLine(" 	 , A.KONGKEP AS SALEPRICE ");
            strSql.AppendLine(" 	 , A.J_BOOKING ");
            strSql.AppendLine("      , G.DANGA AS PURCHUNITPRICE ");
            strSql.AppendLine("	     , G.IKONGKEP ");
            strSql.AppendLine(" 	 , IFNULL(A.DANGA, 0) - IFNULL(G.DANGA, 0) AS DIFFUNITPRICE ");
            strSql.AppendLine(" 	 , F.JUNPYOID ");
            strSql.AppendLine("   FROM INLIST A ");
            strSql.AppendLine("   LEFT OUTER JOIN IPCHULGO B ");
            strSql.AppendLine("     ON B.J_ID = A.J_ID ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE C ");
            strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
            strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD E ");
            strSql.AppendLine("     ON E.COM_NM = C.GUBUN1 ");
            strSql.AppendLine("    AND E.CD_GB = '0001' ");
            strSql.AppendLine("   LEFT OUTER JOIN MESURING F ");
            strSql.AppendLine("     ON F.IPCHULGO_MACHULID = A.J_ID ");
            strSql.AppendLine("   LEFT OUTER JOIN INLIST G ");
            strSql.AppendLine("     ON G.J_LOTNO = A.J_LOTNO ");
            strSql.AppendLine("    AND G.KERATYPE = '매입' ");
            strSql.AppendLine("        LEFT OUTER JOIN ACC_DEALER_CD H ");
            strSql.AppendLine("    	ON H.DEALER_CD = G.J_ID1 ");
            strSql.AppendLine("    	   LEFT OUTER JOIN SAWON I ");
            strSql.AppendLine("    	ON I.S_SERIAL = H.CHRG_ID ");
            strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            strSql.AppendLine("    AND F.KERATYPE = '직송' ");
            strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr.DataSource = dt;

            Cursor = Cursors.Default;
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

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void BtnExcel_Click(object sender, EventArgs e)
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
                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridRetr.ExportToXls(FileName + ".xls");
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

        private void DailyDirectSend_KeyDown(object sender, KeyEventArgs e)
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
                BtnExcel_Click(null, null);
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void gridView1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}