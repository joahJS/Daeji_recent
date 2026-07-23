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
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*            2. 레이아웃 전체 저장 설정
*            
* 수정일자 : 2023-01-13
* 수정자   : 정은영
* ID       : #0001
* 수정내용 : frm.Owner에 this 가 null로 들어감. 그래서 ProdMgtReport에 ProdStatus변수 선언해서 this 전달.
*/
namespace AccAdm
{
    public partial class ProdStatus : DevExpress.XtraEditors.XtraForm
    {
        public ProdStatus()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void ProdStatus_Load(object sender, EventArgs e)
        {
            
            ComnEtcFunc.SetDateFromToValue(DateEditFrom, DateEditTo);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { BGridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            BtnRetr_Click(null, null);
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT EMP_ID AS CD ");
                strSql.AppendLine("      , EMP_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT IDT_NO AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT J_SERIAL AS CD ");
                strSql.AppendLine("      , GUBUN1 AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY J_SERIAL) AS SEQ ");
                strSql.AppendLine("   FROM JAJAE A");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }
            else
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        public void BtnRetr_Click(object sender, EventArgs e)
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

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MDATE ");
            strSql.AppendLine("      , A.SIGN1 AS TEAM_LEADER_APRV ");
            strSql.AppendLine("      , A.SIGN1a AS TEAM_A ");
            strSql.AppendLine("      , A.SIGN2 AS TEAM_B ");
            strSql.AppendLine("      , A.SIGN3 AS TEAM_C ");
            strSql.AppendLine("      , A.SIGN4 AS DEPT_MANAGER_APRV ");
            strSql.AppendLine("      , A.SIGN5 AS REP_APRV ");
            strSql.AppendLine("      , A.MCLOSED ");
            strSql.AppendLine("      , A.MLATENESS ");
            strSql.AppendLine("      , A.MLEAVE ");
            strSql.AppendLine("      , A.MGOOUT ");
            strSql.AppendLine("      , CASE WHEN A.MCONTENT LIKE '%' + CHAR(13) + '%' AND A.MCONTENT LIKE '%' + CHAR(10) + '%' THEN REPLACE(REPLACE(A.MCONTENT, CHAR(10),'') , CHAR(13), ' / ')");
            strSql.AppendLine("        WHEN A.MCONTENT LIKE '%' + CHAR(10) + '%' THEN REPLACE(A.MCONTENT , CHAR(10), ' / ')                                                                    ");
            strSql.AppendLine("        WHEN A.MCONTENT LIKE '%' + CHAR(13) + '%' THEN REPLACE(A.MCONTENT , CHAR(13), ' / ')                                                 ");
            strSql.AppendLine("        ELSE A.MCONTENT END AS MCONTENT ");
            strSql.AppendLine("   FROM MAKE_S A ");
            strSql.AppendLine("  WHERE MDATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND MDATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND GUBUN = '1' ");
            
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count > 0)GridRetr.DataSource = dt;
            if(dt.Rows.Count == 0)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("해당 기간 동안의 검색결과가 없습니다.");
                return;
            }
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

        private void BGridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (!e.Column.FieldName.ToString().Equals("MDATE"))
            {
                return;
            }
            else if (e.Column.FieldName.ToString().Equals("MDATE"))
            {
                if(e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void ProdStatus_KeyDown(object sender, KeyEventArgs e)
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

        private void BGridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                string sWorkYmd = BGridViewRetr.GetFocusedRowCellValue("MDATE")?.ToString().Replace("-", "").Substring(0, 8);
                Cursor = Cursors.WaitCursor;

                //if(FmMainToolBar2.EmpID == "4005")
                //{
                //    ProdPlanA frm1 = new ProdPlanA();

                //    frm1._PRTFRM = this;//#0001
                //    frm1.rowUserInfo = rowUserInfo;
                //    frm1.sYmd = sWorkYmd;
                //    frm1.MdiParent = this.MdiParent;
                //    frm1.Show();
                //}
                if (FmMainToolBar2.EmpID == "4000")
                {
                    ProdPlanAA frm1a = new ProdPlanAA();

                    frm1a._PRTFRM = this;//#0001
                    frm1a.rowUserInfo = rowUserInfo;
                    frm1a.sYmd = sWorkYmd;
                    frm1a.MdiParent = this.MdiParent;
                    frm1a.Show();
                }
                //else if (FmMainToolBar2.EmpID == "3001")  //4004
                //{
                //    ProdPlanB frm2 = new ProdPlanB();

                //    frm2._PRTFRM = this;//#0001
                //    frm2.rowUserInfo = rowUserInfo;
                //    frm2.sYmd = sWorkYmd;
                //    frm2.MdiParent = this.MdiParent;
                //    frm2.Show();
                //}
                else if (FmMainToolBar2.EmpID == "4001")
                {
                    ProdPlanC frm3 = new ProdPlanC();

                    frm3._PRTFRM = this;//#0001
                    frm3.rowUserInfo = rowUserInfo;
                    frm3.sYmd = sWorkYmd;
                    frm3.MdiParent = this.MdiParent;
                    frm3.Show();
                }
                else if (FmMainToolBar2.EmpID == "3001" || FmMainToolBar2.EmpID == "2001" || FmMainToolBar2.EmpID == "1000" || FmMainToolBar2.EmpID == "2000" || FmMainToolBar2.EmpID == "2002") 
                {
                    ProdPlanAdder frm4 = new ProdPlanAdder();

                    frm4._PRTFRM = this;//#0001
                    frm4.rowUserInfo = rowUserInfo;
                    frm4.sYmd = sWorkYmd;
                    frm4.MdiParent = this.MdiParent;
                    frm4.Show();
                }
               
                //ProdPlanAdder frm = new ProdPlanAdder();

                //frm._PRTFRM = this;//#0001
                //frm.rowUserInfo = rowUserInfo;
                //frm.sYmd = sWorkYmd;
                //frm.MdiParent = this.MdiParent;
                //frm.Show();

                Cursor = Cursors.Default;
            }
        }

        private void BGridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
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
                string sFileNM = "생산현황 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = Path.GetFileNameWithoutExtension(fileDlg.FileName);
                    GridRetr.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
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

        private void ProdStatus_TextChanged(object sender, EventArgs e)
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

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}