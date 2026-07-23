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
using DevExpress.XtraGrid.Views.Grid;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*/
namespace AccAdm
{
    public partial class SYS002F01 : DevExpress.XtraEditors.XtraForm
    {
        public SYS002F01()
        {
            InitializeComponent();
        }

        public GridView[] arrGrdView;
        private void SYS002F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;
            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);

            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            //DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle("The Bezier", "Office White");

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            string sFindIdx = CboFindIdx.SelectedIndex.ToString();
            string sFindWord = TxtFindWord.EditValue?.ToString().Trim();

            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("이력일자를 입력하세요.");
                DateEditFrom.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("이력일자를 입력하세요.");
                DateEditTo.Focus();
                return;
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", sYmdFrom);
            dicParams.Add("DATE_T", sYmdTo);
            dicParams.Add("FIND_IDX", sFindIdx);
            dicParams.Add("FIND_WORD", sFindWord);

            Cursor = Cursors.WaitCursor;
            GridRetr.DataSource = GetLog(dicParams);
            Cursor = Cursors.Default;
        }

        private DataTable GetLog(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.OCCUR_DATE ");
            strSql.AppendLine("      , A.USRCD ");
            strSql.AppendLine("      , B.USRNM ");
            strSql.AppendLine("      , A.LOG_SEQ ");
            strSql.AppendLine("      , A.EDIT_KIND ");
            strSql.AppendLine("      , D.COM_NM ");
            strSql.AppendLine("      , A.PGM_ID ");
            strSql.AppendLine("      , E.PGMNM ");
            strSql.AppendLine("      , A.ACS_IP ");
            strSql.AppendLine("      , A.STD_COLS ");
            strSql.AppendLine("      , A.EDIT_RMK ");
            strSql.AppendLine("   FROM ZSYS_LOG A  ");
            strSql.AppendLine("   LEFT JOIN ZUSRLST B ");
            strSql.AppendLine("     ON A.USRCD = B.USRCD ");
            strSql.AppendLine("   LEFT JOIN COM_BASE_CD D  ");
            strSql.AppendLine("     ON A.EDIT_KIND = D.COM_CD ");
            strSql.AppendLine("    AND D.CD_GB = 'SYS00201_01' ");
            strSql.AppendLine("   LEFT JOIN ZPGMLST E ");
            strSql.AppendLine("     ON A.PGM_ID = E.PGMID ");
            strSql.AppendLine("  WHERE A.OCCUR_DATE BETWEEN '" + dicParams["DATE_F"] + " 00:00:00' AND '" + dicParams["DATE_T"] + " 23:59:59' ");
            strSql.AppendLine("    AND (( '" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ( '" + dicParams["FIND_IDX"] + "' = '0' AND B.USRNM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ( '" + dicParams["FIND_IDX"] + "' = '1' AND A.USRCD = '" + dicParams["FIND_WORD"] + "' ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ( '" + dicParams["FIND_IDX"] + "' = '2' AND D.COM_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ( '" + dicParams["FIND_IDX"] + "' = '3' AND A.ACS_IP LIKE '%" + dicParams["FIND_WORD"] + "%' )) ");
            strSql.AppendLine("  ORDER BY OCCUR_DATE DESC ");
            
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            ComnEtcFunc.ExportExcelFile(this.Text + "_", GridRetr);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void SYS002F01_TextChanged(object sender, EventArgs e)
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

        private void SYS002F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}