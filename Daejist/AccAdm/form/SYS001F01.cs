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
    public partial class SYS001F01 : DevExpress.XtraEditors.XtraForm
    {
        public SYS001F01()
        {
            InitializeComponent();
        }

        public GridView[] arrGrdView;
        private void SYS001F01_Load(object sender, EventArgs e)
        {
            int year = DateTime.Now.Year;
            DateTime firstDayOfCurYear = new DateTime(year, 1, 1);
            DateTime today = DateTime.Now;
            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);
            DateEditFrom.EditValue = firstDayOfCurYear;
            DateEditTo.EditValue = today;

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
            string sYmdFrom = string.IsNullOrEmpty(DateEditFrom.EditValue?.ToString()) ? string.Empty : DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = string.IsNullOrEmpty(DateEditTo.EditValue?.ToString()) ? string.Empty : DateEditTo.EditValue.ToString().Substring(0, 10);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("YMD_FROM", sYmdFrom);
            dicParams.Add("YMD_TO", sYmdTo);

            GridRetr.DataSource = GetFileInfo(dicParams);
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            SYS001F02 frm = new SYS001F02();
            frm.PARENT_FORM = this;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                SYS001F02 frm = new SYS001F02();
                frm.PARENT_FORM = this;
                frm.DR_FILE_INFO = GridViewRetr.GetFocusedDataRow();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr_Click(null, null);
                }
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private DataTable GetFileInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT A.VERSION_ID  ");
            strSql.AppendLine("      , '.....' AS EXEFILE");
            strSql.AppendLine("      , A.FILE_NAME ");
            strSql.AppendLine("      , A.FILE_BYTE ");
            strSql.AppendLine("      , A.UPLOAD_DT ");
            strSql.AppendLine("      , A.VERSION_RMK ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
            strSql.AppendLine("      , A.ENT_DT                                                                                                ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MOD_ID AS NUMERIC) IS NULL THEN A.MOD_ID ELSE DBO.FN_USRNM(A.MOD_ID) END AS MOD_ID");
            strSql.AppendLine("      , A.MOD_DT ");
            strSql.AppendLine("   FROM ZSYS_VERSION A ");
            strSql.AppendLine("  WHERE (('" + dicParams["YMD_FROM"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["YMD_FROM"] + "' <> '' AND A.UPLOAD_DT BETWEEN '" + dicParams["YMD_FROM"] + "' AND '" + dicParams["YMD_TO"] + "')) ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void SYS001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
            {
            }
            else if(e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void SYS001F01_TextChanged(object sender, EventArgs e)
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