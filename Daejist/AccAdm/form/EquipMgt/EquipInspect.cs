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

namespace AccAdm
{
    public partial class EquipInspect : DevExpress.XtraEditors.XtraForm
    {
        public EquipInspect()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void EquipInspect_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateEditFrom.EditValue = DateTime.Now;
            DateEditTo.EditValue = DateTime.Now;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetGridRetr();
            GetValues();
        }

        private void GetGridRetr()
        {
            string sFromDt = DateEditFrom.EditValue.ToString().Replace("-","").Substring(0,8);
            string sToDt = DateEditTo.EditValue.ToString().Replace("-", "").Substring(0, 8);
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" SELECT B.MDATE ");
            strSql.AppendLine("      , B.MUSER");
            strSql.AppendLine("      , B.MUSER_ID");
            strSql.AppendLine("      , A.MAKENO ");
            strSql.AppendLine("      , A.M7_CHECK01_A ");
            strSql.AppendLine("      , A.M7_CHECK01_B ");
            strSql.AppendLine("      , A.M7_CHECK02_A ");
            strSql.AppendLine("      , A.M7_CHECK02_B ");
            strSql.AppendLine("      , A.M7_CHECK03_A ");
            strSql.AppendLine("      , A.M7_CHECK03_B ");
            strSql.AppendLine("      , A.M7_CHECK04_A ");
            strSql.AppendLine("      , A.M7_CHECK04_B ");
            strSql.AppendLine("      , A.M7_CHECK05_A ");
            strSql.AppendLine("      , A.M7_CHECK05_B ");
            strSql.AppendLine("      , A.M7_CHECK06_A ");
            strSql.AppendLine("      , A.M7_CHECK06_B ");
            strSql.AppendLine("      , A.M7_CHECK07_A ");
            strSql.AppendLine("      , A.M7_CHECK07_B ");
            strSql.AppendLine("      , A.M7_CHECK08_A ");
            strSql.AppendLine("      , A.M7_CHECK08_B ");
            strSql.AppendLine("      , A.M7_CHECK09_A ");
            strSql.AppendLine("      , A.M7_CHECK09_B ");
            strSql.AppendLine("      , A.M7_CHECK10_A ");
            strSql.AppendLine("      , A.M7_CHECK10_B ");
            strSql.AppendLine("      , A.M7_CHECK11_A ");
            strSql.AppendLine("      , A.M7_CHECK11_B ");
            strSql.AppendLine("      , A.M7_CHECK12_A ");
            strSql.AppendLine("      , A.M7_CHECK12_B ");
            strSql.AppendLine("      , A.M7_CHECK13_A ");
            strSql.AppendLine("      , A.M7_CHECK13_B ");
            strSql.AppendLine("      , A.M7_CHECK14_A ");
            strSql.AppendLine("      , A.M7_CHECK14_B ");
            strSql.AppendLine(" 	 , A.ENT_DT ");
            strSql.AppendLine(" 	 , A.ENT_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_7 A ");
            strSql.AppendLine("   LEFT JOIN MAKE_M B ON A.MAKENO=B.MAKENO ");
            strSql.AppendLine("  WHERE 1=1 ");
            strSql.AppendLine("    AND B.MDATE >=  '"+sFromDt+"' ");
            strSql.AppendLine("    AND B.MDATE <= '" + sToDt + "' ");



            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GetValues()
        {
            RadioGroup[] Rdgb = { Rdgb01, Rdgb02, Rdgb03, Rdgb04, Rdgb05, Rdgb06, Rdgb07, Rdgb08,Rdgb09, Rdgb10, Rdgb11, Rdgb12, Rdgb13, Rdgb14 };
            TextEdit[] Txt = { Txt01, Txt02, Txt03, Txt04, Txt05, Txt06, Txt07, Txt08, Txt09, Txt10, Txt11, Txt12, Txt13, Txt14 };
            for (int i = 1; i < 15; i++)
            {
                Rdgb[i - 1].EditValue = GridViewRetr.GetFocusedRowCellValue("M7_CHECK" + i.ToString("D2") + "_A");
                Txt[i - 1].EditValue = GridViewRetr.GetFocusedRowCellValue("M7_CHECK" + i.ToString("D2") + "_B");
            }
          
        }

        private void EquipInspect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                //BtnCrete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                //BtnSave_Click(null, null);
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

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void EquipInspect_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            GetValues();
        }

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}