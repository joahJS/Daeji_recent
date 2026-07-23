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
using System.IO;

/*
 * 수정일자 : 2023-02-03
 * 수정자   : 정은영
 * ID       : #0001
 * 수정내용 : 이자금액 계산식 변경 (금액*이자율/365*일수). 소수점이하 버림. ex) 208,414,900*5.179%/365*53 = 1,567,320
 */
namespace AccAdm
{
    public partial class AC19001F00 : DevExpress.XtraEditors.XtraForm
    {
        public AC19001F00()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;

        private void AC19001F00_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

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

            BtnRetr.PerformClick();
        }

        #region Grid Style
        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }
        #endregion

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetAcbilData();
        }

        private void GetAcbilData()
        {
            string sDateFrom = DateFrom.EditValue?.ToString().Substring(0,10);
            string sDateTo = DateTo.EditValue?.ToString().Substring(0, 10);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("WITH TEMP1 AS(      ");
            strSql.AppendLine("    SELECT A1.BSDAT ");
            strSql.AppendLine("         , CASE WHEN A3.ACRDR = '1' THEN A2.ADAMT - A2.ACAMT");
            strSql.AppendLine("                WHEN A3.ACRDR = '2' THEN A2.ACAMT - A2.ADAMT END AS AMT");
            strSql.AppendLine("         , A1.BEDAT   ");
            strSql.AppendLine("         , A2.CVNAM   ");
            strSql.AppendLine("         , DATEDIFF(DAY, A1.BSDAT, A1.BEDAT) AS DAYS");
            strSql.AppendLine("         , REPLACE((SELECT TOP 1 value FROM STRING_SPLIT(A1.DBIGO, '*')), '%', '') AS PERSENT");
            strSql.AppendLine("           , B1.COM_NM AS BSTAT             ");
	        strSql.AppendLine("         , B2.COM_NM AS BKIND               ");
	        strSql.AppendLine("         , DBO.FN_USRNM(A1.CUSER) AS CUSER  ");
            strSql.AppendLine("         , A1.CDATE                         ");
	        strSql.AppendLine("         , DBO.FN_USRNM(A1.MUSER) AS MUSER  ");
            strSql.AppendLine("         , A1.MDATE                         ");
            strSql.AppendLine("      FROM ACBILL A1                        ");
            strSql.AppendLine("      LEFT JOIN ACTRAN A2                   ");
            strSql.AppendLine("        ON A1.TDATE = A2.TDATE              ");
            strSql.AppendLine("       AND A1.ATGUB = A2.ATGUB              ");
            strSql.AppendLine("       AND A1.SEQNO = A2.SEQNO              ");
            strSql.AppendLine("       AND A1.LINNO = A2.LINNO              ");
            strSql.AppendLine("      LEFT JOIN ACMSTF A3                   ");
            strSql.AppendLine("        ON A2.ACCOD = A3.ACCOD              ");
            strSql.AppendLine("      LEFT JOIN COM_BASE_CD B1              ");
            strSql.AppendLine("        ON A1.BSTAT = B1.COM_CD             ");
            strSql.AppendLine("       AND B1.CD_GB = 'AC02001_02'          ");
            strSql.AppendLine("      LEFT JOIN COM_BASE_CD B2              ");
            strSql.AppendLine("        ON A1.BKIND = B2.COM_CD             ");
            strSql.AppendLine("       AND B2.CD_GB = 'AC02001_03'          ");
            strSql.AppendLine("     WHERE A1.BSDAT BETWEEN '"+sDateFrom+"' AND '"+sDateTo+"' ");
            strSql.AppendLine(")                                                             ");
            strSql.AppendLine("SELECT * ");
            strSql.AppendLine("     , CASE WHEN TRY_PARSE(PERSENT AS NUMERIC) IS NOT NULL THEN FLOOR(AMT* (CONVERT(NUMERIC(5,3), PERSENT)/100)/365*DAYS) ");//#0001
            strSql.AppendLine("            ELSE NULL END AS IJA                                                                    ");
            strSql.AppendLine("  FROM TEMP1                                                                                        ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt != null)
            {
                if(dt.Rows.Count > 0)
                {
                    GridRetr.Focus();
                }
                else
                {
                    DateFrom.Focus();
                    DateFrom.SelectAll();
                }

                GridRetr.DataSource = dt;
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ComnEtcFunc.ExportExcelFile("어음리스트_" + DateTime.Today.ToString("yyyyMMdd"), GridRetr);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC19001F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
                //BtnClose.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void DateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void AC19001F00_TextChanged(object sender, EventArgs e)
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
    }
}