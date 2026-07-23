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
*/
namespace AccAdm
{
    public partial class AC04001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC04001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC04001F01_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateTime today = DateTime.Now.Date;
            DateEditFrom.EditValue = today.AddDays(1 - today.Day);
            DateEditTo.EditValue = today;

            DataTable dtAcCod = GetLookUpData("2", "Y", "Y"); //전표구분
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAccod, dtAcCod, GridBal, BGridColAcntCd, "CD", "NM", "");
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { AdvBGridViewBal };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }


            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            GridBal.DataSource = GetSlipsDailyBalance(sYmdFrom, sYmdTo);
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
                string sFileNM = "일계표_" + DateTime.Now.ToLongDateString().Replace(" ", "");
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    GridBal.ExportToXls(FileName + ".xls");
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

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC04001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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

        #region[Execute By Query]

        private DataTable GetSlipsDailyBalance(string sYmdFrom, string sYmdTo)
        {
            string sSQL =
                         " select ACCOD, ACDSP, ACRDR, SUM(CAMT)CTOT, SUM(HAMT_C)CHAMT, SUM(IAMT_C)CAAMT " +
                "\r\n" + "      , SUM(DAMT)DTOT, SUM(HAMT_D)DHAMT, SUM(IAMT_D)DAAMT " +
                "\r\n" + " from  (SELECT A1.ACCOD, B1.ACDSP, B1.ACRDR, SUM(ACAMT) CAMT, SUM(ADAMT) DAMT " +
                "\r\n" + "             , SUM(case when ATGUB <> '3' then ACAMT else 0 END)HAMT_C " +
                "\r\n" + "             , SUM(case when ATGUB <> '3' then ADAMT else 0 END)HAMT_D " +
                "\r\n" + "             , SUM(case when ATGUB = '3' then ACAMT else 0 END)IAMT_C " +
                "\r\n" + "             , SUM(case when ATGUB = '3' then ADAMT else 0 END)IAMT_D " +
                "\r\n" + "        FROM   ACTRAN A1 " +
                "\r\n" + "               LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD " +
                "\r\n" + "        where  A1.TDATE between :StrDt and :EndDt " +
                "\r\n" + "        group  by A1.ACCOD, B1.ACDSP, B1.ACRDR " +

                "\r\n" + "        union  all " +
                "\r\n" + "        select '0101', '현                금', '1', SUM(ADAMT), SUM(ACAMT), SUM(ADAMT), SUM(ACAMT), 0, 0 " +
                "\r\n" + "        from   ACTRAN " +
                "\r\n" + "        where  TDATE between :StrDt and :EndDt and ATGUB <> '3' " +
                "\r\n" + "        ) A1 " +
                "\r\n" + "  group  by ACCOD, ACDSP, ACRDR " +
                "\r\n" + " order by 1 ";
            sSQL = sSQL.Replace(":StrDt", "'" + DateEditFrom.Text.Replace("-","") + "'");
            sSQL = sSQL.Replace(":EndDt", "'" + DateEditTo.Text.Replace("-", "") + "'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, sSQL); 
            return dt;
        }        

        #endregion[Execute By Query]

        #region[GetLookupData]

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
            else if (sNullYn.Equals("N"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1")) //전표구분
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_01 '");
            }
            if (sGb.Equals("2")) //계정코드
            {
                strSql.AppendLine(" SELECT ACCOD AS CD  ");
                strSql.AppendLine("      , ACNAM AS NM  ");
                strSql.AppendLine("      , ACCOD AS SEQ ");
                strSql.AppendLine("   FROM ACMSTF A  ");
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

        #endregion[GetLookupData]

        #region[GridView's Design]

        private void AdvBGridViewBal_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AdvBGridViewBal_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        #endregion[GridView's Design]

        private void AC04001F01_TextChanged(object sender, EventArgs e)
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