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
using DevExpress.XtraGrid.Views.Grid;
using System.Diagnostics;
using ComLib;
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
    public partial class AC03001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC03001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC03001F01_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateTime today = DateTime.Now.Date;
            DateEditFrom.EditValue = today.AddDays(1 - today.Day);
            DateEditTo.EditValue = today;

            DataTable dtAtGub = GetLookUpData("1", "Y", "Y"); //전표구분
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAtGub, dtAtGub, GridSlip, GridColAtGub, "CD", "NM", "");

            DataTable dtFindWord = GetLookUpData("2", "N", "N");

            ComGrid.SetLookUpEdit(LkupFindWord, dtFindWord, "CD", "CD", "Y");
            LkupFindWord.Properties.PopulateColumns();
            LkupFindWord.Properties.Columns[1].Visible = false;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewSlip };
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


            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                if (string.IsNullOrEmpty(sYmdFrom) || string.IsNullOrEmpty(sYmdTo))
                {
                    XtraMessageBox.Show("검색기간을 올바르게 설정하세요.");
                    return;
                }
                GridSlip.DataSource = GetSlipInfo(sYmdFrom, sYmdTo, ReturningByComboBoxValues(CboFindSbj.EditValue?.ToString(), LkupFindWord.EditValue?.ToString().Replace("-", "")));

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("ERROR : \r\n" + ex.ToString());
            }
            finally
            {
                Cursor = Cursors.Default;
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
                string sFileNM = "분개장리스트_" + DateTime.Now.ToLongDateString().Replace(" ", "");
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    GridSlip.ExportToXls(FileName + ".xls");
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

        #region[Execute By Query]

        private DataTable GetSlipInfo(string sYmdFrom, string sYmdTo, string sAddingQuery)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" SELECT DATE_FORMAT(A.TDATE,'%Y-%m-%d') AS TDATE ");
            //strSql.AppendLine(" 	 , A.ATGUB  ");
            //strSql.AppendLine(" 	 , A.SEQNO  ");
            //strSql.AppendLine(" 	 , A.SEQNO  ");
            //strSql.AppendLine(" 	 , A.LINNO  ");
            //strSql.AppendLine(" 	 , A.ACCOD  ");
            //strSql.AppendLine(" 	 , B.ACDSP AS ACNAM ");
            //strSql.AppendLine(" 	 , CAST(A.CVCOD AS CHAR) AS CVCOD ");
            //strSql.AppendLine(" 	 , A.CVNAM  ");
            //strSql.AppendLine(" 	 , A.ACTCD  ");
            //strSql.AppendLine(" 	 , A.ACTNM  ");
            //strSql.AppendLine(" 	 , A.ATEXT  ");
            //strSql.AppendLine(" 	 , A.ACAMT  ");
            //strSql.AppendLine(" 	 , A.ADAMT  ");
            //strSql.AppendLine(" 	 , A.ADPCD  ");
            //strSql.AppendLine(" 	 , A.ADPNM  ");
            //strSql.AppendLine(" 	 , A.AAUTO  ");
            //strSql.AppendLine(" 	 , A.ADATE  ");
            //strSql.AppendLine(" 	 , A.AUSER  ");
            //strSql.AppendLine(" 	 , CAST(A.CUSER AS CHAR) AS CUSER ");
            //strSql.AppendLine(" 	 , A.CDATE  ");
            //strSql.AppendLine(" 	 , CAST(A.MUSER AS CHAR) AS MUSER ");
            //strSql.AppendLine(" 	 , A.MDATE  ");
            //strSql.AppendLine("   FROM ACTRAN A ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            //strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            //strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C ");
            //strSql.AppendLine("     ON A.ATGUB = C.COM_CD ");
            //strSql.AppendLine("    AND C.CD_GB = 'AC01001_03' ");//계정성격
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine(sAddingQuery);
            //strSql.AppendLine("  ORDER BY A.TDATE, A.SEQNO, A.SEQNO, A.LINNO ");
            #endregion

            strSql.AppendLine("SELECT CONVERT(DATE, A.TDATE) AS TDATE");
            strSql.AppendLine("     , A.ATGUB                        ");
 	        strSql.AppendLine("     , A.SEQNO                        ");
 	        strSql.AppendLine("     , A.SEQNO                        ");
 	        strSql.AppendLine("     , A.LINNO                        ");
 	        strSql.AppendLine("     , A.ACCOD                        ");
 	        strSql.AppendLine("     , B.ACDSP AS ACNAM               ");
 	        strSql.AppendLine("     , A.CVCOD AS CVCOD               ");
 	        strSql.AppendLine("     , A.CVNAM                        ");
 	        strSql.AppendLine("     , A.ACTCD                        ");
 	        strSql.AppendLine("     , A.ACTNM                        ");
 	        strSql.AppendLine("     , A.ATEXT                        ");
 	        strSql.AppendLine("     , A.ACAMT                        ");
 	        strSql.AppendLine("     , A.ADAMT                        ");
 	        strSql.AppendLine("     , A.ADPCD                        ");
 	        strSql.AppendLine("     , A.ADPNM                        ");
 	        strSql.AppendLine("     , D.COM_NM AS AAUTO                        ");
 	        strSql.AppendLine("     , A.ADATE                        ");
            strSql.AppendLine("     , A.AUSER                        ");
 	        strSql.AppendLine("     , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER ");
 	        strSql.AppendLine("     , A.CDATE                                                                                              ");
 	        strSql.AppendLine("     , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER ");
 	        strSql.AppendLine("     , A.MDATE                                                                                              ");
            strSql.AppendLine("   FROM ACTRAN A                                                                                            ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B                                                                                 ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD                                                                                   ");
            strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C                                                                            ");
            strSql.AppendLine("     ON A.ATGUB = C.COM_CD                                                                                  ");
            strSql.AppendLine("    AND C.CD_GB = 'AC01001_03'                                                                              ");
            strSql.AppendLine("   LEFT JOIN COM_BASE_CD D    ");
            strSql.AppendLine("     ON A.AAUTO = D.COM_CD    ");
            strSql.AppendLine("    AND D.CD_GB = 'AC02001_06'");
            strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            strSql.AppendLine(sAddingQuery);
            strSql.AppendLine("  ORDER BY A.TDATE, A.SEQNO, A.LINNO                                                                        ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private string ReturningByComboBoxValues(string sValue, string sText)
        {
            if (string.IsNullOrEmpty(sText))
                return string.Empty;

            StringBuilder strSql = new StringBuilder();

            if (sValue.Equals("계정코드"))
                strSql.AppendLine("    AND A.ACCOD = '" + sText + "' ");
            else if (sValue.Equals("계정명"))
                strSql.AppendLine("    AND A.ACNAM = '" + sText + "' ");
            else if (sValue.Equals("성격"))
                strSql.AppendLine("    AND C.COM_NM = '" + sText + "'");
            else if (sValue.Equals("관계계정코드"))
                strSql.AppendLine("    AND B.ASMCD = '" + sText + "' ");
            else if (sValue.Equals("관계계정명"))
            {
                strSql.AppendLine("    AND B.ASMCD IN (SELECT X1.ACCOD ");
                strSql.AppendLine("                      FROM ACMSTF X1 ");
                strSql.AppendLine("                     WHERE X1.ACNAM = '" + sText + "' ) ");
            }

            return strSql.ToString();
        }
        #endregion[Execute By Query]

        #region[KeyDown Event]

        private void LkupFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void AC03001F01_KeyDown(object sender, KeyEventArgs e)
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

        #endregion[KeyDown Event]

        #region[GridView Row's Design]

        private void GridViewSlip_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewSlip_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridView Row's Design]
        
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
            else if(sGb.Equals("2") || sGb.Equals("5")) //계정코드 
            {
                strSql.AppendLine(" SELECT DISTINCT ACCOD AS CD");
                strSql.AppendLine("                 , '' AS NM");
                strSql.AppendLine(" FROM ACTRAN                ");
                strSql.AppendLine(" WHERE ACCOD != ''");
                strSql.AppendLine("GROUP BY ACCOD");
            }
            else if (sGb.Equals("3") || sGb.Equals("6")) //계정명
            {
                strSql.AppendLine(" SELECT DISTINCT ACNAM AS CD");
                strSql.AppendLine("                 , '' AS NM");
                strSql.AppendLine(" FROM ACTRAN                ");
                strSql.AppendLine(" WHERE ACNAM != ''");
                strSql.AppendLine("GROUP BY ACNAM");
            }
            else if (sGb.Equals("4")) //성격
            {
                strSql.AppendLine(" SELECT COM_NM AS CD ");
                strSql.AppendLine("                 , '' AS NM");
                strSql.AppendLine(" FROM COM_BASE_CD");
                strSql.AppendLine(" WHERE CD_GB = 'AC01001_03'");
                strSql.AppendLine("GROUP BY COM_NM");
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

        private void CboFindSbj_SelectedIndexChanged(object sender, EventArgs e)
        {
            int cboFindIdx = CboFindSbj.SelectedIndex;
            cboFindIdx = cboFindIdx + 2;

            DataTable dtFindWord = GetLookUpData(cboFindIdx.ToString(), "N", "N");

            ComGrid.SetLookUpEdit(LkupFindWord, dtFindWord, "CD", "CD", "Y");
            LkupFindWord.Properties.PopulateColumns();
            LkupFindWord.Properties.Columns[1].Visible = false;
        }

        private void AC03001F01_TextChanged(object sender, EventArgs e)
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