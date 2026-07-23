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
using Popbill.Fax;
using Popbill;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Net;
using System.IO;
using DevExpress.XtraGrid;
using DevExpress.XtraReports.UI;
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
    public partial class AccMeasureDev : DevExpress.XtraEditors.XtraForm
    {
        public AccMeasureDev()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        
        private void AccMeasureDev_Load(object sender, EventArgs e)
        {
            DateEditYmd.EditValue = DateTime.Today;
            
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            SetGridLookupEdit(GLkupEditCate, "N", "1");
            GLkupEditCate.Properties.View.PopulateColumns(GLkupEditCate.Properties.DataSource);
            GLkupEditCate.Properties.View.Columns[GLkupEditCate.Properties.ValueMember].Visible = false;

            arrGrdView = new GridView[] { GridViewRetr };
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
        }

        private void AccMeasureDev_TextChanged(object sender, EventArgs e)
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

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ClearFps();

            string sYmd = DateEditYmd.EditValue?.ToString();
            if (string.IsNullOrEmpty(sYmd))
            {
                XtraMessageBox.Show("일자를 지정하세요.");
                DateEditYmd.SelectAll();
                DateEditYmd.Focus();
                return;
            }
            DateTime dtDate = DateTime.Parse(sYmd);
            sYmd = dtDate.ToString("yyyy-MM-dd");

            /*
             * 전체 : ALL
             * 입고 :IN
             * 출고 : OUT
             */
            string sKeraType = RdgbKeraType.EditValue?.ToString();

            string sBigCate = string.Empty;
            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.ValueMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0)
            {
                sBigCate = sBigCate.Substring(0, sBigCate.Length - 1);
            }
            else
            {
                sBigCate = null;
            }
            
            string strSql = string.Empty;
            
            strSql = strSql + "\r\n" + " SELECT A.SUN";
            strSql = strSql + "\r\n" + "      , A.JUNPYOID ";
            strSql = strSql + "\r\n" + "      , A.KERATYPE ";
            strSql = strSql + "\r\n" + "      , A.J_BNUM ";
            strSql = strSql + "\r\n" + "      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS J_COMPANY ";
            strSql = strSql + "\r\n" + "      , A.GUBUN1 ";
            strSql = strSql + "\r\n" + "      , CASE WHEN A.OWEIGHT = 0 THEN A.IWEIGHT ELSE A.OWEIGHT END AS OWEIGHT ";
            strSql = strSql + "\r\n" + "      , A.FIRSTWEIGHT ";
            //strSql = strSql + "\r\n" + "      , DATE_FORMAT(A.FIRSTTIME, '%H : %i') AS FIRSTTIME ";
            strSql = strSql + "\r\n" + "      , FORMAT(CONVERT(DATETIME, A.FIRSTTIME),'HH:mm') AS FIRSTTIME ";
            strSql = strSql + "\r\n" + "      , A.SECONDWEIGHT ";
            //strSql = strSql + "\r\n" + "      , DATE_FORMAT(A.SECONDTIME, '%H : %i') AS SECONDTIME ";
            strSql = strSql + "\r\n" + "      , FORMAT(CONVERT(DATETIME, A.SECONDTIME),'HH:mm') AS SECONDTIME ";
            strSql = strSql + "\r\n" + " 	  , A.WEIGHT   ";
            strSql = strSql + "\r\n" + " 	  , CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END AS OCHAGAM ";
            strSql = strSql + "\r\n" + " 	  , A.J_STATE ";
            strSql = strSql + "\r\n" + " 	  , A.GUMSUBIGO	 ";
            strSql = strSql + "\r\n" + "      , B.EMP_NM ";
            strSql = strSql + "\r\n" + "      , C.WEB_FAX_YN ";
            strSql = strSql + "\r\n" + "      , C.FAX ";
            strSql = strSql + "\r\n" + "   FROM MESURING A ";
            strSql = strSql + "\r\n" + "   LEFT OUTER JOIN HR_EMP_BASIS B ";
            strSql = strSql + "\r\n" + "     ON A.GUMSU_SERIAL = B.EMP_ID ";
            strSql = strSql + "\r\n" + "   LEFT JOIN ACC_DEALER_CD C  ";
            strSql = strSql + "\r\n" + "     ON C.DEALER_CD = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHERID ELSE A.J_ASSIGNID END ";
            strSql = strSql + "\r\n" + "   LEFT JOIN JAJAE D ";
            strSql = strSql + "\r\n" + "     ON A.J_SERIAL = D.J_SERIAL ";
            strSql = strSql + "\r\n" + "  WHERE A.J_DATE = '" + sYmd + "' ";
            strSql = strSql + "\r\n" + "    AND A.KERATYPE <> '직송' ";
            strSql = strSql + "\r\n" + "    AND (('" + sKeraType + "' = 'ALL' AND 1 = 1) ";
            strSql = strSql + "\r\n" + "         OR ";
            strSql = strSql + "\r\n" + "         ('" + sKeraType + "' = 'IN' AND A.KERATYPE = '입고' )";
            strSql = strSql + "\r\n" + "         OR ";
            strSql = strSql + "\r\n" + "         ('" + sKeraType + "' = 'OUT' AND A.KERATYPE = '출고' ))";
            if (!string.IsNullOrEmpty(sBigCate))
            {
                strSql = strSql + "\r\n" + "    AND D.DAEGUBUN IN (" + sBigCate + ") ";
            }
            strSql = strSql + "\r\n" + "  ORDER BY A.SUN";

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql);
            GridRetr.DataSource = dt;
        }

        private void SetGridLookupEdit(DevExpress.XtraEditors.GridLookUpEdit gLkup, string sNullYn, string sGb)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '전체' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }
            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT DAEGUBUN AS CD ");
                strSql.AppendLine("      , DAEGUBUN AS NM ");
                strSql.AppendLine("   FROM JAJAE ");
                strSql.AppendLine("  WHERE DAEGUBUN IN ('고철A', '고철B', '슈레더') ");
                strSql.AppendLine("  GROUP BY DAEGUBUN  ");
            }
            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gLkup.Properties.DataSource = dt;
            gLkup.Properties.DisplayMember = "NM";
            gLkup.Properties.ValueMember = "CD";
        }

        private void DtpRetr_ValueChanged(object sender, EventArgs e)
        {
            ClearFps();
        }

        private void ClearFps()
        {
            GridRetr.DataSource = null;
        }

        private void AccMeasureDev_KeyDown(object sender, KeyEventArgs e)
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
                BtnExcel.PerformClick();
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccMeasureDev_FormClosed(object sender, FormClosedEventArgs e)
        {
            
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
                string sFileNM = "일일계근 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
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

        private void BtnMesuringPrint_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                
                string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                #region MariaDB
                //strSql.AppendLine(" ");
                //strSql.AppendLine(" SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM ");
                //strSql.AppendLine(" 	 , CAST(DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS CHAR) AS J_DATE");
                //strSql.AppendLine(" 	 , A.KERATYPE  ");
                //strSql.AppendLine(" 	 , A.SUN ");
                //strSql.AppendLine(" 	 , A.GUBUN1 ");
                //strSql.AppendLine(" 	 , A.J_BNUM ");
                //strSql.AppendLine(" 	 , B.EMP_NM ");
                ////strSql.AppendLine("      , CONCAT(IFNULL(CHRG_RGN_NO, ''), REPLACE(FAX, '-', '')) AS FAX_NO ");
                //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H:%i') ELSE DATE_FORMAT(A.FIRSTTIME, '%H:%i') END AS FIRSTTIME ");
                //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H:%i') ELSE DATE_FORMAT(A.SECONDTIME, '%H:%i') END AS SECONDTIME ");
                //strSql.AppendLine("      , CONCAT(FORMAT(A.SECONDWEIGHT, 0), ' KG') AS TOT_WEIGHT   ");
                //strSql.AppendLine("      , CONCAT(FORMAT(A.FIRSTWEIGHT, 0), ' KG') AS EMPTY_WEIGHT  ");
                //strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, 0) ELSE FORMAT(A.OWEIGHT, 0) END, ' KG') AS ACTL_WEIGHT  ");
                //strSql.AppendLine("      , A.GUBUN1 AS GUBUN2  ");
                //strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, 0) ELSE FORMAT(A.OCHAGAM, 0) END, ' KG') AS LOSS ");
                //strSql.AppendLine(" 	 , A.J_STATE ");
                //strSql.AppendLine("   FROM MESURING A ");
                //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                //strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID ");
                //strSql.AppendLine("  WHERE A.JUNPYOID = '" + sJunpyoId + "' ");
                //strSql.AppendLine("  ORDER BY J_DATE DESC ");
                #endregion

                strSql.AppendLine(" SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM ");
                strSql.AppendLine("      , A.J_DATE AS J_DATE                                                              ");
 	            strSql.AppendLine("      , A.KERATYPE                                                                      ");
 	            strSql.AppendLine("      , A.SUN                                                                           ");
 	            strSql.AppendLine("      , A.GUBUN1                                                                        ");
 	            strSql.AppendLine("      , A.J_BNUM                                                                        ");
                strSql.AppendLine("      , B.EMP_NM                                                                        ");
                strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(CONVERT(DATETIME, A.SECONDTIME), 'HH:mm') ELSE FORMAT(CONVERT(DATETIME, A.FIRSTTIME), 'HH:mm') END AS FIRSTTIME ");
                strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(CONVERT(DATETIME, A.FIRSTTIME), 'HH:mm') ELSE FORMAT(CONVERT(DATETIME, A.SECONDTIME), 'HH:mm') END AS SECONDTIME");
                strSql.AppendLine("      , CONCAT(FORMAT(A.SECONDWEIGHT, '#,0'), ' KG') AS TOT_WEIGHT                                                                                                ");
                strSql.AppendLine("      , CONCAT(FORMAT(A.FIRSTWEIGHT, '#,0'), ' KG') AS EMPTY_WEIGHT                                                                                               ");
                strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, '#,0') ELSE FORMAT(A.OWEIGHT, '#,0') END, ' KG') AS ACTL_WEIGHT                               ");
                strSql.AppendLine("      , A.GUBUN1 AS GUBUN2                                                                                                                                        ");
                strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, '#,0') ELSE FORMAT(A.OCHAGAM, '#,0') END, ' KG') AS LOSS                                      ");
                strSql.AppendLine("       , A.J_STATE                                                                                                                                                ");
                strSql.AppendLine("    FROM MESURING A                                                                                                                                               ");
                strSql.AppendLine("    LEFT OUTER JOIN HR_EMP_BASIS B                                                                                                                                ");
                strSql.AppendLine("      ON A.GUMSU_SERIAL = B.EMP_ID                                                                                                                                ");
                strSql.AppendLine("   WHERE A.JUNPYOID = '"+ sJunpyoId + "'                                                                                                                                    ");
                strSql.AppendLine("   ORDER BY J_DATE DESC                                                                                                                                           ");
                                                                                                                                                                                                    
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt.Rows.Count == 0)
                {
                    XtraMessageBox.Show("전표ID : " + sJunpyoId + "\r\n해당 데이터가 Mesuring에 존재하지 않습니다.");
                    return;
                }

                Dictionary<string, Image> Image = GetImage(sJunpyoId, dt.Rows[0]["J_DATE"]?.ToString());

                Dictionary<string, Image> result = new Dictionary<string, Image>();
                result.Add("1", Image["1_1"]);
                result.Add("2", Image["1_2"]);
                result.Add("3", Image["1_3"]);
                result.Add("4", Image["2_1"]);
                result.Add("5", Image["2_2"]);
                result.Add("6", Image["2_3"]);

                if (result["1"] == null)
                    result["1"] = (Image)Properties.Resources.No_Img;
                if (result["2"] == null)
                    result["2"] = (Image)Properties.Resources.No_Img;
                if (result["3"] == null)
                    result["3"] = (Image)Properties.Resources.No_Img;
                if (result["4"] == null)
                    result["4"] = (Image)Properties.Resources.No_Img;
                if (result["5"] == null)
                    result["5"] = (Image)Properties.Resources.No_Img;
                if (result["6"] == null)
                    result["6"] = (Image)Properties.Resources.No_Img;

                Cursor = Cursors.Default;

                //ReportViewer fm = new ReportViewer(dt, "RptMesuring2", result);
                //fm.ShowDialog();

                RptMesuring2 report = new RptMesuring2(dt.Rows[0], result);
                ReportPrintTool printTool = new ReportPrintTool(report);
                report.CreateDocument();
                printTool.Print();
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnWebFex_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT A.JUNPYOID                                                                     ");
                strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM");
                strSql.AppendLine("      , A.J_DATE J_DATE                                                                ");
                strSql.AppendLine("      , A.KERATYPE                                                                     ");
                strSql.AppendLine("      , A.SUN                                                                          ");
                strSql.AppendLine("      , A.GUBUN1                                                                       ");
                strSql.AppendLine("      , A.J_BNUM                                                                       ");
                strSql.AppendLine("      , A.GUMSU_SERIAL                                                                 ");
                strSql.AppendLine("      , B.EMP_NM                                                                       ");
                strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(CONVERT(DATETIME, A.SECONDTIME), 'HH:mm') ELSE FORMAT(CONVERT(DATETIME, A.FIRSTTIME), 'HH:mm') END AS FIRSTTIME ");
                strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(CONVERT(DATETIME, A.FIRSTTIME), 'HH:mm') ELSE FORMAT(CONVERT(DATETIME, A.SECONDTIME), 'HH:mm') END AS SECONDTIME");
                strSql.AppendLine("      , CONCAT(FORMAT(A.SECONDWEIGHT, '#,0'), ' KG') AS TOT_WEIGHT  ");
                strSql.AppendLine("      , CONCAT(FORMAT(A.FIRSTWEIGHT, '#,0'), ' KG') AS EMPTY_WEIGHT ");
                strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, '#,0') ELSE FORMAT(A.OWEIGHT, '#,0') END, ' KG') AS ACTL_WEIGHT");
                strSql.AppendLine("      , A.GUBUN1 AS GUBUN2                                                                                                         ");
                strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, '#,0') ELSE FORMAT(A.OCHAGAM, '#,0') END, ' KG') AS LOSS       ");
                strSql.AppendLine("      , A.J_STATE                                                                                                                  ");
                strSql.AppendLine("      , CONCAT(ISNULL(C.CHRG_RGN_NO, ''), REPLACE(C.FAX, '-', '')) AS FAX                                                          ");
                strSql.AppendLine("   FROM MESURING A                                                                                                                 ");
                strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B                                                                                                  ");
                strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID                                                                                                  ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C                                                                                                 ");
                strSql.AppendLine("       ON C.DEALER_NM = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END                                         ");
                strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + "                                                                                             ");
                strSql.AppendLine("  ORDER BY J_DATE DESC                                                                                                             ");

                #region MariaDB
                //strSql.AppendLine(" SELECT A.JUNPYOID ");
                //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM  ");
                //strSql.AppendLine("      , CAST(DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS CHAR) AS J_DATE ");
                //strSql.AppendLine("      , A.KERATYPE   ");
                //strSql.AppendLine("      , A.SUN  ");
                //strSql.AppendLine("      , A.GUBUN1  ");
                //strSql.AppendLine("      , A.J_BNUM  ");
                //strSql.AppendLine("      , A.GUMSU_SERIAL ");
                //strSql.AppendLine("      , B.EMP_NM  ");
                //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H:%i') ELSE DATE_FORMAT(A.FIRSTTIME, '%H:%i') END AS FIRSTTIME ");
                //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H:%i') ELSE DATE_FORMAT(A.SECONDTIME, '%H:%i') END AS SECONDTIME ");
                //strSql.AppendLine("      , CONCAT(FORMAT(A.SECONDWEIGHT, 0), ' KG') AS TOT_WEIGHT   ");
                //strSql.AppendLine("      , CONCAT(FORMAT(A.FIRSTWEIGHT, 0), ' KG') AS EMPTY_WEIGHT  ");
                //strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, 0) ELSE FORMAT(A.OWEIGHT, 0) END, ' KG') AS ACTL_WEIGHT  ");
                //strSql.AppendLine("      , A.GUBUN1 AS GUBUN2  ");
                //strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, 0) ELSE FORMAT(A.OCHAGAM, 0) END, ' KG') AS LOSS ");
                ////strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%m-%d %H:%i') ELSE DATE_FORMAT(A.FIRSTTIME, '%m-%d %H:%i') END AS FIRSTTIME ");
                ////strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%m-%d %H:%i') ELSE DATE_FORMAT(A.SECONDTIME, '%m-%d %H:%i') END AS SECONDTIME ");
                ////strSql.AppendLine("      , FORMAT(A.SECONDWEIGHT, 0) AS TOT_WEIGHT   ");
                ////strSql.AppendLine("      , FORMAT(A.FIRSTWEIGHT, 0) AS EMPTY_WEIGHT  ");
                ////strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, 0) ELSE FORMAT(A.OWEIGHT, 0) END AS ACTL_WEIGHT  ");
                ////strSql.AppendLine("      , A.GUBUN1 AS GUBUN2  ");
                ////strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, 0) ELSE FORMAT(A.OCHAGAM, 0) END AS LOSS ");
                //strSql.AppendLine("      , A.J_STATE  ");
                //strSql.AppendLine("      , CONCAT(IFNULL(C.CHRG_RGN_NO, ''), REPLACE(C.FAX, '-', '')) AS FAX ");
                //strSql.AppendLine("   FROM MESURING A  ");
                //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B  ");
                //strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID  ");
                //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C  ");
                //strSql.AppendLine("  	 ON C.DEALER_NM = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END ");
                //strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");
                //strSql.AppendLine("  ORDER BY J_DATE DESC  ");
                #endregion

                #region[이전 쿼리 2020-12-24]

                //strSql.AppendLine(" SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM ");
                //strSql.AppendLine(" 	 , CAST(DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS CHAR) AS J_DATE");
                //strSql.AppendLine(" 	 , A.KERATYPE  ");
                //strSql.AppendLine(" 	 , A.SUN ");
                //strSql.AppendLine(" 	 , A.GUBUN1 ");
                //strSql.AppendLine(" 	 , A.J_BNUM ");
                //strSql.AppendLine(" 	 , B.EMP_NM ");
                //strSql.AppendLine(" 	 , FORMAT(A.WEIGHT, 0) AS TOT_WEIGHT  ");
                //strSql.AppendLine(" 	 , FORMAT(A.FIRSTWEIGHT, 0) AS EMPTY_WEIGHT ");
                //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, 0) ELSE FORMAT(A.OWEIGHT, 0) END AS ACTL_WEIGHT ");
                //strSql.AppendLine(" 	 , A.GUBUN1 AS GUBUN2 ");
                //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, 0) ELSE FORMAT(A.OCHAGAM, 0) END AS LOSS ");
                //strSql.AppendLine(" 	 , A.J_STATE ");
                //strSql.AppendLine(" 	 , C.FAX ");
                //strSql.AppendLine("   FROM MESURING A ");
                //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                //strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID ");
                //strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD C ");
                //strSql.AppendLine("  	ON C.DEALER_NM = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END");
                //strSql.AppendLine("  WHERE A.JUNPYOID = '" + sJunpyoId + "' ");
                //strSql.AppendLine("  ORDER BY J_DATE DESC ");

                #endregion[이전 쿼리 2020-12-24]

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt.Rows.Count == 0)
                {
                    XtraMessageBox.Show("전표ID : " + sJunpyoId + "\r\n해당 데이터가 Mesuring에 존재하지 않습니다.");
                    return;
                }

                Dictionary<string, Image> Image = GetImage(sJunpyoId, dt.Rows[0]["J_DATE"]?.ToString());
                Dictionary<string, Image> result = new Dictionary<string, Image>();
                result.Add("1", Image["1_1"]);
                result.Add("2", Image["1_2"]);
                result.Add("3", Image["1_3"]);
                result.Add("4", Image["2_1"]);
                result.Add("5", Image["2_2"]);
                result.Add("6", Image["2_3"]);
                
                if(result["1"] == null)
                    result["1"] = (Image)Properties.Resources.No_Img;
                if (result["2"] == null)
                    result["2"] = (Image)Properties.Resources.No_Img;
                if (result["3"] == null)
                    result["3"] = (Image)Properties.Resources.No_Img;
                if (result["4"] == null)
                    result["4"] = (Image)Properties.Resources.No_Img;
                if (result["5"] == null)
                    result["5"] = (Image)Properties.Resources.No_Img;
                if (result["6"] == null)
                    result["6"] = (Image)Properties.Resources.No_Img;

                Cursor = Cursors.Default;
                
                //FaxViewer fm = new FaxViewer(dt, "RptMesuring", result);
                FaxViewer fm = new FaxViewer(dt, "RptWebFax", result);
                fm.ShowDialog();
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
            
        }

        private Dictionary<string, Image> GetImage(string sJunpyoID, string sJ_DATE)
        {
            string[] sJDateArr = sJ_DATE.Split(' ');
            string sJDate = sJDateArr[0];
            string[] strArr = sJDate.Split('-');
            //string ftpPath = @"ftp://192.168.0.202/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
            string ftpPath = @"ftp://"+ComnEtcFunc.FTP_IP+"/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
            string user = ComnEtcFunc.FTP_USER;
            string pw = ComnEtcFunc.FTP_PW;

            FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

            req1.Credentials = new NetworkCredential(user, pw);
            req1.Method = WebRequestMethods.Ftp.ListDirectory;

            //FTP 이미지를 Byte[]로 파싱하여 저장할 Dictionary 객체 초기세팅
            Dictionary<string, Image> dicPicture = new Dictionary<string, Image>();
            dicPicture.Add("1_1", null);
            dicPicture.Add("1_2", null);
            dicPicture.Add("1_3", null);
            dicPicture.Add("2_1", null);
            dicPicture.Add("2_2", null);
            dicPicture.Add("2_3", null);

            Dictionary<string, Image> dicCopy = new Dictionary<string, Image>();
            foreach (KeyValuePair<string, Image> item in dicPicture)
                dicCopy.Add(item.Key, null);

            try
            {
                string[] filesInDirectory = null;
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    foreach (KeyValuePair<string, Image> item in dicPicture)
                    {
                        //해당 파일 Index
                        int findIndex = Array.FindIndex(filesInDirectory, i => i == string.Format("{0}_{1}.jpg", sJunpyoID, item.Key));
                        if (findIndex >= 0)
                        {
                            string fileName = filesInDirectory[findIndex];
                            Image img = DownloadFTPFile(string.Format(@"{0}\{1}", ftpPath, fileName), user, pw);
                            dicCopy[item.Key] = img;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
            }

            return dicCopy;
        }

        #region FTP 파일 다운로드하기 - DownloadFTPFile(sourceFileURI, targetFilePath, userID, password)

        /// <summary>
        /// FTP 파일 다운로드하기
        /// </summary>
        /// <param name="sourceFileURI">소스 파일 URI</param>
        /// <param name="targetFilePath">타겟 파일 경로</param>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <returns>처리 결과</returns>

        public Image DownloadFTPFile(string sourceFileURI, string user, string pw)
        {
            Image img = null;
            try
            {
                Uri sourceFileUri = new Uri(sourceFileURI);
                FtpWebRequest ftpWebRequest = WebRequest.Create(sourceFileUri) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(user, pw);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;

                Stream sourceStream = ftpWebResponse.GetResponseStream();
                img = Image.FromStream(sourceStream);
                sourceStream.Close();
            }
            catch (Exception ex)
            {
                return null;
            }

            return img;
        }

        #endregion

        #region [정렬기능(2020-06-02 정은영)]
        private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        {
            GridView view = (GridView)sender;
            GridHitInfo hitInfo = view.CalcHitInfo(e.Location);
            
            //if (hitInfo.InColumn)
            //{
            //    if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.None)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
            //        GridViewRetr.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Descending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            //        GridViewRetr.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Ascending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.None;
            //        GridViewRetr.FocusedRowHandle = 0;
            //    }
            //    // if ((ModifierKeys & Keys.Control) == Keys.Control) return;
            //    //if ((ModifierKeys & Keys.Shift) != Keys.Shift) view.ClearSorting();
            //}
        }

        #endregion

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(e.FocusedRowHandle < 0)
            {
                BtnMesuringPrint.Enabled = false;
                BtnWebFex.Enabled = false;
                return;
            }

            BtnMesuringPrint.Enabled = true;

            string sYn = GridViewRetr.GetFocusedRowCellValue(GridColWebFaxYn)?.ToString();
            /*
            if (string.IsNullOrEmpty(sYn))
                BtnWebFex.Enabled = false;
            else if (sYn.Equals("Y"))
                BtnWebFex.Enabled = true;
            else
                BtnWebFex.Enabled = false;
            */
            if (string.IsNullOrEmpty(sYn))
                BtnWebFex.Enabled = false;
            else
                BtnWebFex.Enabled = true;
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GLkupEditCate_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            string sBigCate = string.Empty;

            foreach (int idx in GLkupEditCate.Properties.View.GetSelectedRows())
            {
                sBigCate = sBigCate + "'" + GLkupEditCate.Properties.View.GetRowCellValue(idx, GLkupEditCate.Properties.DisplayMember).ToString() + "' ,";
            }

            if (sBigCate.Length > 0) e.DisplayText = sBigCate.Substring(0, sBigCate.Length - 1);
        }

        private void GLkupEditCate_CloseUp(object sender, DevExpress.XtraEditors.Controls.CloseUpEventArgs e)
        {
            BtnRetr.Focus();
        }

        private void RdgbKeraType_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewRetr_RowClick(object sender, RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                try
                {
                    Cursor = Cursors.WaitCursor;

                    string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();
                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    #region MariaDB
                    //strSql.AppendLine(" SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM ");
                    //strSql.AppendLine(" 	 , CAST(DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS CHAR) AS J_DATE");
                    //strSql.AppendLine(" 	 , A.KERATYPE  ");
                    //strSql.AppendLine(" 	 , A.SUN ");
                    //strSql.AppendLine(" 	 , A.GUBUN1 ");
                    //strSql.AppendLine(" 	 , A.J_BNUM ");
                    //strSql.AppendLine(" 	 , B.EMP_NM ");
                    ////strSql.AppendLine("      , CONCAT(IFNULL(CHRG_RGN_NO, ''), REPLACE(FAX, '-', '')) AS FAX_NO ");
                    //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H:%i') ELSE DATE_FORMAT(A.FIRSTTIME, '%H:%i') END AS FIRSTTIME ");
                    //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H:%i') ELSE DATE_FORMAT(A.SECONDTIME, '%H:%i') END AS SECONDTIME ");
                    //strSql.AppendLine("      , CONCAT(FORMAT(A.SECONDWEIGHT, 0), ' KG') AS TOT_WEIGHT   ");
                    //strSql.AppendLine("      , CONCAT(FORMAT(A.FIRSTWEIGHT, 0), ' KG') AS EMPTY_WEIGHT  ");
                    //strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, 0) ELSE FORMAT(A.OWEIGHT, 0) END, ' KG') AS ACTL_WEIGHT  ");
                    //strSql.AppendLine("      , A.GUBUN1 AS GUBUN2  ");
                    //strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, 0) ELSE FORMAT(A.OCHAGAM, 0) END, ' KG') AS LOSS ");
                    //strSql.AppendLine(" 	 , A.J_STATE ");
                    //strSql.AppendLine("   FROM MESURING A ");
                    //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                    //strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID ");
                    //strSql.AppendLine("  WHERE A.JUNPYOID = '" + sJunpyoId + "' ");
                    //strSql.AppendLine("  ORDER BY J_DATE DESC ");
                    #endregion

                    strSql.AppendLine(" SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM ");
                    strSql.AppendLine("      , A.J_DATE AS J_DATE                                                              ");
                    strSql.AppendLine("      , A.KERATYPE                                                                      ");
                    strSql.AppendLine("      , A.SUN                                                                           ");
                    strSql.AppendLine("      , A.GUBUN1                                                                        ");
                    strSql.AppendLine("      , A.J_BNUM                                                                        ");
                    strSql.AppendLine("      , B.EMP_NM                                                                        ");
                    strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(CONVERT(DATETIME, A.SECONDTIME), 'HH:mm') ELSE FORMAT(CONVERT(DATETIME, A.FIRSTTIME), 'HH:mm') END AS FIRSTTIME ");
                    strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN FORMAT(CONVERT(DATETIME, A.FIRSTTIME), 'HH:mm') ELSE FORMAT(CONVERT(DATETIME, A.SECONDTIME), 'HH:mm') END AS SECONDTIME");
                    strSql.AppendLine("      , CONCAT(FORMAT(A.SECONDWEIGHT, '#,0'), ' KG') AS TOT_WEIGHT                                                                                                ");
                    strSql.AppendLine("      , CONCAT(FORMAT(A.FIRSTWEIGHT, '#,0'), ' KG') AS EMPTY_WEIGHT                                                                                               ");
                    strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, '#,0') ELSE FORMAT(A.OWEIGHT, '#,0') END, ' KG') AS ACTL_WEIGHT                               ");
                    strSql.AppendLine("      , A.GUBUN1 AS GUBUN2                                                                                                                                        ");
                    strSql.AppendLine("      , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, '#,0') ELSE FORMAT(A.OCHAGAM, '#,0') END, ' KG') AS LOSS                                      ");
                    strSql.AppendLine("       , A.J_STATE                                                                                                                                                ");
                    strSql.AppendLine("    FROM MESURING A                                                                                                                                               ");
                    strSql.AppendLine("    LEFT OUTER JOIN HR_EMP_BASIS B                                                                                                                                ");
                    strSql.AppendLine("      ON A.GUMSU_SERIAL = B.EMP_ID                                                                                                                                ");
                    strSql.AppendLine("   WHERE A.JUNPYOID = '" + sJunpyoId + "'                                                                                                                                    ");
                    strSql.AppendLine("   ORDER BY J_DATE DESC                     ");

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    if (dt.Rows.Count == 0)
                    {
                        XtraMessageBox.Show("전표ID : " + sJunpyoId + "\r\n해당 데이터가 Mesuring에 존재하지 않습니다.");
                        return;
                    }

                    Dictionary<string, Image> Image = GetImage(sJunpyoId, dt.Rows[0]["J_DATE"]?.ToString());

                    Dictionary<string, Image> result = new Dictionary<string, Image>();
                    result.Add("1", Image["1_1"]);
                    result.Add("2", Image["1_2"]);
                    result.Add("3", Image["1_3"]);
                    result.Add("4", Image["2_1"]);
                    result.Add("5", Image["2_2"]);
                    result.Add("6", Image["2_3"]);

                    if (result["1"] == null)
                        result["1"] = (Image)Properties.Resources.No_Img;
                    if (result["2"] == null)
                        result["2"] = (Image)Properties.Resources.No_Img;
                    if (result["3"] == null)
                        result["3"] = (Image)Properties.Resources.No_Img;
                    if (result["4"] == null)
                        result["4"] = (Image)Properties.Resources.No_Img;
                    if (result["5"] == null)
                        result["5"] = (Image)Properties.Resources.No_Img;
                    if (result["6"] == null)
                        result["6"] = (Image)Properties.Resources.No_Img;

                    Cursor = Cursors.Default;

                    ReportViewer fm = new ReportViewer(dt, "RptWebFax", result);
                    fm.ShowDialog();

                    //RptWebFax report = new RptWebFax(dt.Rows[0], result);
                    //ReportPrintTool printTool = new ReportPrintTool(report);
                    //report.CreateDocument();
                    //printTool.Print();
                }
                catch (Exception ex)
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }

        private void GLkupEditCate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateEditYmd.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateEditYmd.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateEditYmd.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateEditYmd.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }
    }
}