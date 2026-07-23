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
using Popbill.Fax;
using System.Net;
using System.IO;
using DevExpress.XtraReports.UI;
using DevExpress.XtraGrid.Views.Grid;

/*
 * 작성일자 : 2021-01-06
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-01-24
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            1. 레이아웃 저장기능 추가
 *            
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 */

namespace AccAdm
{
    public partial class SYS003F01 : DevExpress.XtraEditors.XtraForm
    {
        public SYS003F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void SYS003F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr, GridViewSub };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            string sFile = Application.StartupPath + @"\xaml\" + FmMainToolBar2.UserID + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            string sFindIdx = CboFindIdx.SelectedIndex.ToString();
            string sFIndWord = TxtFindWord.EditValue?.ToString().Trim();

            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("검색일자를 올바르게 입력하세요.");
                DateEditFrom.SelectAll();
                DateEditFrom.Focus();
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("검색일자를 올바르게 입력하세요.");
                DateEditTo.SelectAll();
                DateEditTo.Focus();
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("DATE_F", sYmdFrom);
            dicParams.Add("DATE_T", sYmdTo);
            dicParams.Add("FIND_IDX", sFindIdx);
            dicParams.Add("FIND_WORD", sFIndWord);

            GridRetr.DataSource = GetMesureInfo(dicParams);
        }

        private DataTable GetMesureInfo(Dictionary<string, string> dicParams)
        {
            DataTable dt = new DataTable();
            try
            {
                Cursor = Cursors.WaitCursor;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT A.SUN");
                strSql.AppendLine("      , A.JUNPYOID ");
                strSql.AppendLine("      , A.KERATYPE ");
                strSql.AppendLine("      , A.J_BNUM ");
                strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS J_COMPANY ");
                strSql.AppendLine("      , A.GUBUN1 ");
                strSql.AppendLine("      , CONCAT(ISNULL(CHRG_RGN_NO, ''), REPLACE(FAX, '-', '')) AS FAX_NO ");
                strSql.AppendLine("      , CASE WHEN A.OWEIGHT = 0 THEN A.IWEIGHT ELSE A.OWEIGHT END AS OWEIGHT ");
                strSql.AppendLine("      , A.FIRSTWEIGHT ");
                strSql.AppendLine("      , SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.FIRSTTIME), 14), 1, 5) AS FIRSTTIME  ");
                strSql.AppendLine("      , A.SECONDWEIGHT                                                                         ");
                strSql.AppendLine("      , SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.SECONDTIME), 14), 1, 5) AS SECONDTIME");
                strSql.AppendLine(" 	 , A.WEIGHT   ");
                strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END AS OCHAGAM ");
                strSql.AppendLine(" 	 , A.J_STATE ");
                strSql.AppendLine(" 	 , A.GUMSUBIGO	 ");
                strSql.AppendLine("      , B.EMP_NM ");
                strSql.AppendLine("      , C.WEB_FAX_YN ");
                strSql.AppendLine("      , ISNULL(D.CNT, 0) AS CNT ");
                strSql.AppendLine("      , 0 AS SCS_CNT ");
                strSql.AppendLine("      , 0 AS FAIL_CNT ");
                strSql.AppendLine("   FROM MESURING A ");
                strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID ");
                strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C  ");
                strSql.AppendLine("     ON C.DEALER_CD = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHERID ELSE A.J_ASSIGNID END ");
                strSql.AppendLine("   LEFT JOIN ( SELECT X1.JUNPYOID ");
                strSql.AppendLine("                    , COUNT(*) AS CNT ");
                strSql.AppendLine("                 FROM WEB_FAX_LOG X1 ");
                strSql.AppendLine("                GROUP BY JUNPYOID ) D ");
                strSql.AppendLine("     ON A.JUNPYOID = D.JUNPYOID ");
                strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
                strSql.AppendLine("    AND A.KERATYPE <> '직송' ");
                strSql.AppendLine("    AND C.WEB_FAX_YN = 'Y' ");
                strSql.AppendLine("    AND (('" + dicParams["FIND_WORD"] + "' IS NULL OR '" + dicParams["FIND_WORD"]+ "' = '') AND 1 = 1 ");
                strSql.AppendLine("         OR ");
                strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "' = '0' AND CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END LIKE '%" + dicParams["FIND_WORD"] + "%')");
                strSql.AppendLine("         OR ");
                strSql.AppendLine("         ('" + dicParams["FIND_IDX"] + "'= '1' AND A.GUBUN1 LIKE '%" + dicParams["FIND_WORD"] + "%'))");
                strSql.AppendLine("  ORDER BY A.KERATYPE, A.J_COMPANY, A.SUN, A.FIRSTTIME, A.SECONDTIME ");

                dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                FaxInfo faxInfo = new FaxInfo();
                FaxService faxService = new FaxService(faxInfo.LinkID, faxInfo.SECRET_KEY);
                faxService.IsTest = ComnEtcFunc.IsFaxTest;
                int iSuccessCnt = 0;
                int iFailCnt = 0;
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    Dictionary<string, string> dicParam = new Dictionary<string, string>();

                    dicParam.Add("JUNPYOID", dt.Rows[j]["JUNPYOID"]?.ToString());

                    DataTable dtChk = GetFaxInfo(dicParam);
                    for (int i = 0; i < dtChk.Rows.Count; i++)
                    {
                        try
                        {
                            List<FaxResult> result = faxService.GetFaxResult(faxInfo.CorpNo, dtChk.Rows[i]["WEB_FAX_NO"]?.ToString());
                            int iresultCode = Convert.ToInt32(result[0].result); //전송결과 https://docs.popbill.com/fax/sendCode?lang=dotnet 참고
                            if (iresultCode == 100)
                                iSuccessCnt++;
                            else
                                iFailCnt++;
                        }
                        catch (Exception ex)
                        {
                            //iFailCnt++;
                        }
                    }
                    dt.Rows[j]["SCS_CNT"] = iSuccessCnt;
                    dt.Rows[j]["FAIL_CNT"] = iFailCnt;
                    iSuccessCnt = 0;
                    iFailCnt = 0;
                }

                //foreach (DataRow row in dt.Rows)
                //{
                //    Dictionary<string, string> dicParam = new Dictionary<string, string>();

                //    dicParam.Add("JUNPYOID", row["JUNPYOID"]?.ToString());

                //    DataTable dtChk = GetFaxInfo(dicParams);
                //    for (int i = 0; i < dt.Rows.Count; i++)
                //    {
                //        try
                //        {
                //            List<FaxResult> result = faxService.GetFaxResult(faxInfo.CorpNo, dtChk.Rows[i]["WEB_FAX_NO"]?.ToString());
                //            int iresultCode = Convert.ToInt32(result[0].result); //전송결과 https://docs.popbill.com/fax/sendCode?lang=dotnet 참고
                //            if (iresultCode == 100)
                //                iSuccessCnt++;
                //            else
                //                iFailCnt++;
                //        }
                //        catch (Exception ex)
                //        {
                //            //iFailCnt++;
                //        }
                //    }
                //    row["SCS_CNT"] = iSuccessCnt;
                //    row["FAIL_CNT"] = iFailCnt;
                //}

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
            }
            
            return dt;
        }

        private DataTable GetFaxInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.JUNPYOID ");
            strSql.AppendLine("      , A.WEB_FAX_NO ");
            strSql.AppendLine("      , A.FAX_SND_NO ");
            strSql.AppendLine("      , B.PGMNM ");
            strSql.AppendLine("      , C.USRNM ");
            strSql.AppendLine("      , A.REG_DT ");
            strSql.AppendLine("      , '' AS STSGB ");
            strSql.AppendLine("      , '' AS RESULT ");
            strSql.AppendLine("   FROM WEB_FAX_LOG A  ");
            strSql.AppendLine("   LEFT JOIN ZPGMLST B ");
            strSql.AppendLine("     ON A.PGM_ID = B.PGMID ");
            strSql.AppendLine("   LEFT JOIN ZUSRLST C  ");
            strSql.AppendLine("     ON A.REG_ID = C.USRCD ");
            strSql.AppendLine("  WHERE A.JUNPYOID = @JUNPYOID  ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {

        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
            {
                GridSub.DataSource = null;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                
                FaxInfo faxInfo = new FaxInfo();
                FaxService faxService = new FaxService(faxInfo.LinkID, faxInfo.SECRET_KEY);
                faxService.IsTest = ComnEtcFunc.IsFaxTest;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("JUNPYOID", GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString());

                DataTable dt = GetFaxInfo(dicParams);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        List<FaxResult> result = faxService.GetFaxResult(faxInfo.CorpNo, dt.Rows[i]["WEB_FAX_NO"]?.ToString());
                        int istateCode = Convert.ToInt32(result[0].state); //전송상태 https://docs.popbill.com/fax/sendCode?lang=dotnet 참고
                        int iresultCode = Convert.ToInt32(result[0].result); //전송결과 https://docs.popbill.com/fax/sendCode?lang=dotnet 참고
                        dt.Rows[i]["STSGB"] = faxInfo.GetState(istateCode);
                        dt.Rows[i]["RESULT"] = faxInfo.GetResult(iresultCode);
                    }
                    catch(Exception ex)
                    {
                        dt.Rows[i]["STSGB"] = "시스템에러";
                        dt.Rows[i]["RESULT"] = ex.Message;
                    }
                }

                GridSub.DataSource = dt;

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }

        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void SYS003F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnChrg.PerformClick();
        }

        private void BtnChrg_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                Cursor = Cursors.WaitCursor;
                FaxInfo faxInfo = new FaxInfo();
                FaxService faxService = new FaxService(faxInfo.LinkID, faxInfo.SECRET_KEY);
                faxService.IsTest = ComnEtcFunc.IsFaxTest;

                //string sURL = faxService.GetChargeURL(faxInfo.CorpNo, faxInfo.USER_ID);
                string sURL = faxService.GetPartnerURL(faxInfo.CorpNo, "CHRG");
                System.Diagnostics.Process.Start(sURL);

                Cursor = Cursors.Default;
                //WebRequest request = WebRequest.Create(sURL);
                //request.Method = "POST";
                //request.ContentType = "application/x-www-form-urlencoded";

                ////전달할 파라메타
                //string sendData = "param1=1&param2=2";

                //byte[] buffer;
                //buffer = Encoding.Default.GetBytes(sendData);
                //request.ContentLength = buffer.Length;
                //Stream sendStream = request.GetRequestStream();
                //sendStream.Write(buffer, 0, buffer.Length);
                //sendStream.Close();
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnHomePG_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                FaxInfo faxInfo = new FaxInfo();
                FaxService faxService = new FaxService(faxInfo.LinkID, faxInfo.SECRET_KEY);
                faxService.IsTest = ComnEtcFunc.IsFaxTest;

                //string sURL = faxService.GetChargeURL(faxInfo.CorpNo, faxInfo.SYNC_USER_ID);
                string sURL = faxService.GetAccessURL(faxInfo.CorpNo, faxInfo.SYNC_USER_ID);
                System.Diagnostics.Process.Start(sURL);

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if(e.Column == GridColCnt)
            {
                int iCnt = 0;
                int.TryParse(e.CellValue?.ToString(), out iCnt);
                if(iCnt > 0)
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else
                {
                    e.Appearance.BackColor = Color.LightSalmon;
                }
            }
            else if(e.Column == GridColSuceessCnt)
            {
                int iCnt = 0;
                int.TryParse(e.CellValue?.ToString(), out iCnt);
                if (iCnt > 0)
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else
                {
                    e.Appearance.BackColor = Color.LightSalmon;
                }
            }
            else if(e.Column == GridColFailCnt)
            {
                int iCnt = 0;
                int.TryParse(e.CellValue?.ToString(), out iCnt);
                if (iCnt > 0)
                {
                    e.Appearance.BackColor = Color.LightSalmon;
                }
                else
                {
                    int iTransferCnt = 0;
                    int.TryParse(GridViewRetr.GetRowCellValue(e.RowHandle, GridColCnt)?.ToString(), out iTransferCnt);
                    if (iTransferCnt == 0)
                    {
                        e.Appearance.BackColor = Color.LightSalmon;
                    }
                    else
                    {
                        e.Appearance.BackColor = Color.PaleGreen;
                    }
                }
            }
        }

        private void BtnWebFex_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                Cursor = Cursors.WaitCursor;
                string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");

                strSql.AppendLine("SELECT A.JUNPYOID                                                                                                                 ");
                strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM                                            ");
                strSql.AppendLine("     , A.J_DATE AS J_DATE                                                                                                         ");
                strSql.AppendLine("     , A.KERATYPE                                                                                                                 ");
                strSql.AppendLine("     , A.SUN                                                                                                                      ");
                strSql.AppendLine("     , A.GUBUN1                                                                                                                   ");
                strSql.AppendLine("     , A.J_BNUM                                                                                                                   ");
                strSql.AppendLine("     , A.GUMSU_SERIAL                                                                                                             ");
                strSql.AppendLine("     , B.EMP_NM                                                                                                                   ");
                strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.SECONDTIME), 14), 1, 5)               ");
                strSql.AppendLine("             ELSE SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.FIRSTTIME), 14), 1, 5) END AS FIRSTTIME                       ");
                strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.FIRSTTIME), 14), 1, 5)                ");
                strSql.AppendLine("             ELSE SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.SECONDTIME), 14), 1, 5) END AS SECONDTIME                     ");
                strSql.AppendLine("     , CONCAT(FORMAT(A.SECONDWEIGHT, '#,#'), ' KG') AS TOT_WEIGHT                                                                 ");
                strSql.AppendLine("     , CONCAT(FORMAT(A.FIRSTWEIGHT, '#,#'), ' KG') AS EMPTY_WEIGHT                                                                ");
                strSql.AppendLine("     , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, '#,#') ELSE FORMAT(A.OWEIGHT, '#,#') END, ' KG') AS ACTL_WEIGHT");
                strSql.AppendLine("     , A.GUBUN1 AS GUBUN2                                                                                                         ");
                strSql.AppendLine("     , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, '#,#') ELSE FORMAT(A.OCHAGAM, '#,#') END, ' KG') AS LOSS       ");
                strSql.AppendLine("     , A.J_STATE                                                                                                                  ");
                strSql.AppendLine("     , CONCAT(ISNULL(C.CHRG_RGN_NO, ''), REPLACE(C.FAX, '-', '')) AS FAX                                                          ");
                strSql.AppendLine(" FROM MESURING A                                                                                                                  ");
                strSql.AppendLine(" LEFT OUTER JOIN HR_EMP_BASIS B                                                                                                   ");
                strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID                                                                                                 ");
                strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD C                                                                                                  ");
                strSql.AppendLine("     ON C.DEALER_NM = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END                                          ");
                strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");
                strSql.AppendLine(" ORDER BY J_DATE DESC                                                                                                             ");

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

                #region mariaDB
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
                //strSql.AppendLine("      , CONCAT(ISNULL(C.CHRG_RGN_NO, ''), REPLACE(C.FAX, '-', '')) AS FAX ");
                //strSql.AppendLine("   FROM MESURING A  ");
                //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B  ");
                //strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID  ");
                //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C  ");
                //strSql.AppendLine("  	 ON C.DEALER_NM = CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END ");
                //strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");
                //strSql.AppendLine("  ORDER BY J_DATE DESC  ");
                #endregion

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

                //FaxViewer fm = new FaxViewer(dt, "RptMesuring", result);
                FaxViewer fm = new FaxViewer(dt, "RptWebFax", result);
                fm.ShowDialog();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnMesuringPrint_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("PRINT", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 출력 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                //string sAccCd = GrdViewAccUseCd.GetFocusedRowCellValue("ACC_CD")?.ToString();
                string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine("SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM                                              ");
                strSql.AppendLine("     , A.J_DATE AS J_DATE                                                                                                           ");
 	            strSql.AppendLine("     , A.KERATYPE                                                                                                                   ");
 	            strSql.AppendLine("     , A.SUN                                                                                                                        ");
 	            strSql.AppendLine("     , A.GUBUN1                                                                                                                     ");
 	            strSql.AppendLine("     , A.J_BNUM                                                                                                                     ");
 	            strSql.AppendLine("     , B.EMP_NM                                                                                                                     ");
                strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.SECONDTIME), 14), 1, 5)                 ");
                strSql.AppendLine("            ELSE SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.FIRSTTIME), 14), 1, 5) END AS FIRSTTIME                          ");
                strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.FIRSTTIME), 14), 1, 5)                  ");
                strSql.AppendLine("            ELSE SUBSTRING(CONVERT(VARCHAR(8), CONVERT(DATETIME, A.SECONDTIME), 14), 1, 5) END AS SECONDTIME                        ");
                strSql.AppendLine("     , CONCAT(FORMAT(A.SECONDWEIGHT, '#,#'), ' KG') AS TOT_WEIGHT                                                                   ");
                strSql.AppendLine("     , CONCAT(FORMAT(A.FIRSTWEIGHT, '#,#'), ' KG') AS EMPTY_WEIGHT                                                                  ");
                strSql.AppendLine("     , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.IWEIGHT, '#,#') ELSE FORMAT(A.OWEIGHT, '#,#') END, ' KG') AS ACTL_WEIGHT  ");
                strSql.AppendLine("     , A.GUBUN1 AS GUBUN2                                                                                                           ");
                strSql.AppendLine("     , CONCAT(CASE WHEN A.KERATYPE = '입고' THEN FORMAT(A.ICHAGAM, '#,#') ELSE FORMAT(A.OCHAGAM, '#,#') END, ' KG') AS LOSS         ");
                strSql.AppendLine("     , A.J_STATE                                                                                                                    ");
                strSql.AppendLine("  FROM MESURING A                                                                                                                   ");
                strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS B                                                                                                    ");
                strSql.AppendLine("    ON A.GUMSU_SERIAL = B.EMP_ID                                                                                                    ");
                strSql.AppendLine("  WHERE A.JUNPYOID = '" + sJunpyoId + "' ");
                strSql.AppendLine(" ORDER BY J_DATE DESC                                                                                                               ");

                #region mariaDB
                //strSql.AppendLine(" SELECT CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER_NM ");
                //strSql.AppendLine(" 	 , CAST(DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS CHAR) AS J_DATE");
                //strSql.AppendLine(" 	 , A.KERATYPE  ");
                //strSql.AppendLine(" 	 , A.SUN ");
                //strSql.AppendLine(" 	 , A.GUBUN1 ");
                //strSql.AppendLine(" 	 , A.J_BNUM ");
                //strSql.AppendLine(" 	 , B.EMP_NM ");
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
            catch (Exception ex)
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
            string ftpPath = @"ftp://"+ ComnEtcFunc.FTP_IP + @"/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
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

        private void SYS003F01_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                ////string path = @"C:\STLNT\" + sProject + @"\xaml\" + sId;
                string path = Application.StartupPath + @"\xaml\" + FmMainToolBar2.UserID;
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                
                layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");
                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}