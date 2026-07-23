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
using MySql.Data.MySqlClient;
using System.Diagnostics;
using DevExpress.XtraBars;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.IO;
using System.Data.SqlClient;
using DevExpress.XtraGrid.Columns;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 수정일자 : 2021-02-07
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            거래처초성검색 추가 (쿼리)
*            
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*            2. 레이아웃 전체 저장 설정
*            
* 수정일자 : 2021-03-03
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 검색조건으로 검색 시 검색대상이 속한 전표번호의 전표건도 포함하여 출력하도록 수정
*              메소드 추가하여 적용
*             
* 수정일자 : 2021-03-10
* 수정자   : 고혜성
* Reference Key : #0001
* 수정내용 : (현업요청)
*            1. 검색조건으로 검색 시 검색대상이 속한 전표번호의 전표건도 포함하여 출력하도록 수정
*              메소드 추가하여 적용
* 
* 수정일자 : 2021-03-17
* 수정자   : 고혜성
* Reference Key = #0002
* 수정내용 : (현업요청)
*            1. 로그수정
*               1-1) 기본사항/변경사항을 나누어 입력
* 
* 수정일자 : 2021-03-22
* 수정자   : 고혜성
* Reference Key : #0003
* 수정내용 : (현업요청)
*            승인취소일 경우에만 Log적용
*            
* 수정일자 : 2021-06-07
* 수정자   : 정은영
* ID       : #0004
* 수정내용 : (현업요청)
*            검색조건 계정코드, 성격, 관계계정코드, 관계계정명 삭제 및 자동전표구분 추가
*/
namespace AccAdm
{
    public partial class AC02001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC02001F01_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            DateEditFrom.EditValue = today.AddDays(1 - today.Day);
            DateEditTo.EditValue = today;

            DataTable dtAcrDr = GetLookUpData("1", "Y", "Y");
            DataTable dtAuto = GetLookUpData("4", "Y", "Y");
            DataTable dtEmp = GetLookUpData("5", "Y", "Y");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAtGub, dtAcrDr, GridSlip, GridColAtGub, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAtGub, dtAcrDr, GridSlip, GridColAtGub2, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAAuto, dtAuto, GridSlip, GridColAAuto, "CD", "NM", "");
            //ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAAuto, dtAuto, GridSlip, GridColAUser, "CD", "NM", "");
            RepoGLkupUser.ValueMember = "CD";
            RepoGLkupUser.DisplayMember = "NM";
            RepoGLkupUser.DataSource = dtEmp;

            InitControls();
            UpdateDropDownButton(BtnSlipCopy);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewSlip };
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

            SettingTextEditAutoComplete(CboFindSbj.EditValue?.ToString());
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                //_MODE = SelectionMode.Single;
                //GridViewSlip.OptionsSelection.MultiSelect = false;
                //GridViewSlip.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
                //GridViewSlip.OptionsSelection.CheckBoxSelectorColumnWidth = 40;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                string sAprvGb = RdgbAprvGb.EditValue?.ToString();

                if (string.IsNullOrEmpty(sYmdFrom) || string.IsNullOrEmpty(sYmdTo))
                {
                    XtraMessageBox.Show("검색기간을 올바르게 설정하세요.");
                    return;
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);
                dicParams.Add("FIND_IDX", CboFindSbj.SelectedIndex.ToString());
                dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());
                dicParams.Add("APRV_YN", sAprvGb);

                /*
                 * 2021-03-03 (현업요청)
                 * 검색조건으로 검색 시 검색대상이 속한 전표번호의 전표건도 포함하여 출력하도록 수정
                 * 메소드 추가하여 적용
                 */
                //GridSlip.DataSource = GetSlipInfo(dicParams);

                /*
                 * 2021-03-04(현업요청)
                 * 
                 */
                GridSlip.DataSource = GetSlipInfo(sYmdFrom, sYmdTo, ReturningByComboBoxValues(CboFindSbj.EditValue?.ToString(), TxtFindWord.EditValue?.ToString().Replace("-", "")), sAprvGb);
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

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            AC02001F02 frm = new AC02001F02();
            frm.Owner = this;
            frm.AddModifyGb = "ADD";
            frm.PARENT_FORM = this;
            frm.DataRowSendEvent += new AC02001F02.SendDataHandler(GetDataRow);
            frm.Show();

            //BeginInvoke(new Action(() =>
            //{
            //    AC02001F02 frm = new AC02001F02();
            //    frm.Owner = this;
            //    //frm.MdiParent = this.Owner;
            //    frm.PARENT_FORM = this;
            //    frm.AddModifyGb = "ADD";

            //    //frm.ShowDialog(this);
            //    frm.Show();
            //    //if (frm.ShowDialog() == DialogResult.OK)
            //    //{
            //    //    BtnRetr_Click(null, null);
            //    //}
            //}));

            //AC02001F02 frm = new AC02001F02();

            //frm.PARENT_FORM = this;
            //frm.AddModifyGb = "ADD";
            //Button b = new Button();
            //frm.Controls.Add(b);
            //b.Click += new EventHandler(click);
            //frm.FormClosed += new FormClosedEventHandler(form2_closed);
            //frm.FormClosing += new FormClosingEventHandler(form2_closing);
            //this.Show();
            //do
            //{
            //    frm.ShowDialog();
            //} while (frm.IsDisposed == false);

        }

        public void GetDataRow(Dictionary<string, string> dicParams)
        {
            BtnRetr.PerformClick();
            int[] iArr = new int[dicParams.Count];
            int iResult = 0;

            string sTdate = DateTime.ParseExact(dicParams["TDATE"].ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd");

            iResult = GridViewSlip.LocateByValue(0, GridColTDate, sTdate);
            iResult = GridViewSlip.LocateByValue(iResult, GridColAtGub, dicParams["ATGUB"]);
            iResult = GridViewSlip.LocateByValue(iResult, GridColSeqNo, dicParams["SEQNO"]);

            GridViewSlip.FocusedRowHandle = iResult;
        }

        private void click(object sender, EventArgs e)
        {
            ((Form)((Control)sender).Parent).ShowInTaskbar = !((Form)((Control)sender).Parent).ShowInTaskbar;
        }

        private void form2_closed(object sender, FormClosedEventArgs e)
        {
            ((Form)sender).Dispose();
        }

        private void form2_closing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.None)
                e.Cancel = true;
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sSlipYmd = GridViewSlip.GetFocusedRowCellValue("TDATE")?.ToString();
            string sSlipGb = GridViewSlip.GetFocusedRowCellDisplayText("ATGUB")?.ToString();
            string sSeqNo = GridViewSlip.GetFocusedRowCellValue("SEQNO")?.ToString();

            string sSlipInfo = "전표일자 : " + sSlipYmd
                + "\r\n" + "전표구분 : " + sSlipGb
                + "\r\n" + "전표번호 : " + sSeqNo;

            if (XtraMessageBox.Show(sSlipInfo + "\r\n\r\n" + "선택하신 정보가 삭제됩니다.\r\n정말로 진행하시겠습니까?"
                , "삭제여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                int iIdx = GridViewSlip.FocusedRowHandle;

                DeleteSlipInfo(GridViewSlip.GetFocusedDataRow());

                GridViewSlip.FocusedRowHandle = iIdx - 1;
            }
        }

        private void BtnRmkEdit_Click(object sender, EventArgs e)
        {
            AC02001F04 frm = new AC02001F04();

            if(GridViewSlip.GetFocusedDataRow() == null)
            {
                XtraMessageBox.Show("적요보완 등록할 항목을 선택해주세요.");
                return;
            }

            frm.DrSlipInfo = GridViewSlip.GetFocusedDataRow();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                DataRow row = GridViewSlip.GetFocusedDataRow();

                BtnRetr.PerformClick();
                SetFocusedRow(row["TDATE"].ToString().Replace("-", "").Substring(0, 8), row["ATGUB"].ToString(), row["SEQNO"].ToString());
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
                string sFileNM = "전표리스트_" + DateTime.Now.ToLongDateString().Replace(" ", "");
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

        private void GridViewSlip_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                //전표승인 여부 체크
                string sAprv = GridViewSlip.GetFocusedRowCellValue("APVYN")?.ToString();
                //if (!string.IsNullOrEmpty(sAprv))
                //{
                //    if (sAprv.Equals("Y"))
                //    {
                //        XtraMessageBox.Show("승인된 전표는 수정할 수 없습니다.");
                //        return;
                //    }
                //}

                //자동전표 여부체크
                string sAuto = GridViewSlip.GetFocusedRowCellValue("AAUTO")?.ToString();
                //if (!string.IsNullOrEmpty(sAuto))
                //{
                //    //A01 -> MESURING, INLIST에서 처리되는 전표 REF1 참조
                //    if (sAuto.Equals("A01") || sAuto.Equals("A02"))
                //    {
                //        XtraMessageBox.Show("자동전표는 수정할 수 없습니다.\r\n[출처 : 매입출차료승인]");
                //        return;
                //    }
                //}

                AC02001F02 frm = new AC02001F02();
                frm.Owner = this;
                frm.PARENT_FORM = this;
                frm.AddModifyGb = "MODIFY";
                frm.DrSlipInfo = GridViewSlip.GetFocusedDataRow();
                if (sAuto.Equals("A01") || sAuto.Equals("A02"))
                {
                    frm._EDIT_YN = AC02001F02.EDIT_YN.NonEditable;
                    frm._EDIT_GB = AC02001F02.EDIT_GB.Auto;
                }
                else if (sAprv.Equals("Y"))
                {
                    frm._EDIT_YN = AC02001F02.EDIT_YN.NonEditable;
                    frm._EDIT_GB = AC02001F02.EDIT_GB.Approval;
                }

                /* 
                 * 2020-01-06 로직추가
                 * AAUTO 'D01'(지불관리)를 전표관리에서 수정하여 SUGMF과 SYNC를 맞추기위하여 변수세팅
                 * 변수를 주지 않을 시 Default None으로 세팅
                 */
                if (sAuto.Equals("D01"))
                {
                    frm._SYNC = AC02001F02.SYNC_GB.Jibul;
                }
                frm.DataRowSendEvent += new AC02001F02.SendDataHandler(GetDataRow);
                frm.Show();

                //AC02001F02 frm = new AC02001F02();
                //frm.PARENT_FORM = this;
                //frm.AddModifyGb = "MODIFY";
                //frm.DrSlipInfo = GridViewSlip.GetFocusedDataRow();
                //if (frm.ShowDialog() == DialogResult.OK)
                //{
                //    BtnRetr_Click(null, null);
                //}
            }
            else
            {
                //if(_MODE == SelectionMode.Multi)
                //{
                //    bool bSelected = GridViewSlip.IsRowSelected(e.RowHandle);
                //    DataRow select_row = GridViewSlip.GetDataRow(e.RowHandle);
                //    string sSELECT_TDATE = select_row["TDATE"]?.ToString();
                //    string sSELECT_ATGUB = select_row["ATGUB"]?.ToString();
                //    string sSELECT_SEQNO = select_row["SEQNO"]?.ToString();
                //    for (int i = 0 ;i < GridViewSlip.RowCount; i++)
                //    {
                //        DataRow row = GridViewSlip.GetDataRow(i);
                //        string sTDATE = row["TDATE"]?.ToString();
                //        string sATGUB = row["ATGUB"]?.ToString();
                //        string sSEQNO = row["SEQNO"]?.ToString();
                //        if (sTDATE.Equals(sSELECT_TDATE))
                //        {
                //            if (sATGUB.Equals(sSELECT_ATGUB))
                //            {
                //                if (sSEQNO.Equals(sSELECT_SEQNO))
                //                {
                //                    if (bSelected)
                //                        GridViewSlip.UnselectRow(i);
                //                    else
                //                        GridViewSlip.SelectRow(i);
                //                }
                //                else
                //                {
                //                    continue;
                //                }
                //            }
                //            else
                //            {
                //                continue;
                //            }
                //        }
                //        else
                //        {
                //            continue;
                //        }
                //    }
                //}
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        #region[KEY_DOWN_EVENT]

        private void AC02001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {

            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnDel_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F7)
            {
                BtnRmkEdit_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        #endregion[KEY_DOWN_EVENT]

        #region[Execute By Query]

        private DataTable GetSlipInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" WITH SLIP_INFO AS ( ");
            strSql.AppendLine("      SELECT X1.TDATE, X1.ATGUB, X1.SEQNO ");
            strSql.AppendLine("        FROM ACTRAN X1 ");
            strSql.AppendLine("        LEFT JOIN ACMSTF X2 ");
            strSql.AppendLine("          ON X1.ACCOD = X2.ACCOD ");
            strSql.AppendLine("        LEFT JOIN COM_BASE_CD X3 ");
            strSql.AppendLine("          ON X2.AGUBN = X3.COM_CD ");
            strSql.AppendLine("         AND X3.CD_GB = 'AC01001_03' ");
            strSql.AppendLine("        LEFT JOIN ACMSTF X4 --관계계정코드 ");
            strSql.AppendLine("          ON X2.ASMCD = X4.ACCOD ");
            strSql.AppendLine("       WHERE X1.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("         AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '0' AND X1.ACCOD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '1' AND X2.ACNAM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '2' AND X3.COM_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '3' AND X2.ASMCD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '4' AND X4.ACNAM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '5' AND X1.ATEXT LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("              OR  ");
            strSql.AppendLine("              ('" + dicParams["FIND_IDX"] + "' = '6' AND X1.CVNAM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
            strSql.AppendLine("         AND (('" + dicParams["APRV_YN"] + "' = '0' AND 1 = 1) ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ('" + dicParams["APRV_YN"] + "' = '1' AND X1.APVYN = 'Y') ");
            strSql.AppendLine("              OR ");
            strSql.AppendLine("              ('" + dicParams["APRV_YN"] + "' = '2' AND X1.APVYN <> 'Y')) ");
            strSql.AppendLine("       GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO ");
            strSql.AppendLine("    ) ");
            strSql.AppendLine(" SELECT CONVERT(VARCHAR(10),CONVERT(DATE,A.TDATE),23) AS TDATE1     ");
            strSql.AppendLine("       , A.ATGUB AS ATGUB1    ");
            strSql.AppendLine("       , A.SEQNO AS SEQNO1    ");
            strSql.AppendLine("       , A.LINNO AS LINNO1    ");
            strSql.AppendLine("       , CONVERT(VARCHAR(10),CONVERT(DATE,A.TDATE),23) AS TDATE     ");
            strSql.AppendLine("       , A.ATGUB    ");
            strSql.AppendLine("       , A.SEQNO    ");
            strSql.AppendLine("       , A.LINNO    ");
            strSql.AppendLine("       , A.ACCOD    ");
            strSql.AppendLine("       , B.ACNAM AS ACNAM    ");
            strSql.AppendLine("       , A.CVCOD AS CVCOD     ");
            strSql.AppendLine("       , C.DEALER_NM AS CVNAM   ");
            strSql.AppendLine("       , A.ACTCD     ");
            strSql.AppendLine("       , D.ACNAM AS ACTNM     ");
            strSql.AppendLine("       , A.ATEXT    ");
            strSql.AppendLine("       , A.ACAMT    ");
            strSql.AppendLine("       , A.ADAMT    ");
            strSql.AppendLine("       , A.ADPCD    ");
            strSql.AppendLine("       , E.DEPT_NM AS ADPNM     ");
            strSql.AppendLine("       , A.REF1    ");
            strSql.AppendLine("       , A.REF2    ");
            strSql.AppendLine("       , A.REF3    ");
            strSql.AppendLine("       , A.APVYN   ");
            strSql.AppendLine("       , A.AAUTO    ");
            strSql.AppendLine("       , A.ADATE    ");
            strSql.AppendLine("       , A.AUSER    ");
            strSql.AppendLine("       , F.BILNO    ");
            strSql.AppendLine("       , A.RK    ");
            strSql.AppendLine("       , A.CUSER AS CUSER    ");
            strSql.AppendLine("       , A.CDATE    ");
            strSql.AppendLine("       , A.MUSER AS MUSER    ");
            strSql.AppendLine("       , A.MDATE    ");
            strSql.AppendLine("   FROM SLIP_INFO XX  ");
            strSql.AppendLine("   LEFT JOIN ACTRAN A   ");
            strSql.AppendLine("     ON XX.TDATE = A.TDATE ");
            strSql.AppendLine("    AND XX.ATGUB = A.ATGUB ");
            strSql.AppendLine("    AND XX.SEQNO = A.SEQNO ");
            //strSql.AppendLine("    AND XX.LINNO = A.LINNO ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B    ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD    ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C     ");
            strSql.AppendLine("     ON A.CVCOD = C.DEALER_CD    ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF D    ");
            strSql.AppendLine("     ON A.ACTCD = D.ACCOD    ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD E    ");
            strSql.AppendLine("     ON A.ADPCD = E.DEPT_CD    ");
            strSql.AppendLine("   LEFT OUTER JOIN ACBILL F    ");
            strSql.AppendLine("     ON A.TDATE = F.TDATE    ");
            strSql.AppendLine("    AND A.ATGUB = F.ATGUB     ");
            strSql.AppendLine("    AND A.SEQNO = F.SEQNO     ");
            strSql.AppendLine("    AND A.LINNO = F.LINNO     ");
            strSql.AppendLine("  GROUP BY A.TDATE, A.ATGUB, A.SEQNO, A.LINNO ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetSlipInfo(string sYmdFrom, string sYmdTo, string AddingQuery, string sAprvGb)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT DATE_FORMAT(A.TDATE,'%Y-%m-%d') AS TDATE1    ");
            //strSql.AppendLine("      , A.ATGUB AS ATGUB1   ");
            //strSql.AppendLine("      , A.SEQNO AS SEQNO1   ");
            //strSql.AppendLine("      , A.LINNO AS LINNO1   ");
            //strSql.AppendLine("      , DATE_FORMAT(A.TDATE,'%Y-%m-%d') AS TDATE    ");
            //strSql.AppendLine("      , A.ATGUB   ");
            //strSql.AppendLine("      , A.SEQNO   ");
            //strSql.AppendLine("      , A.LINNO   ");
            //strSql.AppendLine("      , A.ACCOD   ");
            //strSql.AppendLine("      , B.ACNAM AS ACNAM   ");
            //strSql.AppendLine("      , CAST(A.CVCOD AS CHAR) AS CVCOD    ");
            //strSql.AppendLine("      , C.DEALER_NM AS CVNAM  ");
            //strSql.AppendLine("      , A.ACTCD    ");
            //strSql.AppendLine("      , D.ACNAM AS ACTNM    ");
            //strSql.AppendLine("      , A.ATEXT   ");
            //strSql.AppendLine("      , A.ACAMT   ");
            //strSql.AppendLine("      , A.ADAMT   ");
            //strSql.AppendLine("      , A.ADPCD   ");
            //strSql.AppendLine("      , E.DEPT_NM AS ADPNM    ");
            //strSql.AppendLine("      , A.REF1   ");
            //strSql.AppendLine("      , A.REF2   ");
            //strSql.AppendLine("      , A.REF3   ");
            //strSql.AppendLine("      , A.APVYN  ");
            //strSql.AppendLine("      , A.AAUTO   ");
            //strSql.AppendLine("      , A.ADATE   ");
            //strSql.AppendLine("      , A.AUSER   ");
            //strSql.AppendLine("      , F.BILNO   ");
            //strSql.AppendLine("      , A.RK   ");
            //strSql.AppendLine("      , CAST(A.CUSER AS CHAR) AS CUSER   ");
            //strSql.AppendLine("      , A.CDATE   ");
            //strSql.AppendLine("      , CAST(A.MUSER AS CHAR) AS MUSER   ");
            //strSql.AppendLine("      , A.MDATE   ");
            //strSql.AppendLine("   FROM ACTRAN A  ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B   ");
            //strSql.AppendLine("     ON A.ACCOD = B.ACCOD   ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C    ");
            //strSql.AppendLine("     ON A.CVCOD = C.DEALER_CD   ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF D   ");
            //strSql.AppendLine("     ON A.ACTCD = D.ACCOD   ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD E   ");
            //strSql.AppendLine("     ON A.ADPCD = E.DEPT_CD   ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACBILL F   ");
            //strSql.AppendLine("     ON A.TDATE = F.TDATE   ");
            //strSql.AppendLine("    AND A.ATGUB = F.ATGUB    ");
            //strSql.AppendLine("    AND A.SEQNO = F.SEQNO    ");
            //strSql.AppendLine("    AND A.LINNO = F.LINNO    ");
            //strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD G   ");
            //strSql.AppendLine("     ON A.ATGUB = G.COM_CD   ");
            //strSql.AppendLine("    AND G.CD_GB = 'AC01001_03'   ");
            #endregion

            strSql.AppendLine("SELECT CONVERT(VARCHAR(10),CONVERT(DATE, A.TDATE),23) AS TDATE1   ");
            strSql.AppendLine("     , A.ATGUB AS ATGUB1                                          ");
            strSql.AppendLine("     , A.SEQNO AS SEQNO1                                          ");
            strSql.AppendLine("     , A.LINNO AS LINNO1                                          ");
            strSql.AppendLine("     , CONVERT(VARCHAR(10), CONVERT(DATE, A.TDATE), 23) AS TDATE  ");
            strSql.AppendLine("     , A.ATGUB                                                    ");
            strSql.AppendLine("     , A.SEQNO                                                    ");
            strSql.AppendLine("     , A.LINNO                                                    ");
            strSql.AppendLine("     , A.ACCOD                                                    ");
            strSql.AppendLine("     , B.ACNAM AS ACNAM                                           ");
            strSql.AppendLine("     , A.CVCOD AS CVCOD                             ");
            strSql.AppendLine("     , C.DEALER_NM AS CVNAM                                       ");
            strSql.AppendLine("     , A.ACTCD                                                    ");
            strSql.AppendLine("     , D.ACNAM AS ACTNM                                           ");
            strSql.AppendLine("     , A.ATEXT                                                    ");
            strSql.AppendLine("     , A.ACAMT                                                    ");
            strSql.AppendLine("     , A.ADAMT                                                    ");
            strSql.AppendLine("     , A.ADPCD                                                    ");
            strSql.AppendLine("     , E.DEPT_NM AS ADPNM                                         ");
            strSql.AppendLine("     , A.REF1                                                     ");
            strSql.AppendLine("     , A.REF2                                                     ");
            strSql.AppendLine("     , A.REF3                                                     ");
            strSql.AppendLine("     , A.APVYN                                                    ");
            strSql.AppendLine("     , A.AAUTO                                                    ");
            strSql.AppendLine("     , A.ADATE                                                    ");
            strSql.AppendLine("     , A.AUSER                                                    ");
            strSql.AppendLine("     , F.BILNO                                                    ");
            //strSql.AppendLine("     , A.CVGB");
            strSql.AppendLine("     , A.RK                                                       ");
            strSql.AppendLine("     , A.CUSER AS CUSER                             ");
            strSql.AppendLine("     , A.CDATE                                                    ");
            strSql.AppendLine("     , A.MUSER AS MUSER                             ");
            strSql.AppendLine("     , A.MDATE                                                    ");
            strSql.AppendLine("  FROM ACTRAN A                                                   ");
            strSql.AppendLine("  LEFT OUTER JOIN ACMSTF B                                        ");
            strSql.AppendLine("    ON A.ACCOD = B.ACCOD                                          ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD C                                 ");
            strSql.AppendLine("    ON A.CVCOD = C.DEALER_CD                                      ");
            strSql.AppendLine("  LEFT OUTER JOIN ACMSTF D                                        ");
            strSql.AppendLine("    ON A.ACTCD = D.ACCOD                                          ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEPT_CD E                                   ");
            strSql.AppendLine("    ON A.ADPCD = E.DEPT_CD                                        ");
            strSql.AppendLine("  LEFT OUTER JOIN ACBILL F                                        ");
            strSql.AppendLine("    ON A.TDATE = F.TDATE                                          ");
            strSql.AppendLine("   AND A.ATGUB = F.ATGUB                                          ");
            strSql.AppendLine("   AND A.SEQNO = F.SEQNO                                          ");
            strSql.AppendLine("   AND A.LINNO = F.LINNO                                          ");
            strSql.AppendLine("  LEFT OUTER JOIN COM_BASE_CD G                                   ");
            strSql.AppendLine("    ON A.ATGUB = G.COM_CD                                         ");
            strSql.AppendLine("   AND G.CD_GB = 'AC01001_03'                                     ");
            strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            strSql.AppendLine("    AND (('" + sAprvGb + "' = '0' AND 1 = 1) ");
            strSql.AppendLine("         OR  ");
            strSql.AppendLine("         ('" + sAprvGb + "' = '1' AND A.APVYN = 'Y' )  ");
            strSql.AppendLine("         OR  ");
            strSql.AppendLine("         ('" + sAprvGb + "' = '2' AND A.APVYN <> 'Y' ))  ");
            strSql.AppendLine(AddingQuery);
            strSql.AppendLine("  ORDER BY A.TDATE, A.ATGUB, A.SEQNO, A.LINNO  ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private string ReturningByComboBoxValues(string sValue, string sText)
        {
            if (string.IsNullOrEmpty(sText))
                return string.Empty;

            string[] cnWHERE = { "A.ACNAM", "ATEXT", "C.DEALER_NM"}; string sWHERE = "";
            /*
             * 2021-02-07(현업요청)
             * 1. 거래처 초성검색 추가
             */
            if (cnWHERE[CboFindSbj.SelectedIndex].Equals("C.DEALER_NM"))
            {
                sWHERE = "                  AND (C.DEALER_NM LIKE '" + "%" + TxtFindWord.Text + "%" + "' OR C.INITIAL_NM LIKE '" + "%" + TxtFindWord.Text + "%" + "') ";
            }
            else if (TxtFindWord.Text != "")
            {
                sWHERE = "                  AND " + cnWHERE[CboFindSbj.SelectedIndex] + " LIKE '" + "%" + TxtFindWord.Text + "%" + "' ";
            }
            return sWHERE;

            //StringBuilder strSql = new StringBuilder();

            //if (sValue.Equals("계정코드"))
            //    strSql.AppendLine("    AND A.ACCOD LIKE '" + sText + "%' ");
            //else if (sValue.Equals("계정명"))
            //    strSql.AppendLine("    AND A.ACNAM LIKE '%" + sText + "%' ");
            //else if (sValue.Equals("성격"))
            //    strSql.AppendLine("    AND G.COM_NM LIKE '%" + sText + "%");
            //else if (sValue.Equals("관계계정코드"))
            //    strSql.AppendLine("    AND A.ASMCD LIKE '%" + sText + "%' ");
            //else if (sValue.Equals("관계계정명"))
            //{
            //    strSql.AppendLine("    AND B.ASMCD IN (SELECT X1.ACCOD ");
            //    strSql.AppendLine("                      FROM ACMSTF X1 ");
            //    strSql.AppendLine("                     WHERE X1.ACNAM LIKE '%" + sText + "%' ) ");
            //}

            //return strSql.ToString();
        }

        private void DeleteSlipInfo(DataRow drSlip)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sTDate = drSlip["TDATE"]?.ToString().Replace("-", "").Substring(0, 8);
                string sAtGub = drSlip["ATGUB"]?.ToString();
                string sSeqNo = drSlip["SEQNO"]?.ToString();
                string sAText = drSlip["ATEXT"]?.ToString();
                string sAuto = drSlip["AAUTO"]?.ToString();
                string sRef1 = drSlip["REF1"]?.ToString();
                string sRef2 = drSlip["REF2"]?.ToString();
                string sRef3 = drSlip["REF3"]?.ToString();

                StringBuilder strSql = new StringBuilder();

                //strSql.Clear();
                //strSql.AppendLine(" UPDATE ACC_TAX_MGT A  ");
                //strSql.AppendLine("  INNER JOIN ACTRAN B  ");
                //strSql.AppendLine("     ON A.BILL_KEY = B.REF1 ");
                //strSql.AppendLine("    SET A.SLIP_ISSUE_YN = 'N' ");
                //strSql.AppendLine("      , A.SLIP_YMD = '' ");
                //strSql.AppendLine("      , A.SLIP_NO = '' ");
                //strSql.AppendLine("  WHERE B.REF1 = '" + sRef1 + "' ");

                //cmd.CommandType = CommandType.Text;
                //cmd.CommandText = strSql.ToString();
                //cmd.ExecuteNonQuery();


                /*
                 * 2021-03-10
                 * Reference Key : #0001
                 * 로그 적용
                 */
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CONVERT(VARCHAR(10), CONVERT(DATE,X1.TDATE), 23) AS TDATE ");
                strSql.AppendLine("      , X1.ATGUB ");
                strSql.AppendLine("      , X2.COM_NM AS ATGUB_NM ");
                strSql.AppendLine("      , X1.SEQNO ");
                strSql.AppendLine("      , X1.LINNO ");
                strSql.AppendLine("      , X1.ACCOD ");
                strSql.AppendLine("      , X1.ACNAM ");
                strSql.AppendLine("      , X1.CVCOD ");
                strSql.AppendLine("      , X1.CVNAM ");
                strSql.AppendLine("      , X1.ATEXT ");
                strSql.AppendLine("      , X1.ACAMT ");
                strSql.AppendLine("      , X1.ADAMT ");
                strSql.AppendLine("      , X1.ADPCD ");
                strSql.AppendLine("      , X1.ADPNM ");
                strSql.AppendLine("   FROM ACTRAN X1 ");
                strSql.AppendLine("   LEFT JOIN COM_BASE_CD X2 ");
                strSql.AppendLine("     ON X1.ATGUB = X2.COM_CD ");
                strSql.AppendLine("    AND X2.CD_GB = 'AC02001_01' ");
                strSql.AppendLine("  WHERE TDATE = '" + sTDate + "'");
                strSql.AppendLine("    AND ATGUB = '" + sAtGub + "'");
                strSql.AppendLine("    AND SEQNO = '" + sSeqNo + "'");

                DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                for(int i = 0; i< dtPrv.Rows.Count; i++)
                {
                    string sLog_Msg = string.Empty;
                    string sPrv_TDate = dtPrv.Rows[i]["TDATE"]?.ToString();
                    string sPrv_AtGub = dtPrv.Rows[i]["ATGUB"]?.ToString();
                    string sPrv_AtGub_Nm = dtPrv.Rows[i]["ATGUB_NM"]?.ToString();
                    string sPrv_SeqNo = dtPrv.Rows[i]["SEQNO"]?.ToString();
                    string sPrv_LinNo = dtPrv.Rows[i]["LINNO"]?.ToString();

                    string sPrv_AcCod = dtPrv.Rows[i]["ACCOD"]?.ToString();
                    string sPrv_AcNam = dtPrv.Rows[i]["ACNAM"]?.ToString();
                    string sPrv_CvCod = dtPrv.Rows[i]["CVCOD"]?.ToString();
                    string sPrv_CvNam = dtPrv.Rows[i]["CVNAM"]?.ToString();
                    string sPrv_AText = dtPrv.Rows[i]["ATEXT"]?.ToString();

                    double dPrv_AcAmt = 0;
                    double.TryParse(dtPrv.Rows[i]["ACAMT"]?.ToString(), out dPrv_AcAmt);

                    double dPrv_AdAmt = 0;
                    double.TryParse(dtPrv.Rows[i]["ADAMT"]?.ToString(), out dPrv_AdAmt);

                    string sPrv_AdpCd = dtPrv.Rows[i]["ADPCD"]?.ToString();
                    string sPrv_AdpNm = dtPrv.Rows[i]["ADPNM"]?.ToString();

                    /*
                     * #0002
                     */
                    string sSTD_COLS = string.Empty;
                    string sREF_RMK = string.Empty;

                    sSTD_COLS = string.Format("{0}/{1}전표번호:{2}/순번:{3}"
                        , sPrv_TDate, sPrv_AtGub, sPrv_SeqNo, sPrv_LinNo);

                    sLog_Msg += string.Format("계정:{0} ", sPrv_AcNam);
                    sLog_Msg += string.Format("/거래처:{0}", sPrv_CvNam);
                    sLog_Msg += string.Format("/차변:{0}", dPrv_AcAmt);
                    sLog_Msg += string.Format("/대변:{0}", dPrv_AdAmt);

                    sREF_RMK += string.Format("TABLE : ACTRAN / TDATE : {0}, ATGUB : {1}, SEQNO : {2}, LINNO : {3} ", sPrv_TDate.Replace("-", ""), sPrv_AtGub, sPrv_SeqNo, sPrv_LinNo);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine("           ( OCCUR_DATE ");
                    strSql.AppendLine("           , USRCD ");
                    strSql.AppendLine("           , LOG_SEQ ");
                    strSql.AppendLine("           , EDIT_KIND ");
                    strSql.AppendLine("           , PGM_ID ");
                    strSql.AppendLine("           , ACS_IP ");
                    strSql.AppendLine("           , STD_COLS ");
                    strSql.AppendLine("           , REF_RMK ");
                    strSql.AppendLine("           , EDIT_RMK ) ");
                    strSql.AppendLine("     VALUES( @OCCUR_DATE ");
                    strSql.AppendLine("           , @USRCD ");
                    strSql.AppendLine("           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                    strSql.AppendLine("           , 'D' ");
                    strSql.AppendLine("           , @PGM_ID ");
                    strSql.AppendLine("           , @ACS_IP ");
                    strSql.AppendLine("           , @STD_COLS ");
                    strSql.AppendLine("           , @REF_RMK ");
                    strSql.AppendLine("           , @EDIT_RMK ) ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                    cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                    cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                    cmd.Parameters.AddWithValue("@STD_COLS", sSTD_COLS);
                    cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                    cmd.Parameters.AddWithValue("@EDIT_RMK", sLog_Msg = sLog_Msg.Length > 500 ? sLog_Msg.Substring(0, 500) : sLog_Msg);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
                
                strSql.Clear();
                strSql.AppendLine(" DELETE FROM ACTRAN ");
                strSql.AppendLine("  WHERE TDATE = '" + sTDate + "' ");
                strSql.AppendLine("    AND ATGUB = '" + sAtGub + "' ");
                strSql.AppendLine("    AND SEQNO = '" + sSeqNo + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                if (!string.IsNullOrEmpty(sAuto))
                {
                    if (sAuto.Equals("A01")) //자동전표 여부 -> 매입출 (계근 데이터)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = NULL ");
                        strSql.AppendLine("      , APRV_DATE = NULL ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", Convert.ToDouble(sRef1));
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        strSql.Clear();
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = 'N' ");
                        strSql.AppendLine("  WHERE ISNULL(JUNPYOID, 0) = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.AddWithValue("@JUNPYOID", Convert.ToDouble(sRef1));
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        //string sRmk = string.Format("[자동전표삭제]JUNPYOID : {0}, 적요 : {1} ", sRef1, sAText);

                        //strSql.Clear();
                        //strSql.AppendLine(" CALL DP_IST_LOG (NOW(), " + FmMainToolBar2.UserID + ", 'U', 'INLIST', '" + sRmk + "'); ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();
                    }
                    else if (sAuto.Equals("A02")) //자동전표 여부 -> 직송
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = NULL ");
                        strSql.AppendLine("      , APRV_DATE = NULL ");
                        strSql.AppendLine("  WHERE J_RID = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", Convert.ToDouble(sRef3));
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        strSql.Clear();
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET APRV_YN = 'N' ");
                        strSql.AppendLine("  WHERE ISNULL(J_RID, 0) = @JUNPYOID ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.AddWithValue("@JUNPYOID", Convert.ToDouble(sRef3));
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        //string sRmk = string.Format("[자동전표삭제][직송]JUNPYOID : {0}, 적요 : {1} ", sRef3, sAText);

                        //strSql.Clear();
                        //strSql.AppendLine(" CALL DP_IST_LOG (NOW(), " + FmMainToolBar2.UserID + ", 'U', 'INLIST', '" + sRmk + "'); ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();
                    }
                    else if (sAuto.Equals("D01")) //자동전표 구분 -> 지불
                    {
                        strSql.Clear();
                        strSql.AppendLine(" DELETE FROM SUGMF ");
                        strSql.AppendLine("       WHERE SLIPNO = @SLIPNO ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.AddWithValue("@SLIPNO", sRef1);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();

                        string sRmk = string.Format("[자동전표삭제][직송]JUNPYOID : {0}, 적요 : {1} ", sRef1, sAText);

                        //strSql.Clear();
                        //strSql.AppendLine(" CALL DP_IST_LOG (NOW(), " + FmMainToolBar2.UserID + ", 'U', 'INLIST', '" + sRmk + "'); ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();
                    }
                }

                //출처 일반비용일 경우 해당 설비이력 전표여부 해제
                if (sRef3.Equals("MAKE_EXPENSE"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" UPDATE EQUIP_CD_HISTORY  ");
                    strSql.AppendLine("    SET SLIP_YN = 'N' ");
                    strSql.AppendLine("  WHERE MG_NO = '" + sRef1 + "' ");
                    strSql.AppendLine("    AND MG_HIS_SEQ = " + sRef2 + " ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                //어음삭제
                strSql.Clear();
                strSql.AppendLine(" DELETE FROM ACBILL ");
                strSql.AppendLine("  WHERE TDATE = '" + sTDate + "' ");
                strSql.AppendLine("    AND ATGUB = '" + sAtGub + "' ");
                strSql.AppendLine("    AND SEQNO = '" + sSeqNo + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료하였습니다.");

                BtnRetr_Click(null, null);
                //SetFocusedRow(sTDate, sAtGub, sSeqNo);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }


        }

        private DataTable GetRemarkInfo()
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("               SELECT A.ATEXT AS CD ");
            strSql.AppendLine("                 FROM ACTRAN A ");
            strSql.AppendLine("                WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            strSql.AppendLine("                GROUP BY A.ATEXT ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #endregion[Execute By Query]

        #region[기타 메소드]

        //삭제 후 삭제된 정보의 전표번호가 존재하는 데이터에 포커스
        private void SetFocusedRow(string sTDate, string sAtGub, string sSeqNo)
        {
            for (int i = 0; i < GridViewSlip.RowCount; i++)
            {
                string sSlipYmd = GridViewSlip.GetDataRow(i)["TDATE"]?.ToString().Replace("-", "").Substring(0, 8);
                string sSlipGb = GridViewSlip.GetDataRow(i)["ATGUB"]?.ToString();
                string sSlipNo = GridViewSlip.GetDataRow(i)["SEQNO"]?.ToString();

                if (sSlipYmd.Equals(sTDate) && sSlipGb.Equals(sAtGub) && sSlipNo.Equals(sSeqNo))
                {
                    GridViewSlip.FocusedRowHandle = i;
                    break;
                }
            }
        }

        private void SettingTextEditAutoComplete(string sGb)
        {
            AutoCompleteStringCollection collection = new AutoCompleteStringCollection();

            DataTable dt = new DataTable();

            if (sGb.Equals("계정코드") || sGb.Equals("관계계정코드"))
            {
                dt = GetLookUpData("2", "Y", "Y");
            }
            if (sGb.Equals("계정명") || sGb.Equals("관계계정명"))
            {
                dt = GetLookUpData("3", "Y", "Y");
            }
            if (sGb.Equals("적요"))
            {
                dt = GetRemarkInfo();
            }
            if (sGb.Equals("자동전표구분"))
            {
                dt = GetLookUpData("6", "Y", "Y");
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                collection.Add(dt.Rows[i]["CD"]?.ToString());
            }

            TxtFindWord.MaskBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            TxtFindWord.MaskBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            TxtFindWord.MaskBox.AutoCompleteCustomSource = collection;
        }

        #endregion[기타 메소드]

        #region[GridView Row's Design]

        private void GridViewSlip_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);

            //if (_MODE == SelectionMode.Single)
            //{
            //    ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
            //}
            //else if (_MODE == SelectionMode.Multi)
            //{
            //    string sApvYn = GridViewSlip.GetRowCellValue(e.RowHandle, GridColApvYn)?.ToString();
            //    if (sApvYn == null || sApvYn.Equals("N"))
            //    {
            //        e.Appearance.BackColor = SystemColors.Info;
            //    }
            //    else
            //    {
            //        e.Appearance.BackColor = Color.Gray;
            //    }
            //}
        }

        private void GridViewSlip_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
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
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_01'");
            }
            else if (sGb.Equals("2")) //계정코드
            {
                strSql.AppendLine(" SELECT A.ACCOD AS CD ");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , '' AS SEQ");
                strSql.AppendLine("   FROM ACMSTF A ");
            }
            else if (sGb.Equals("3")) //계정명
            {
                strSql.AppendLine(" SELECT A.ACNAM AS CD ");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , '' AS SEQ");
                strSql.AppendLine("   FROM ACMSTF A ");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_06'");
            }
            else if (sGb.Equals("5"))
            {
                #region mariaDB
                //strSql.AppendLine(" SELECT EMP_ID ");
                //strSql.AppendLine("      , EMP_NM ");
                //strSql.AppendLine("      , @SEQ := @SEQ + 1 AS SEQ ");
                //strSql.AppendLine("   FROM HR_EMP_BASIS A, (SELECT @SEQ := 0) A ");
                #endregion

                strSql.AppendLine("SELECT EMP_ID                                     ");
                strSql.AppendLine("     , EMP_NM                                     ");
                strSql.AppendLine("     , ROW_NUMBER () OVER(ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("  FROM HR_EMP_BASIS A                             ");
            }
            else if (sGb.Equals("6"))
            {
                strSql.AppendLine(" SELECT A.COM_NM AS CD ");
                strSql.AppendLine("      , '' AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_06'");
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

        private void CboFindSbj_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                ComboBoxEdit cbo = (ComboBoxEdit)sender;
                SettingTextEditAutoComplete(cbo.EditValue?.ToString());
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
            }
        }

        //private enum SelectionMode { Multi, Single }
        //private SelectionMode _MODE = SelectionMode.Single;
        //private enum ApplyMode { Apply, Cancle }
        //private ApplyMode _APY_MODE;

        private void BtnAprv_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sSlipYmd = GridViewSlip.GetFocusedRowCellValue("TDATE")?.ToString().Replace("-", "").Substring(0, 8);
            string sSlipGb_Text = GridViewSlip.GetFocusedRowCellDisplayText("ATGUB")?.ToString();
            string sSeqNo = GridViewSlip.GetFocusedRowCellValue("SEQNO")?.ToString();
            string sSlipGb = GridViewSlip.GetFocusedRowCellValue("ATGUB")?.ToString();
            string sAuto = GridViewSlip.GetFocusedRowCellValue("AAUTO")?.ToString();

            //_MODE = SelectionMode.Multi;
            //_APY_MODE = ApplyMode.Apply;

            //GridViewSlip.OptionsSelection.MultiSelect = true;
            //GridViewSlip.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
            //GridViewSlip.OptionsSelection.CheckBoxSelectorColumnWidth = 40;

            if (!string.IsNullOrEmpty(sAuto))
            {
                if (sAuto.Equals("A01") || sAuto.Equals("A02"))
                {
                    XtraMessageBox.Show("자동전표는 승인취소 할 수 없습니다.");
                    return;
                }
            }

            string sSlipInfo = "전표일자 : " + sSlipYmd
                + "\r\n" + "전표구분 : " + sSlipGb_Text
                + "\r\n" + "전표번호 : " + sSeqNo;

            string sText = BtnAprv.Text;
            if (sText.Equals("전표승인"))
            {
                if (XtraMessageBox.Show(sSlipInfo + "\r\n\r\n" + "선택하신 정보가 승인처리됩니다.\r\n정말로 진행하시겠습니까?"
                , "전표승인여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    ApproveSlipData(sSlipYmd, sSlipGb, sSeqNo, "Y");
                }
            }
            else if (sText.Equals("승인취소"))
            {
                if (XtraMessageBox.Show(sSlipInfo + "\r\n\r\n" + "선택하신 정보가 승인취소처리됩니다.\r\n정말로 진행하시겠습니까?"
                , "승인취소여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    ApproveSlipData(sSlipYmd, sSlipGb, sSeqNo, "N");
                }
            }
        }

        private void CheckGridView(GridView view)
        {
            for (int i = 0; i < view.RowCount; i++)
            {
                string sApvYn = view.GetRowCellValue(i, GridColApvYn)?.ToString();

            }
        }

        private void BtnCopySlip_Click(object sender, EventArgs e)
        {

        }

        private void GridViewSlip_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SettingButtonApproval(GridViewSlip.GetFocusedRowCellValue("APVYN")?.ToString());
        }

        /// <summary>
        ///     APVYN 값에 따라 BtnAprv 텍스트 Toggle 형식으로 바꾸기
        /// </summary>
        /// <param name="sAprvYn">ACTRAN / APVYN</param>
        private void SettingButtonApproval(string sAprvYn)
        {
            if (!string.IsNullOrEmpty(sAprvYn))
            {
                if (sAprvYn.Equals("Y"))
                {
                    BtnAprv.Text = "승인취소";
                }
                else
                {
                    BtnAprv.Text = "전표승인";
                }
            }
            else
            {
                BtnAprv.Text = "전표승인";
            }
        }

        /// <summary>
        ///     전표승인/취소 처리
        /// </summary>
        /// <param name="sTDate">전표일자</param>
        /// <param name="sATgub">전표구분</param>
        /// <param name="sSeqNo">전표번호</param>
        /// <param name="sAprvGB">승인/취소 구분(Y : 승인처리, N: 취소처리)</param>
        private void ApproveSlipData(string sTDate, string sATgub, string sSeqNo, string sAprvGB)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" UPDATE ACTRAN  ");
                strSql.AppendLine("    SET APVYN = @APVYN ");
                strSql.AppendLine("      , ADATE = @ADATE ");
                strSql.AppendLine("      , AUSER = @AUSER ");
                strSql.AppendLine("      , MUSER = @MUSER ");
                strSql.AppendLine("      , MDATE = CONVERT(VARCHAR(19),GETDATE(),20)");
                strSql.AppendLine("  WHERE TDATE = @TDATE ");
                strSql.AppendLine("    AND ATGUB = @ATGUB ");
                strSql.AppendLine("    AND SEQNO = @SEQNO ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.AddWithValue("@APVYN", sAprvGB);
                if (sAprvGB.Equals("Y"))
                {
                    cmd.Parameters.AddWithValue("@ADATE", DateTime.Now);
                    cmd.Parameters.AddWithValue("@AUSER", FmMainToolBar2.UserID);
                }
                else if (sAprvGB.Equals("N"))
                {
                    cmd.Parameters.AddWithValue("@ADATE", "");
                    cmd.Parameters.AddWithValue("@AUSER", "");
                }

                cmd.Parameters.AddWithValue("@MUSER", FmMainToolBar2.UserID);
                cmd.Parameters.AddWithValue("@TDATE", sTDate);
                cmd.Parameters.AddWithValue("@ATGUB", sATgub);
                cmd.Parameters.AddWithValue("@SEQNO", sSeqNo);
                cmd.ExecuteNonQuery();

                /*
                 * 2021-03-22
                 * #0003
                 * 승인취소일 경우에만 Log적용
                 */
                if (sAprvGB.Equals("N"))
                {
                    /*
                     * 2021-03-10
                     * #0001
                     * Log적용
                     */
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT CONVERT(VARCHAR(10),CONVERT(DATE,X1.TDATE), 23) AS TDATE ");
                    strSql.AppendLine("      , X1.ATGUB ");
                    strSql.AppendLine("      , X2.COM_NM AS ATGUB_NM ");
                    strSql.AppendLine("      , X1.SEQNO ");
                    strSql.AppendLine("   FROM ACTRAN X1 ");
                    strSql.AppendLine("   LEFT JOIN COM_BASE_CD X2 ");
                    strSql.AppendLine("     ON X1.ATGUB = X2.COM_CD ");
                    strSql.AppendLine("    AND X2.CD_GB = 'AC02001_01' ");
                    strSql.AppendLine("  WHERE TDATE = '" + sTDate + "'");
                    strSql.AppendLine("    AND ATGUB = '" + sATgub + "'");
                    strSql.AppendLine("    AND SEQNO = '" + sSeqNo + "'");

                    DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    string sPrv_TDate = dtPrv.Rows[0]["TDATE"]?.ToString();
                    string sPrv_AtGub = dtPrv.Rows[0]["ATGUB"]?.ToString();
                    string sPrv_AtGub_Nm = dtPrv.Rows[0]["ATGUB_NM"]?.ToString();
                    string sPrv_SeqNo = dtPrv.Rows[0]["SEQNO"]?.ToString();

                    string sSTD_COLS = string.Empty;
                    string sREF_RMK = string.Empty;
                    string sLog_Msg = string.Empty;
                    sSTD_COLS = string.Format("{0}/{1}/전표번호:{2}"
                                    , sPrv_TDate
                                    , sPrv_AtGub_Nm
                                    , sPrv_SeqNo);

                    sLog_Msg += string.Format("[전표{0}]", sAprvGB.Equals("Y") ? "승인" : "취소");
                    sREF_RMK += string.Format("TABLE : ACTRAN / TDATE : {0}, ATGUB : {1}, SEQNO : {2}", sPrv_TDate, sPrv_AtGub, sPrv_SeqNo);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine("           ( OCCUR_DATE ");
                    strSql.AppendLine("           , USRCD ");
                    strSql.AppendLine("           , LOG_SEQ ");
                    strSql.AppendLine("           , EDIT_KIND ");
                    strSql.AppendLine("           , PGM_ID ");
                    strSql.AppendLine("           , ACS_IP ");
                    strSql.AppendLine("           , STD_COLS ");
                    strSql.AppendLine("           , REF_RMK ");
                    strSql.AppendLine("           , EDIT_RMK ) ");
                    strSql.AppendLine("     VALUES( @OCCUR_DATE ");
                    strSql.AppendLine("           , @USRCD ");
                    strSql.AppendLine("           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                    strSql.AppendLine("           , 'U' ");
                    strSql.AppendLine("           , @PGM_ID ");
                    strSql.AppendLine("           , @ACS_IP ");
                    strSql.AppendLine("           , @STD_COLS ");
                    strSql.AppendLine("           , @REF_RMK ");
                    strSql.AppendLine("           , @EDIT_RMK ) ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                    cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                    cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                    cmd.Parameters.AddWithValue("@STD_COLS", sSTD_COLS);
                    cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                    cmd.Parameters.AddWithValue("@EDIT_RMK", sLog_Msg);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                string sGb = sAprvGB.Equals("Y") ? "승인" : "취소";
                XtraMessageBox.Show(string.Format("{0}처리를 완료하였습니다.", sGb));

                BtnRetr_Click(null, null);
                SettingButtonApproval("N");
                SetFocusedRow(sTDate, sATgub, sSeqNo);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void AC02001F01_TextChanged(object sender, EventArgs e)
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

        private void RdgbAprvGb_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void DropBtnSlip_Click(object sender, EventArgs e)
        {
            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "전표복사")
            {
                SlipCopy();
            }
            else if (tag == "전표이동")
            {
                SlipMove();
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnSlipCopy;
        BarButtonItem BtnSlipMove;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnSlipCopy = new BarButtonItem(barManager1, "전표복사");
            BtnSlipMove = new BarButtonItem(barManager1, "전표이동");
            popupMenu1.AddItem(BtnSlipCopy);
            popupMenu1.AddItem(BtnSlipMove);

            DropBtnSlip.DropDownControl = popupMenu1;

            BtnSlipCopy.Tag = "전표복사";
            BtnSlipCopy.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlipCopy_ItemClick);

            BtnSlipMove.Tag = "전표이동";
            BtnSlipMove.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlipMove_ItemClick);
        }

        private void BtnSlipCopy_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            SlipCopy();
        }

        private void BtnSlipMove_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            SlipMove();
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnSlip.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnSlip.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnSlip.Tag = submenuItem.Tag;
            DropBtnSlip.Text = string.Format("{0}", submenuItem.Tag);
        }

        private void SlipCopy()
        {
            if (GridViewSlip.RowCount == 0)
            {
                XtraMessageBox.Show("전표리스트에 데이터가 존재하지 않습니다.");
                return;
            }

            //자동전표일 경우 전표복사 할 수 없음
            string sAAuto = GridViewSlip.GetFocusedRowCellValue(GridColAAuto)?.ToString().Trim();
            if (!string.IsNullOrEmpty(sAAuto))
            {
                XtraMessageBox.Show("자동전표는 복사할 수 없습니다.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CONVERT(VARCHAR(10),CONVERT(DATE,A.TDATE),23) AS TDATE ");
            strSql.AppendLine("  	 , A.ATGUB ");
            strSql.AppendLine("  	 , A.SEQNO ");
            strSql.AppendLine("  	 , A.LINNO ");
            strSql.AppendLine("  	 , A.ACCOD ");
            strSql.AppendLine("  	 , A.ACNAM AS ACNAM ");
            strSql.AppendLine("  	 , A.CVCOD ");
            strSql.AppendLine("  	 , A.CVNAM  ");
            strSql.AppendLine("  	 , A.ATEXT ");
            strSql.AppendLine("  	 , A.ACAMT ");
            strSql.AppendLine("  	 , A.ADAMT  ");
            strSql.AppendLine("      , A.ADPCD ");
            strSql.AppendLine("      , A.ADPNM ");
            //strSql.AppendLine("      , A.CVGB");
            strSql.AppendLine("  	 , A.RK ");
            strSql.AppendLine("   FROM ACTRAN A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C ");
            strSql.AppendLine("     ON A.CVCOD = C.DEALER_CD ");
            strSql.AppendLine("  WHERE A.TDATE = REPLACE('" + GridViewSlip.GetFocusedRowCellValue(GridColTDate)?.ToString() + "', '-', '') ");
            strSql.AppendLine("    AND A.ATGUB = '" + GridViewSlip.GetFocusedRowCellValue(GridColAtGub)?.ToString() + "' ");
            strSql.AppendLine("    AND A.SEQNO = '" + GridViewSlip.GetFocusedRowCellValue(GridColSeqNo)?.ToString() + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            AC02001F02 frm = new AC02001F02();
            frm.PARENT_FORM = this;
            frm.AddModifyGb = "ADD";
            frm.COPYGB = "1";
            frm.SLIPDT = GridViewSlip.GetFocusedRowCellValue(GridColTDate)?.ToString();
            frm.SLIPGB = GridViewSlip.GetFocusedRowCellValue(GridColAtGub)?.ToString();
            frm.DataRowSendEvent += new AC02001F02.SendDataHandler(GetDataRow);
            frm.DT_COPY = dt;
            frm.Show();
            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //    BtnRetr.PerformClick();
            //}
        }

        private void SlipMove()
        {
            if (GridViewSlip.RowCount == 0)
            {
                XtraMessageBox.Show("전표리스트에 데이터가 존재하지 않습니다.");
                return;
            }

            string sAprvYn = GridViewSlip.GetFocusedRowCellValue(GridColApvYn)?.ToString().Trim();
            if (sAprvYn.Equals("Y"))
            {
                XtraMessageBox.Show("해당 전표 건은 전표승인 상태입니다.");
                return;
            }

            //자동전표일 경우 전표복사 할 수 없음
            string sAAuto = GridViewSlip.GetFocusedRowCellValue(GridColAAuto)?.ToString().Trim();
            if (!string.IsNullOrEmpty(sAAuto))
            {
                XtraMessageBox.Show("자동전표는 복사할 수 없습니다.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CONVERT(VARCHAR(10),CONVERT(DATE,A.TDATE),23) AS TDATE ");
            strSql.AppendLine("  	 , A.ATGUB ");
            strSql.AppendLine("  	 , A.SEQNO ");
            strSql.AppendLine("  	 , A.LINNO ");
            strSql.AppendLine("  	 , A.ACCOD ");
            strSql.AppendLine("  	 , A.ACNAM AS ACNAM ");
            strSql.AppendLine("  	 , A.CVCOD ");
            strSql.AppendLine("  	 , A.CVNAM  ");
            strSql.AppendLine("  	 , A.ATEXT ");
            strSql.AppendLine("  	 , A.ACAMT ");
            strSql.AppendLine("  	 , A.ADAMT  ");
            strSql.AppendLine("      , A.ADPCD ");
            strSql.AppendLine("      , A.ADPNM ");
            //strSql.AppendLine("      , A.CVGB");
            strSql.AppendLine("  	 , A.RK ");
            strSql.AppendLine("   FROM ACTRAN A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C ");
            strSql.AppendLine("     ON A.CVCOD = C.DEALER_CD ");
            strSql.AppendLine("  WHERE A.TDATE = REPLACE('" + GridViewSlip.GetFocusedRowCellValue(GridColTDate)?.ToString() + "', '-', '') ");
            strSql.AppendLine("    AND A.ATGUB = '" + GridViewSlip.GetFocusedRowCellValue(GridColAtGub)?.ToString() + "' ");
            strSql.AppendLine("    AND A.SEQNO = '" + GridViewSlip.GetFocusedRowCellValue(GridColSeqNo)?.ToString() + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            AC02001F02 frm = new AC02001F02();
            frm.PARENT_FORM = this;
            frm.AddModifyGb = "ADD";
            frm.COPYGB = "2";
            frm.SLIPDT = GridViewSlip.GetFocusedRowCellValue(GridColTDate)?.ToString();
            frm.SLIPGB = GridViewSlip.GetFocusedRowCellValue(GridColAtGub)?.ToString();
            frm.DT_COPY = dt;
            frm.DataRowSendEvent += new AC02001F02.SendDataHandler(GetDataRow);
            frm.Show();
            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //    BtnRetr.PerformClick();
            //}
        }


        BarManager barManager2;
        PopupMenu popupMenu2;
        BarButtonItem BtnDelOne;
        BarButtonItem BtnDelOverOne;
        private void InitControls2()
        {
            barManager2 = new BarManager();
            barManager2.Form = this;

            popupMenu2 = new PopupMenu(barManager2);
            BtnDelOne = new BarButtonItem(barManager2, "건별삭제");
            BtnDelOverOne = new BarButtonItem(barManager2, "선택삭제");
            popupMenu2.AddItem(BtnDelOne);
            popupMenu2.AddItem(BtnDelOverOne);

            DropBtnDel.DropDownControl = popupMenu1;

            BtnDelOne.Tag = "건별삭제";
            BtnDelOne.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnDelOne_ItemClick);

            BtnDelOverOne.Tag = "선택삭제";
            BtnDelOverOne.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnDelOverOne_ItemClick);
        }

        private void BtnDelOne_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            DeleteOne();
        }

        private void BtnDelOverOne_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton2(e.Item);
            //...
            DeleteOneOver();
        }

        private void UpdateDropDownButton2(BarItem submenuItem)
        {
            DropBtnDel.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnDel.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnDel.Tag = submenuItem.Tag;
            DropBtnDel.Text = string.Format("{0}", submenuItem.Tag);
        }

        private void DeleteOne()
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (GridViewSlip.RowCount == 0)
            {
                XtraMessageBox.Show("전표리스트에 데이터가 존재하지 않습니다.");
                return;
            }

            string sSlipYmd = GridViewSlip.GetFocusedRowCellValue("TDATE")?.ToString();
            string sSlipGb = GridViewSlip.GetFocusedRowCellDisplayText("ATGUB")?.ToString();
            string sSeqNo = GridViewSlip.GetFocusedRowCellValue("SEQNO")?.ToString();

            string sSlipInfo = "전표일자 : " + sSlipYmd
                + "\r\n" + "전표구분 : " + sSlipGb
                + "\r\n" + "전표번호 : " + sSeqNo;

            if (XtraMessageBox.Show(sSlipInfo + "\r\n\r\n" + "선택하신 정보가 삭제됩니다.\r\n정말로 진행하시겠습니까?"
                , "삭제여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DeleteSlipInfo(GridViewSlip.GetFocusedDataRow());
            }
        }

        private void DeleteOneOver()
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditFrom.EditValue?.ToString().Substring(0, 10);


        }

        private void DropBtnDel_Click(object sender, EventArgs e)
        {
            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "건별삭제")
            {
                DeleteOne();
            }
            else if (tag == "선택삭제")
            {
                DeleteOneOver();
            }
        }

        private void BtnCancle_Click(object sender, EventArgs e)
        {
            //_MODE = SelectionMode.Multi;
            //_APY_MODE = ApplyMode.Cancle;

            //GridViewSlip.OptionsSelection.MultiSelect = true;
            //GridViewSlip.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
            //GridViewSlip.OptionsSelection.CheckBoxSelectorColumnWidth = 40;
        }
    }
}