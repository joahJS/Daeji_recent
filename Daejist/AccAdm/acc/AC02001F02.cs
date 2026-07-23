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
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using System.Resources;
using AccAdm.Properties;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Collections;
using System.Data.SqlClient;

/*
 * 작성일자 : 2020-02월 초
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-01-27
 * 수정자 : 고혜성
 * 수정내용 : 입출금전표 저장 시 현금계정을 상대계정으로 넣는 것을 다시 원래대로 세팅
 * 저장로직 살펴볼 것 
 * 
 * 수정일자 : 2021-02-07
 * 수정자 : 고혜성
 * 수정내용 : 거래처초성검색 추가 (Repository 이벤트 관련 참조)
 * 
 * 수정일자 : 2021-03-10
 * 수정자   : 고혜성
 * Reference Key = #0001
 * 수정내용 : 저장 시 로그적용
 * 
 * 수정일자 : 2021-03-17
 * 수정자   : 고혜성
 * Reference Key = #0002
 * 수정내용 : (현업요청)
 *            1. 로그수정
 *               1-1) 기본사항/변경사항을 나누어 입력
 *               
 * 수정일자 : 2021-03-18
 * 수정자   : 고혜성
 * 수정내용 : (현업요청) #0002와 관련
 *            1. 기본항목에 입력되는 값 단순화
 */
namespace AccAdm
{
    public partial class AC02001F02 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F02()
        {
            InitializeComponent();
        }

        /*
         * 2021-01-06 현업요청
         * 지불관리에서 등록된 건에 대하여 전표관리에서 수정가능하도록 하며
         * 지불관리 SUGMF의 데이터와 동기화
         */
        public enum SYNC_GB { None, Maipchul, JikSong, Jibul }
        public SYNC_GB _SYNC = SYNC_GB.None; //초기 디폴트는 None 세팅

        /*
         * 2021-03-04 (현업요청)
         * 자동전표 및 승인전표의 경우에도 전표창은 뜨되 수정을 할수없도록 수정
         */
        public enum EDIT_YN { Editable, NonEditable }
        public EDIT_YN _EDIT_YN = EDIT_YN.Editable;
        public enum EDIT_GB { Add, Auto, Approval }
        public EDIT_GB _EDIT_GB = EDIT_GB.Add;

        public string AddModifyGb { get; set; }
        public string COPYGB = string.Empty; //1: 전표복사, 2: 전표이동
        public string SLIPDT = string.Empty;
        public string SLIPGB = string.Empty;
        public DataTable DT_COPY;

        public Dictionary<string, string> ExternalParam;
        public DataRow DrSlipInfo { get; set; }
        public Form PARENT_FORM; //전표관리

        public delegate void SendDataHandler(Dictionary<string, string> sGb);
        public event SendDataHandler DataRowSendEvent;

        public DataRow rowUserInfo { get; set; }
        private void AC02001F02_Load(object sender, EventArgs e)
        {
            DataTable dtSlipType = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupSlipType, dtSlipType, "CD", "NM", "Y");

            /*
             * 2021-03-04
             * 수정 불가능한 전표(승인전표, 자동전표)의 경우 수정할 수 없다는 메시지 출력
             */
            if(_EDIT_YN == EDIT_YN.NonEditable)
            {
                string sMSG = string.Empty;
                if (_EDIT_GB == EDIT_GB.Auto)
                {
                    sMSG = "해당 건은 자동전표이므로 수정하실 수 없습니다.";
                }
                else if(_EDIT_GB == EDIT_GB.Approval)
                {
                    sMSG = "해당 건은 승인이 완료되었으므로 수정하실 수 없습니다.";
                }

                LayoutGrpSlip.Text = string.Format(sMSG);
                LayoutGrpSlip.AppearanceGroup.ForeColor = Color.Blue;

                DateEditSlip.ReadOnly = true;
                LkupSlipType.ReadOnly = true;
                BtnAdd.Enabled = false;
                BtnCntsAdd.Enabled = false;
                BtnSave.Enabled = false;
                GridViewSlip.OptionsBehavior.Editable = false;
            }

            if (COPYGB.Equals("1"))
            {
                DateEditSlip.EditValue = SLIPDT;
                LkupSlipType.EditValue = SLIPGB;
                GridSlip.DataSource = DT_COPY;
                LayoutGrpSlip.Text = string.Format("전표복사 -> ( 전표복사의 경우 전표일자를 수정하셔야합니다. 기존 전표일자는 '{0}' 입니다.)", SLIPDT);
                LayoutGrpSlip.AppearanceGroup.ForeColor = Color.Blue;
            }
            else if (COPYGB.Equals("2"))
            {
                DateEditSlip.EditValue = SLIPDT;
                LkupSlipType.EditValue = SLIPGB;
                GridSlip.DataSource = DT_COPY;
                LayoutGrpSlip.Text = string.Format("전표이동 -> ( 전표이동의 경우 전표일자를 수정하셔야합니다. 기존 전표일자는 '{0}' 입니다.)", SLIPDT);
                LayoutGrpSlip.AppearanceGroup.ForeColor = Color.Blue;
            }
            else
            {
                LkupSlipType.EditValue = "3";
                DateEditSlip.EditValue = DateTime.Now.Date;

                //SetLookUpEdit(RepoLkupAcNam, "1", "N");
                //SetLookUpEdit(RepoLkupDealerNm, "2", "Y");
                SettingInitializedControl();
            }
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, PARENT_FORM.Name);
            FmMainToolBar2._FontSetting.SetGridView(GridViewSlip);
            SetPriceDifferency();
        }

        private Dictionary<string, string> GetDeptInfoByUserCd(string sUserCd)
        {
            /*
             * 2020-11-06 부서정보 로직추가
             */

            Dictionary<string, string> dicResult = new Dictionary<string, string>();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("USRCD", sUserCd);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT B.DEPT_CD, C.DEPT_NM ");
            strSql.AppendLine("   FROM ZUSRLST A  ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B  ");
            strSql.AppendLine("     ON A.INSANO = B.EMP_ID ");
            strSql.AppendLine("   LEFT JOIN ACC_DEPT_CD C  ");
            strSql.AppendLine("     ON B.DEPT_CD = C.DEPT_CD ");
            strSql.AppendLine("  WHERE A.USRCD = @USRCD ");

            DataTable dtInfo = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
            if (dtInfo.Rows.Count > 0)
            {
                dicResult.Add("ADPCD", dtInfo.Rows[0]["DEPT_CD"]?.ToString());
                dicResult.Add("ADPNM", dtInfo.Rows[0]["DEPT_NM"]?.ToString());
            }

            dtInfo = null;
            return dicResult;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (GridViewSlip.FocusedRowHandle > 0)
            {
                if (VerificatePreviousRowValue(GridViewSlip.GetDataRow(GridViewSlip.FocusedRowHandle)))
                    return;
            }

            
            if(GridViewSlip.FocusedRowHandle == GridViewSlip.RowCount - 1)
            {
                DataRow row = GridViewSlip.GetDataRow(GridViewSlip.RowCount - 1);

                DataTable dt = (DataTable)GridSlip.DataSource;
                DataRow rowCopy = dt.NewRow();
                rowCopy["LINNO"] = (dt.Rows.Count + 1);
                rowCopy["ATEXT"] = row["ATEXT"];
                rowCopy["CVCOD"] = row["CVCOD"];
                rowCopy["CVNAM"] = row["CVNAM"];

                //rowCopy["CVGB"] = "거래처";

                Dictionary<string,string> dicResult = GetDeptInfoByUserCd(FmMainToolBar2.UserID);
                if(dicResult.Count == 2)
                {
                    rowCopy["ADPCD"] = dicResult["ADPCD"];
                    rowCopy["ADPNM"] = dicResult["ADPNM"];
                }
                
                //dt.ImportRow(rowCopy);
                dt.Rows.Add(rowCopy);
                GridSlip.DataSource = dt;
                //GridViewSlip.AddNewRow();
                //GridViewSlip.UpdateCurrentRow();
                //GridViewSlip.SetRowCellValue(GridViewSlip.RowCount - 1, "LINNO", GridViewSlip.RowCount);
                
                //GridViewSlip.SetRowCellValue(GridViewSlip.RowCount - 1, "ATEXT" , row["ATEXT"]);
                //GridViewSlip.SetRowCellValue(GridViewSlip.RowCount - 1, "CVCOD", row["CVCOD"]);
                //GridViewSlip.SetRowCellValue(GridViewSlip.RowCount - 1, "CVNAM", row["CVNAM"]);

                //GridViewSlip.UpdateCurrentRow();
                GridViewSlip.FocusedRowHandle = GridViewSlip.RowCount - 1;
                GridViewSlip.FocusedColumn = GridColAcCod;
                GridViewSlip.SelectCell(GridViewSlip.RowCount - 1, GridColAcCod);
            }
        }

        private void BtnCntsAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            SaveSlipInfo("CONTINUE");
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            SaveSlipInfo("SAVE");
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Dispose();
        }
        
        //전표구분 값에 따라 컬럼 Visible Setting
        private void LkupSlipType_EditValueChanged(object sender, EventArgs e)
        {
            if (AddModifyGb.Equals("MODIFY"))
                return;

            string sSlipType = LkupSlipType.EditValue?.ToString();
            SetColumnVisibleIndex(sSlipType);
        }

        #region[RepositoryItem_Event(컬럼명 : ACCOD, CVCOD, ACAMT, ADAMT, ATEXT)]

        //RepositoryItem Input 시 해당 셀에 바로 적용
        private void RepoBtnEditAccCd_EditValueChanged(object sender, EventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            GridViewSlip.SetFocusedRowCellValue("ACNAM", btnEdit.EditValue?.ToString());
        }

        public DataRow DrAccInfo;
        private void RepoBtnEditAccCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            string sAccCd = GridViewSlip.GetFocusedRowCellValue("ACCOD")?.ToString();

            AC01001F03 frm = new AC01001F03();
            frm.AccCd = sAccCd;
            frm.PAC02001F02 = this;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (DrAccInfo.ItemArray.Length > 0)
                {
                    GridViewSlip.SetFocusedRowCellValue("ACCOD", DrAccInfo["ACCOD"]);
                    GridViewSlip.SetFocusedRowCellValue("ACNAM", DrAccInfo["ACNAM"]);
                }
            }
        }

        //RepositoryItem Input 시 해당 셀에 바로 적용
        private void RepoBtnEditDealerCd_EditValueChanged(object sender, EventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            GridViewSlip.SetFocusedRowCellValue("CVNAM", btnEdit.EditValue?.ToString());
        }

        public DataRow DrDealerInfo;
        private void RepoBtnEditDealerCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;

            string sDealerCd = GridViewSlip.GetFocusedRowCellValue("CVCOD")?.ToString();

            AC02001F03 frm = new AC02001F03();

            frm.DealerCd = sDealerCd;
            frm.P_AC02001F02 = this;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (DrDealerInfo.ItemArray.Length > 0)
                {
                    GridViewSlip.SetFocusedRowCellValue("CVCOD", DrDealerInfo["DEALER_CD"]);
                    GridViewSlip.SetFocusedRowCellValue("CVNAM", DrDealerInfo["DEALER_NM"]);
                }
            }

            #region 구분선택해서 적용하는 VER
            //string sCvgb = GridViewSlip.GetFocusedRowCellValue("CVGB")?.ToString();

            //if (string.IsNullOrEmpty(sCvgb))
            //{
            //    XtraMessageBox.Show("거래처구분을 먼저 선택해주세요.");
            //    GridViewSlip.FocusedColumn = GridColCvgb;
            //    return;
            //}

            //if (sCvgb.Equals("거래처"))
            //{
            //    string sDealerCd = GridViewSlip.GetFocusedRowCellValue("CVCOD")?.ToString();

            //    AC02001F03 frm = new AC02001F03();

            //    frm.DealerCd = sDealerCd;
            //    frm.P_AC02001F02 = this;
            //    if (frm.ShowDialog() == DialogResult.OK)
            //    {
            //        if (DrDealerInfo.ItemArray.Length > 0)
            //        {
            //            GridViewSlip.SetFocusedRowCellValue("CVCOD", DrDealerInfo["DEALER_CD"]);
            //            GridViewSlip.SetFocusedRowCellValue("CVNAM", DrDealerInfo["DEALER_NM"]);
            //        }
            //    }
            //}
            //else if (sCvgb.Equals("계좌"))
            //{
            //    string sBankCd = btnEdit.EditValue?.ToString();

            //    if (!string.IsNullOrEmpty(sBankCd))
            //    {
            //        DataTable dt = GetBankInfo(sBankCd);
            //        if (dt.Rows.Count == 1)
            //        {
            //            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, dt.Rows[0]["ACNT_CD"]);

            //            string sCvNam = dt.Rows[0]["BANK_NM"]?.ToString();
            //            if (!string.IsNullOrEmpty(dt.Rows[0]["GGBNM"]?.ToString()))
            //            {
            //                sCvNam += "(" + dt.Rows[0]["GGBNM"]?.ToString() + ")";
            //            }
            //            sCvNam += "(" + dt.Rows[0]["BANK_ACNT_NO"]?.ToString() + ")";
            //            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, sCvNam);
            //        }
            //        else
            //        {
            //            PopUpSchAcnt frm = new PopUpSchAcnt();
            //            frm.Owner = this;
            //            frm._FINDWORD = sBankCd;
            //            frm.DataRowSendEvent += new PopUpSchAcnt.SendDataHandler(GetBankInfo);
            //            if (frm.ShowDialog() == DialogResult.OK)
            //            {
            //                GridViewSlip.FocusedColumn = GridColAText;
            //            }
            //        }
            //    }
            //    else
            //    {
            //        PopUpSchAcnt frm = new PopUpSchAcnt();
            //        frm.Owner = this;
            //        frm._FINDWORD = sBankCd;
            //        frm.DataRowSendEvent += new PopUpSchAcnt.SendDataHandler(GetBankInfo);
            //        if (frm.ShowDialog() == DialogResult.OK)
            //        {
            //            GridViewSlip.FocusedColumn = GridColAText;
            //        }
            //    }
            //}
            #endregion
        }

        //RepositoryItem Input 시 해당 셀에 바로 적용
        private void RepoTxtAcAmt_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txtEdit = (TextEdit)sender;
            GridViewSlip.SetFocusedRowCellValue("ACAMT", txtEdit.EditValue?.ToString());
        }

        //RepositoryItem Input 시 해당 셀에 바로 적용
        private void RepoTxtAdAmt_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txtEdit = (TextEdit)sender;
            GridViewSlip.SetFocusedRowCellValue("ADAMT", txtEdit.EditValue?.ToString());
        }

        //RepostoryItme Input 시 해당 셀에 바로 적용
        private void RepoTxtJukyo_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewSlip.SetFocusedRowCellValue("ATEXT", txt.EditValue);
        }

        #endregion[RepositoryItem_Event(컬럼명 : ACCOD, CVCOD, ACAMT, ADAMT, ATEXT)]

        #region[Visible Index Setting]

        //전표구분 Columns VisibleIndex Setting
        private void SetColumnVisibleIndex(string AtGub)
        {
            switch (AtGub)
            {
                case "1":
                    GridColTDate.Visible = false;
                    GridColAtGub.Visible = false;
                    GridColSeqNo.Visible = false;
                    GridColLinNo.Visible = false;
                    GridColAdAmt.Visible = false;

                    GridColAcCod.VisibleIndex = 0;
                    GridColAcNam.VisibleIndex = 1;
                    GridColCvCod.VisibleIndex = 2;
                    GridColCvNam.VisibleIndex = 3;
                    GridColAText.VisibleIndex = 4;
                    GridColAcAmt.VisibleIndex = 5;
                    GridColRK.VisibleIndex = 6;

                    #region 거래처구분 VER
                    //GridColAcCod.VisibleIndex = 0;
                    //GridColAcNam.VisibleIndex = 1;
                    //GridColCvgb.VisibleIndex = 2;
                    //GridColCvCod.VisibleIndex = 3;
                    //GridColCvNam.VisibleIndex = 4;
                    //GridColAText.VisibleIndex = 5;
                    //GridColAcAmt.VisibleIndex = 6;
                    //GridColRK.VisibleIndex = 7;
                    #endregion

                    break;
                case "2":
                    GridColTDate.Visible = false;
                    GridColAtGub.Visible = false;
                    GridColSeqNo.Visible = false;
                    GridColLinNo.Visible = false;
                    GridColAcAmt.Visible = false;

                    GridColAcCod.VisibleIndex = 0;
                    GridColAcNam.VisibleIndex = 1;
                    GridColCvCod.VisibleIndex = 2;
                    GridColCvNam.VisibleIndex = 3;
                    GridColAText.VisibleIndex = 4;
                    GridColAdAmt.VisibleIndex = 5;
                    GridColRK.VisibleIndex = 6;

                    #region 거래처구분 버전
                    //GridColAcCod.VisibleIndex = 0;
                    //GridColAcNam.VisibleIndex = 1;
                    //GridColCvgb.VisibleIndex = 2;
                    //GridColCvCod.VisibleIndex = 3;
                    //GridColCvNam.VisibleIndex = 4;
                    //GridColAText.VisibleIndex = 5;
                    //GridColAdAmt.VisibleIndex = 6;
                    //GridColRK.VisibleIndex = 7;
                    #endregion

                    break;
                default:
                    GridColTDate.Visible = false;
                    GridColAtGub.Visible = false;
                    GridColSeqNo.Visible = false;
                    GridColLinNo.Visible = false;

                    GridColAcCod.VisibleIndex = 0;
                    GridColAcNam.VisibleIndex = 1;
                    GridColCvCod.VisibleIndex = 2;
                    GridColCvNam.VisibleIndex = 3;
                    GridColAText.VisibleIndex = 4;
                    GridColAcAmt.VisibleIndex = 5;
                    GridColAdAmt.VisibleIndex = 6;
                    GridColRK.VisibleIndex = 7;

                    #region 거래처구분 버전
                    //GridColAcCod.VisibleIndex = 0;
                    //GridColAcNam.VisibleIndex = 1;
                    //GridColCvgb.VisibleIndex = 2;
                    //GridColCvCod.VisibleIndex = 3;
                    //GridColCvNam.VisibleIndex = 4;
                    //GridColAText.VisibleIndex = 5;
                    //GridColAcAmt.VisibleIndex = 6;
                    //GridColAdAmt.VisibleIndex = 7;
                    //GridColRK.VisibleIndex = 8;
                    #endregion
                    break;
            }
        }

        #endregion[Visible Index Setting]

        #region[KeyDown Event]

        private void AC02001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                if (GridViewSlip.RowCount < 1)
                    return;

                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F2)
            {
                BtnCntsAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }


        private void GridViewSlip_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewSlip.RowCount < 1)
                return;

            if (e.KeyCode == Keys.Down)
            {
                BtnAdd_Click(null, null);
            }
            //else if (e.KeyCode == Keys.Up)
            //{
            //    if (CancelCurrentRowValueCheck(GridViewSlip.GetDataRow(GridViewSlip.FocusedRowHandle)))
            //        return;

            //    //GridViewSlip.RowStyle -= GridViewSlip_RowStyle;
            //    //GridViewSlip.CustomDrawRowIndicator -= GridViewSlip_CustomDrawRowIndicator;

            //    //GridViewSlip.DeleteRow(GridViewSlip.RowCount);
            //    //GridViewSlip.DeleteSelectedRows();
            //    //DataTable dt = (DataTable)GridSlip.DataSource;
            //    //dt.Rows.RemoveAt(dt.Rows.Count - 1);
            //    //GridSlip.DataSource = null;
            //    //GridSlip.DataSource = dt;
            //    //GridViewSlip.FocusedRowHandle = GridViewSlip.RowCount;
            //    //GridViewSlip.FocusedRowHandle = (dt.Rows.Count - 1);
            //    //GridViewSlip.UpdateCurrentRow();
            //    //DelaySystem(3000);
            //    //GridViewSlip.FocusedRowHandle = (dt.Rows.Count - 1);
            //    //GridViewSlip.SelectRow(GridViewSlip.FocusedRowHandle);
            //    //GridViewSlip.SelectCell((dt.Rows.Count - 1), GridColAcCod);
            //}
        }

        private void GridViewSlip_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up)
            {
                if (CancelCurrentRowValueCheck(GridViewSlip.GetDataRow(GridViewSlip.RowCount - 1)))
                    return;

                //GridViewSlip.RowStyle -= GridViewSlip_RowStyle;
                //GridViewSlip.CustomDrawRowIndicator -= GridViewSlip_CustomDrawRowIndicator;

                GridViewSlip.DeleteRow(GridViewSlip.RowCount - 1);
                //GridViewSlip.DeleteSelectedRows();
                //DataTable dt = (DataTable)GridSlip.DataSource;
                //dt.Rows.RemoveAt(dt.Rows.Count - 1);
                //GridSlip.DataSource = null;
                //GridSlip.DataSource = dt;
                //GridViewSlip.FocusedRowHandle = GridViewSlip.RowCount;
                //GridViewSlip.FocusedRowHandle = (dt.Rows.Count - 1);
                //GridViewSlip.UpdateCurrentRow();
                //DelaySystem(3000);
                //GridViewSlip.FocusedRowHandle = (dt.Rows.Count - 1);
                //GridViewSlip.SelectRow(GridViewSlip.FocusedRowHandle);
                //GridViewSlip.SelectCell((dt.Rows.Count - 1), GridColAcCod);
            }
        }

        //지연 함수... 
        private void DelaySystem(int MS)
        { /* 함수명 : DelaySystem * 1000ms = 1초 * 전달인자 : 얼마나 지연시킬것인가에 대한 변수 * */
            DateTime dtAfter = DateTime.Now;
            TimeSpan dtDuration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime dtThis = dtAfter.Add(dtDuration); while (dtThis >= dtAfter)
            {
                //DoEvent () 를 사용 해서 지연 시간 동안 
                //버튼 클릭 이벤트 및 다른 윈도우 이벤트를 받을 수 있게끔 하는 역할 
                //없으면 지연 동안 다른 이벤트를 받지 못함... 
                System.Windows.Forms.Application.DoEvents();
                //현재 시간 얻어 오기...
                dtAfter = DateTime.Now;
            }
        }

        //그리드 상 계정코드 입력 후 엔터키 시 계정명 Select
        private void RepoBtnEditAccCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewSlip.SelectedRowsCount < 1)
                return;

            if(e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sAcCod = btnEdit.EditValue?.ToString();
                if (!string.IsNullOrEmpty(sAcCod))
                {
                    DataTable dt = GetAccountInfo(sAcCod);
                    if(dt.Rows.Count == 1)
                    {
                        GridViewSlip.SetFocusedRowCellValue("ACCOD", dt.Rows[0]["ACCOD"]);
                        GridViewSlip.SetFocusedRowCellValue("ACNAM", dt.Rows[0]["ACNAM"]);
                    }
                    else
                    {
                        AC01001F03 frm = new AC01001F03();
                        frm.PAC02001F02 = this;
                        frm.AccCd = sAcCod;
                        if(frm.ShowDialog() == DialogResult.OK)
                        {
                            GridViewSlip.SetFocusedRowCellValue("ACCOD", DrAccInfo["ACCOD"]);
                            GridViewSlip.SetFocusedRowCellValue("ACNAM", DrAccInfo["ACNAM"]);
                        }
                    }
                }
                else
                {
                    GridViewSlip.SetFocusedRowCellValue("ACCOD", null);
                    GridViewSlip.SetFocusedRowCellValue("ACNAM", null);
                }
            }
        }

        //그리드 상 거래처코드 입력 후 엔터키 시 거래처명 Select
        private void RepoBtnEditDealerCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewSlip.SelectedRowsCount < 1)
                return;

            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;

                string sDealerCd = btnEdit.EditValue?.ToString();

                if (!string.IsNullOrEmpty(sDealerCd))
                {
                    DataTable dt = GetDealerInfo(sDealerCd);
                    if (dt.Rows.Count == 1)
                    {
                        GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, dt.Rows[0]["DEALER_CD"]);
                        GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, dt.Rows[0]["DEALER_NM"]);
                    }
                    else
                    {
                        AC02001F03 frm = new AC02001F03();
                        frm.P_AC02001F02 = this;
                        frm.DealerCd = sDealerCd;
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, DrDealerInfo["DEALER_CD"]);
                            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, DrDealerInfo["DEALER_NM"]);
                        }
                    }
                }
                else
                {
                    GridViewSlip.SetFocusedRowCellValue("CVCOD", null);
                    GridViewSlip.SetFocusedRowCellValue("CVNAM", null);
                }

                #region 거래처구분(계좌,거래처) 선택하는 VER
                //string sCvgb = GridViewSlip.GetFocusedRowCellValue("CVGB")?.ToString();

                //if (string.IsNullOrEmpty(sCvgb))
                //{
                //    XtraMessageBox.Show("거래처구분을 먼저 선택해주세요.");
                //    GridViewSlip.FocusedColumn = GridColCvgb;
                //    return;
                //}

                //if (sCvgb.Equals("거래처"))
                //{
                //    string sDealerCd = btnEdit.EditValue?.ToString();

                //    if (!string.IsNullOrEmpty(sDealerCd))
                //    {
                //        DataTable dt = GetDealerInfo(sDealerCd);
                //        if (dt.Rows.Count == 1)
                //        {
                //            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, dt.Rows[0]["DEALER_CD"]);
                //            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, dt.Rows[0]["DEALER_NM"]);
                //        }
                //        else
                //        {
                //            AC02001F03 frm = new AC02001F03();
                //            frm.P_AC02001F02 = this;
                //            frm.DealerCd = sDealerCd;
                //            if (frm.ShowDialog() == DialogResult.OK)
                //            {
                //                GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, DrDealerInfo["DEALER_CD"]);
                //                GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, DrDealerInfo["DEALER_NM"]);
                //            }
                //        }
                //    }
                //    else
                //    {
                //        GridViewSlip.SetFocusedRowCellValue("CVCOD", null);
                //        GridViewSlip.SetFocusedRowCellValue("CVNAM", null);
                //    }
                //}
                //else if (sCvgb.Equals("계좌"))
                //{
                //    string sBankCd = btnEdit.EditValue?.ToString();

                //    if (!string.IsNullOrEmpty(sBankCd))
                //    {
                //        DataTable dt = GetBankInfo(sBankCd);
                //        if (dt.Rows.Count == 1)
                //        {
                //            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, dt.Rows[0]["ACNT_CD"]);

                //            string sCvNam = dt.Rows[0]["BANK_NM"]?.ToString();
                //            if (!string.IsNullOrEmpty(dt.Rows[0]["GGBNM"]?.ToString()))
                //            {
                //                sCvNam += "(" + dt.Rows[0]["GGBNM"]?.ToString() + ")";
                //            }
                //            sCvNam += "(" + dt.Rows[0]["BANK_ACNT_NO"]?.ToString() + ")";
                //            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, sCvNam);
                //        }
                //        else
                //        {
                //            PopUpSchAcnt frm = new PopUpSchAcnt();
                //            frm.Owner = this;
                //            frm._FINDWORD = sBankCd;
                //            frm.DataRowSendEvent += new PopUpSchAcnt.SendDataHandler(GetBankInfo);
                //            if (frm.ShowDialog() == DialogResult.OK)
                //            {

                //            }
                //        }
                //    }
                //    else
                //    {
                //        GridViewSlip.SetFocusedRowCellValue("CVCOD", null);
                //        GridViewSlip.SetFocusedRowCellValue("CVNAM", null);
                //    }
                //}
                #endregion
            }
        }

        #region 거래처구분 선택하는 VER
        //private void GetBankInfo(DataRow row)
        //{
        //    GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvCod, row["ACNT_CD"]);

        //    string sCvNam = row["BANK_NM"]?.ToString();
        //    if (!string.IsNullOrEmpty(row["GGBNM"]?.ToString()))
        //    {
        //        sCvNam += "(" + row["GGBNM"]?.ToString() + ")";
        //    }
        //    sCvNam += "(" + row["BANK_ACNT_NO"]?.ToString() + ")";
        //    GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColCvNam, sCvNam);
        //}
        #endregion

        private void RepoTxtRemark_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if(GridViewSlip.FocusedRowHandle == GridViewSlip.RowCount - 1)
                {
                    BtnAdd_Click(null, null);
                }
            }
        }

        public DataRow DR_DEPT_INFO;
        private void RepoBtnEditAdpCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            AC02001F06 frm = new AC02001F06();
            frm.PARENT_FORM = this;
            frm.FIND_WORD = btnEdit.EditValue?.ToString();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                GridViewSlip.SetFocusedRowCellValue("ADPCD", DR_DEPT_INFO["DEPT_CD"]);
                GridViewSlip.SetFocusedRowCellValue("ADPNM", DR_DEPT_INFO["DEPT_NM"]);
            }
        }

        private void RepoBtnEditAdpCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewSlip.SelectedRowsCount < 1)
                return;

            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sAdpCd = btnEdit.EditValue?.ToString();
                if (!string.IsNullOrEmpty(sAdpCd))
                {
                    DataTable dt = GetDeptInfo(sAdpCd);
                    if (dt.Rows.Count == 1)
                    {
                        GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColAdpCd, dt.Rows[0]["DEPT_CD"]);
                        GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColAdpNm, dt.Rows[0]["DEPT_NM"]);
                    }
                    else
                    {
                        AC02001F06 frm = new AC02001F06();
                        frm.PARENT_FORM = this;
                        frm.FIND_WORD = sAdpCd;
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColAdpCd, DR_DEPT_INFO["DEPT_CD"]);
                            GridViewSlip.SetRowCellValue(GridViewSlip.FocusedRowHandle, GridColAdpNm, DR_DEPT_INFO["DEPT_NM"]);
                        }
                    }
                }
                else
                {
                    GridViewSlip.SetFocusedRowCellValue("DEPT_CD", null);
                    GridViewSlip.SetFocusedRowCellValue("DEPT_NM", null);
                }
            }
        }

        #endregion[KeyDown Event]

        #region[기타 함수처리]

        private DataTable SetInitDataTable(string[] sColumnNames)
        {
            DataTable dt = new DataTable();

            if (PARENT_FORM.Name.Equals("AC13001F01"))
            {
                string sMakeNo = ExternalParam["MAKENO"];
                string sMakeNoLn = ExternalParam["MAKENO_LN"];
                string sLnEseq = ExternalParam["LN_ESEQ"];
                string sMgCost = ExternalParam["MG_COST"];
                DateEditSlip.EditValue = ExternalParam["OCCUR_DT"];
                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT '' AS TDATE ");
                strSql.AppendLine("      , '' AS ATGUB ");
                strSql.AppendLine("      , '' AS SEQNO ");
                strSql.AppendLine("      , 1 AS LINNO ");
                strSql.AppendLine("      , C.ACCOD AS ACCOD ");
                strSql.AppendLine("      , C.ACNAM AS ACNAM ");
                strSql.AppendLine("      , D.DEALER_CD AS CVCOD ");
                strSql.AppendLine("      , D.DEALER_NM AS CVNAM ");
                strSql.AppendLine("      , CONCAT('고장내역: ',B.ECONTENT) AS ATEXT ");
                strSql.AppendLine("      , " + sMgCost + " AS ACAMT ");
                strSql.AppendLine("      , " + sMgCost + " AS ADAMT ");
                strSql.AppendLine("      , '' AS ADPCD ");
                strSql.AppendLine("      , '' AS ADPNM ");
                strSql.AppendLine("      , CONCAT('수리사항: ',B.EREPAIR) AS RK ");
                strSql.AppendLine("   FROM EQUIP_CD_HISTORY A         ");
                strSql.AppendLine("   LEFT OUTER JOIN MAKE_EXPENSE B  ");
                strSql.AppendLine("     ON A.MAKENO = B.MAKENO        ");
                strSql.AppendLine("    AND A.MAKENO_LN = B.MAKENO_LN  ");
                strSql.AppendLine("    AND A.LN_ESEQ = B.LN_ESEQ      ");
                strSql.AppendLine("   LEFT OUTER JOIN ACMSTF C   ");
                strSql.AppendLine("     ON ACCOD = '0530' /*소모품비(제조-생산)*/  ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D  ");
                strSql.AppendLine("     ON B.ECVNAM = D.DEALER_NM      ");
                strSql.AppendLine("  WHERE A.MAKENO = " + sMakeNo);
                strSql.AppendLine("    AND A.MAKENO_LN = " + sMakeNoLn);
                strSql.AppendLine("    AND A.LN_ESEQ = " + sLnEseq);

                //strSql.AppendLine(" SELECT D.DEALER_CD AS CVNAM        ");
                //strSql.AppendLine(" 	 , CONCAT('고장내역: ',B.ECONTENT) AS ATEXT        ");
                //strSql.AppendLine(" 	 , CONCAT('수리사항: ',B.EREPAIR) AS RK        ");
                //strSql.AppendLine(" 	 , A.MAKENO                  ");
                //strSql.AppendLine(" 	 , A.MAKENO_LN               ");
                //strSql.AppendLine(" 	 , C.ACCOD AS ACNAM          ");
                //strSql.AppendLine("      , C.ACCOD");
                //strSql.AppendLine("      , D.DEALER_CD AS CVCOD");
                //strSql.AppendLine("      , " + sMgCost + " AS ACAMT");
                //strSql.AppendLine("      , " + sMgCost + " AS ADAMT");
                //strSql.AppendLine("      , 1 AS LINNO");
                //strSql.AppendLine("   FROM EQUIP_CD_HISTORY A        ");
                //strSql.AppendLine("   LEFT OUTER JOIN MAKE_EXPENSE B ");
                //strSql.AppendLine("     ON A.MAKENO = B.MAKENO       ");
                //strSql.AppendLine("    AND A.MAKENO_LN = B.MAKENO_LN ");
                //strSql.AppendLine("    AND A.LN_ESEQ = B.LN_ESEQ     ");
                //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF C  ");
                //strSql.AppendLine("     ON ACCOD = '0122' /*소모품비(제조-생산)*/ ");
                //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD D ");
                //strSql.AppendLine("     ON B.ECVNAM = D.DEALER_NM     ");
                //strSql.AppendLine("  WHERE A.MAKENO = " + sMakeNo);
                //strSql.AppendLine("    AND A.MAKENO_LN = " + sMakeNoLn);
                //strSql.AppendLine("    AND A.LN_ESEQ = " + sLnEseq);

                dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            }
            else
            {
                dt.TableName = "Table";
                for (int i = 0; i < sColumnNames.Length; i++)
                {
                    dt.Columns.Add(sColumnNames[i]);
                }

                DataRow row = dt.NewRow();
                Dictionary<string, string> dicResult = GetDeptInfoByUserCd(FmMainToolBar2.UserID);
                if (dicResult.Count == 2)
                {
                    row["ADPCD"] = dicResult["ADPCD"];
                    row["ADPNM"] = dicResult["ADPNM"];
                }

                dt.Rows.Add(row);
            }
            
            dt.Rows[0]["LINNO"] = 1;

            return dt;
        }

        private string[] SetSlipDataList()
        {
            DataTable dt = GetAccDetailInfo("99999999", "T", "999");

            string[] sArrColumns = new string[dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sArrColumns[i] = dt.Columns[i].ColumnName;
            }

            return sArrColumns;
        }

        private void DisableControls(string sGubun)
        {
            if (sGubun.Equals("T"))
            {
                DateEditSlip.Enabled = true;
                LkupSlipType.Enabled = true;
            }
            else
            {
                DateEditSlip.Enabled = false;
                LkupSlipType.Enabled = false;
            }
        }

        private object[] VerifyCvCodAndCvNam(GridView GridViewSlip)
        {
            object[] objArrValue = new object[5];
            objArrValue[0] = "Y";
            objArrValue[1] = 0;
            objArrValue[2] = "";
            objArrValue[3] = "";
            objArrValue[4] = new GridColumn();

            for(int i = 0; i < GridViewSlip.RowCount; i++)
            {
                string sCvCod = GridViewSlip.GetDataRow(i)["CVCOD"]?.ToString();
                string sCvNam = GridViewSlip.GetDataRow(i)["CVNAM"]?.ToString();
                if(string.IsNullOrEmpty(sCvCod) && string.IsNullOrEmpty(sCvNam))
                {
                    objArrValue[0] = "N";
                    objArrValue[1] = i;
                    objArrValue[2] = "거래처코드나 거래처명";
                    objArrValue[3] = "CVCOD";
                    objArrValue[4] = GridColCvCod;

                    break;
                }
            }

            return objArrValue;
        }

        private object[] VerifyDataRow(GridView GridViewSlip, string sAtGub)
        {
            /*
                sArrValue Index [0, 0] = 검증완료여부
                          Index [0, 1] = RowIndex(RowHandle)
                          Index [0, 2] = 에러발생 컬럼 Caption
                          Index [0, 3] = 에러발생 컬럼명
                          Index [0, 4] = 해당 셀에 대한 GridColumns 객체 Name
             */
            object[] objArrValue = new object[5];
            objArrValue[0] = "Y";
            objArrValue[1] = 0;
            objArrValue[2] = "";
            objArrValue[3] = "";
            objArrValue[4] = new GridColumn();

            for (int i = 0; i < GridViewSlip.RowCount; i++)
            {
                string sAccCd = GridViewSlip.GetDataRow(i)["ACCOD"]?.ToString();
                if (string.IsNullOrEmpty(sAccCd))
                {
                    objArrValue[0] = "N";
                    objArrValue[1] = i;
                    objArrValue[2] = "계정코드";
                    objArrValue[3] = "ACCOD";
                    objArrValue[4] = GridColAcCod;
                    break;
                }

                //string sCvCod = GridViewSlip.GetDataRow(i)["CVCOD"]?.ToString();
                //if (string.IsNullOrEmpty(sCvCod))
                //{
                //    objArrValue[0] = "N";
                //    objArrValue[1] = i;
                //    objArrValue[2] = "거래처코드";
                //    objArrValue[3] = "CVCOD";
                //    objArrValue[4] = GridColCvCod;
                //    break;
                //}

                string sAcAmt = GridViewSlip.GetDataRow(i)["ACAMT"]?.ToString();
                double dAcAmt = string.IsNullOrEmpty(sAcAmt) ? 0 : Convert.ToDouble(sAcAmt);
                string sAdAmt = GridViewSlip.GetDataRow(i)["ADAMT"]?.ToString();
                double dAdAmt = string.IsNullOrEmpty(sAdAmt) ? 0 : Convert.ToDouble(sAdAmt);
                if (sAtGub.Equals("3") || sAtGub.Equals("4"))
                {
                    if ((string.IsNullOrEmpty(sAcAmt) && string.IsNullOrEmpty(sAdAmt)) || 
                        (dAcAmt == 0 && dAdAmt == 0))
                    {
                        objArrValue[0] = "N";
                        objArrValue[1] = i;
                        objArrValue[2] = "금액";
                        objArrValue[3] = "ACAMT";
                        objArrValue[4] = GridColAcAmt;
                        break;
                    }
                }
                else if (sAtGub.Equals("1")) //출금
                {
                    if (string.IsNullOrEmpty(sAccCd))
                    {
                        objArrValue[0] = "N";
                        objArrValue[1] = i;
                        objArrValue[2] = "차변금액";
                        objArrValue[3] = "ACAMT";
                        objArrValue[4] = GridColAcAmt;
                        break;
                    }
                }
                else if (sAtGub.Equals("2")) //입금
                {
                    if (string.IsNullOrEmpty(sAdAmt))
                    {
                        objArrValue[0] = "N";
                        objArrValue[1] = i;
                        objArrValue[2] = "대변금액";
                        objArrValue[3] = "ADAMT";
                        objArrValue[4] = GridColAdAmt;
                        break;
                    }
                }

                
            }
            

            return objArrValue;
        }

        //전표번호 채번
        private string GetSlipNo(SqlCommand cmd, string sSlipYmd, string sAtGub)
        {
            string sSlipNo_FirstToken = string.Empty;
            string[] sAtGub_Range = new string[2];

            if (sAtGub.Equals("1") || sAtGub.Equals("2")) //입금, 출금시 전표번호 첫시작은 0으로 시작
            {
                sSlipNo_FirstToken = "0";
                sAtGub_Range[0] = "1";
                sAtGub_Range[1] = "2";
            }
            else if (sAtGub.Equals("3") || sAtGub.Equals("4")) //대체, 결산시 전표번호 첫시작은 5으로 시작
            {
                sSlipNo_FirstToken = "5";
                sAtGub_Range[0] = "3";
                sAtGub_Range[1] = "4";
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE WHEN LEN(CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR)) = 1 THEN CONCAT('" + sSlipNo_FirstToken + "', '00', CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR)) ");
            strSql.AppendLine(" 			WHEN LEN(CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR)) = 2 THEN CONCAT('" + sSlipNo_FirstToken + "', '0', CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR)) ");
            strSql.AppendLine(" 			WHEN LEN(CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR)) = 3 THEN CONCAT('" + sSlipNo_FirstToken + "', '', CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR)) ");
            strSql.AppendLine(" 			ELSE CAST(ISNULL(MAX(A.SEQNO), 0) + 1 AS VARCHAR) ");
            strSql.AppendLine(" 			 END AS RESULT_MAX_VALUE ");
            strSql.AppendLine("   FROM ACTRAN A  ");
            strSql.AppendLine("  WHERE A.TDATE = '" + sSlipYmd + "' ");
            strSql.AppendLine("    AND A.ATGUB IN ('" + sAtGub_Range[0] + "', '" + sAtGub_Range[1] + "') ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            return cmd.ExecuteScalar().ToString();
        }

        //차대구분 합 체크
        private bool VerificateChaDaeSummary(GridView grdView)
        {
            double dCha = 0;
            double dDae = 0;
            for(int i = 0; i < grdView.RowCount; i++)
            {
                string sCha = string.IsNullOrEmpty(grdView.GetDataRow(i)["ACAMT"]?.ToString()) ? "0" : grdView.GetDataRow(i)["ACAMT"].ToString();
                string sDae = string.IsNullOrEmpty(grdView.GetDataRow(i)["ADAMT"]?.ToString()) ? "0" : grdView.GetDataRow(i)["ADAMT"].ToString();
                dCha += Convert.ToDouble(sCha);
                dDae += Convert.ToDouble(sDae);
            }
            if (dCha != dDae)
                return true;
            else
                return false;
        }

        //Row 추가 시 이전 Row Value Check
        private bool VerificatePreviousRowValue(DataRow row)
        {
            if (ComnGridFunc.VerificateCheckLastRowFocusing(GridViewSlip))
                return true;

            bool bResult = false;

            string[] sArr = new string[1];

            sArr[0] = row["ACCOD"]?.ToString();

            for(int i = 0; i < sArr.Length; i++)
            {
                if (string.IsNullOrEmpty(sArr[i]))
                    bResult = true;
            }

            return bResult;
        }
        
        private bool CancelCurrentRowValueCheck(DataRow row)
        {
            if (GridViewSlip.RowCount < 2)
                return true;

            //행추가된 행이더라도 계정코드가 입력되면 취소불가
            string sAcCod = row["ACCOD"]?.ToString();
            if (!string.IsNullOrEmpty(sAcCod))
                return true;

            string sSeqNo = row["SEQNO"]?.ToString();
            if (string.IsNullOrEmpty(sSeqNo))
                return false;
            else
                return true;

            ///*
            //    GridView의 마지막 로우에만 해당 로직 적용
            // */
            //if (GridViewSlip.FocusedRowHandle == (GridViewSlip.RowCount - 1))
            //{
            //    string sSeqNo = row["SEQNO"]?.ToString();
            //    if (string.IsNullOrEmpty(sSeqNo))
            //        return false;
            //    else
            //        return true;

            //}
            //else
            //{
            //    return true;
            //}

            /*
                DB에 존재하는 데이터는 ACTRAN 테이블의 PK값이 존재하여
                전표번호가 존재할 시 ROW 취소는 할 수 없다.
             */
            //string sSeqNo = row["SEQNO"]?.ToString();

            //if (string.IsNullOrEmpty(sSeqNo))
            //    return false;
            //else
            //    return true;
        }

        private void SettingInitializedControl()
        {
            if (AddModifyGb.Equals("ADD"))
            {
                GridSlip.DataSource = SetInitDataTable(SetSlipDataList());

                DisableControls("T");
                GridViewSlip.FocusedRowHandle = 0;
                GridViewSlip.FocusedColumn = GridColAcNam;
                GridViewSlip.SelectCell(GridViewSlip.FocusedRowHandle, GridColAcNam);
                //GridViewSlip.SetFocusedRowCellValue(GridColCvgb, "거래처");
            }
            else if (AddModifyGb.Equals("MODIFY"))
            {
                if (DrSlipInfo == null)
                {
                    XtraMessageBox.Show("선택된 전표에 대한 정보가 존재하지 않습니다.");
                    Dispose();
                }
                DisableControls("F");
                GridSlip.DataSource = GetAccDetailInfo(DrSlipInfo["TDATE"].ToString(), DrSlipInfo["ATGUB"].ToString(), DrSlipInfo["SEQNO"].ToString());
                DateEditSlip.EditValue = DrSlipInfo["TDATE"];
                LkupSlipType.EditValue = DrSlipInfo["ATGUB"];
                GridViewSlip.FocusedRowHandle = 0;
                GridViewSlip.FocusedColumn = GridColAcNam;
                GridViewSlip.SelectCell(GridViewSlip.FocusedRowHandle, GridColAcNam);

                //GridViewSlip.SelectCell(GridViewSlip.FocusedRowHandle, GridColAcCod);
            }

            string sSlipType = LkupSlipType.EditValue?.ToString();
            SetColumnVisibleIndex(sSlipType);
        }

        #endregion[기타 함수처리]

        #region[Execute By Query]

        private DataTable GetDealerInfo(string sDealerCd)
        {
            StringBuilder strSql = new StringBuilder();

            /*
             * 수정일자 : 2021-02-07 (현업요청)
             * 수정자 : 고혜성
             * 수정내용 : 거래처초성검색 추가
             */

            strSql.Clear();
            strSql.AppendLine(" SELECT A.DEALER_CD  ");
            strSql.AppendLine(" 	 , A.DEALER_NM AS DEALER_NM ");
            strSql.AppendLine(" 	 , A.IDT_NO  ");
            strSql.AppendLine(" 	 , A.DEALER_GB  ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A  ");
            strSql.AppendLine("  WHERE (A.DEALER_CD LIKE '" + sDealerCd + "' ");
            strSql.AppendLine("     OR (A.DEALER_NM LIKE '%" + sDealerCd + "%' OR A.INITIAL_NM LIKE '%" + sDealerCd + "%')) ");
            strSql.AppendLine("    AND A.EOB_YN = 'N' ");
            //은행제외
            strSql.AppendLine("    AND A.DEALER_GB NOT IN('계좌', '금융')    ");
            strSql.AppendLine("    AND A.DEALER_CD NOT BETWEEN 8999 AND 10000");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        private DataTable GetDeptInfo(string sAdpCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("   SELECT DEPT_CD ");
            strSql.AppendLine("        , DEPT_NM ");
            strSql.AppendLine("        , ( SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = A.UP_DEPT_CD ) AS UP_DEPT_NM ");
            strSql.AppendLine("     FROM ACC_DEPT_CD A  ");
            strSql.AppendLine("    WHERE A.USE_YN = 'Y' ");
            strSql.AppendLine("      AND (DEPT_CD = '" + sAdpCd + "' " );
            strSql.AppendLine("           OR ");
            strSql.AppendLine("           DEPT_NM LIKE '%" + sAdpCd + "%' ) ");
           

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        private DataTable GetAccountInfo(string sAccCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.ACCOD ");
            strSql.AppendLine("      , A.ACNAM ");
            strSql.AppendLine("   FROM ACMSTF A ");
            strSql.AppendLine("  WHERE A.ACCOD LIKE '" + sAccCd + "' ");
            strSql.AppendLine("     OR A.ACNAM LIKE '%" + sAccCd + "%' ");
            
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private DataTable GetBankInfo(string sBankNam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("SELECT A1.ACNT_CD             ");
            strSql.AppendLine("     , A1.BANK_ACNT_NO        ");
            strSql.AppendLine("     , A1.BANK_CD             ");
            strSql.AppendLine("     , B1.COM_NM AS BANK_NM   ");
            strSql.AppendLine("     , A1.PSSEQ               ");
            strSql.AppendLine("     , A1.GEJAGB              ");
            strSql.AppendLine("     , B2.COM_NM AS GGBNM    ");
            strSql.AppendLine("     , A1.ACC_CD              ");
            strSql.AppendLine("     , A1.RPLC_ACC_CD         ");
            strSql.AppendLine("     , A1.BRANCH_NM           ");
            strSql.AppendLine("     , A1.FINANCE_GOODS_NM    ");
            strSql.AppendLine("     , A1.CNTR_YMD            ");
            strSql.AppendLine("     , A1.CNTR_AMT            ");
            strSql.AppendLine("     , A1.INRST_RATE          ");
            strSql.AppendLine("     , A1.FRST_PAYIN_YMD      ");
            strSql.AppendLine("     , A1.PAYIN_AMT           ");
            strSql.AppendLine("     , A1.PAYIN_DD            ");
            strSql.AppendLine("     , A1.PAYIN_METHOD        ");
            strSql.AppendLine("     , A1.EXPIRE_YMD          ");
            strSql.AppendLine("     , A1.EXPIRE_INRST_AMT    ");
            strSql.AppendLine("     , A1.TRMN_YN             ");
            strSql.AppendLine("     , A1.TRMN_YMD            ");
            strSql.AppendLine("     , A1.LIQ_CPTL_YN         ");
            strSql.AppendLine("     , A1.RPLC_PMNT_YN        ");
            strSql.AppendLine("     , A1.RECEIVE_ACNT_YN     ");
            strSql.AppendLine("  FROM ACC_ACNT_CD A1         ");
            strSql.AppendLine("  LEFT JOIN COM_BASE_CD B1    ");
            strSql.AppendLine("    ON A1.BANK_CD = B1.COM_CD ");
            strSql.AppendLine("   AND B1.CD_GB = 'BANK_CD'   ");
            strSql.AppendLine("  LEFT JOIN COM_BASE_CD B2    ");
            strSql.AppendLine("    ON A1.GEJAGB = B2.COM_CD  ");
            strSql.AppendLine("   AND B2.CD_GB = 'GEJAGB'    ");
            strSql.AppendLine("  WHERE A1.BANK_ACNT_NO LIKE '%" + sBankNam + "%' ");
            strSql.AppendLine("     OR B1.COM_NM LIKE '%" + sBankNam + "%' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private DataTable GetAccDetailInfo(string sTDate, string sAtGub, string sSeqno)
        {
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
            strSql.AppendLine("  WHERE A.TDATE = '" + sTDate.Replace("-", "").Substring(0, 8) + "' ");
            strSql.AppendLine("    AND A.ATGUB = '" + sAtGub + "' ");
            strSql.AppendLine("    AND A.SEQNO = '" + sSeqno + "' ");
            /*
             * 2021-01-18
             * 입출금전표의 경우 현금은 제외하고 조회
             * 
             * 2021-12-16
             * 주석처리 - 현금전표 수정할때 현금작성 내용이 안떠서
             */
            //if(sAtGub.Equals("1") || sAtGub.Equals("2"))
            //{
            //    strSql.AppendLine("  AND A.ACCOD <> '" + ComnEtcFunc.CashCode + "'");
            //}
            

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        
        private void SaveSlipInfo(string sSaveGb)
        {
            string sChkSlipYmd = DateEditSlip.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sChkSlipGb = LkupSlipType.EditValue?.ToString();

            /*
             * 전표복사일 경우 전역변수 COPY_DT와 입력받은 전표일자를 비교하여 
             * 같을경우 수정할 수 없도록 구현
             */
            if (COPYGB.Equals("1"))
            {
                if (SLIPDT.Replace("-", "").Equals(sChkSlipYmd))
                {
                    XtraMessageBox.Show("전표복사일 경우 전표일자를 수정하여야합니다.\r\n기존데이터를 수정하려면 해당 데이터를 클릭하세요.");
                    return;
                }
            }
            else if (COPYGB.Equals("2"))
            {
                if (SLIPDT.Replace("-", "").Equals(sChkSlipYmd))
                {
                    XtraMessageBox.Show("전표이동일 경우 전표일자를 수정하여야합니다.\r\n기존데이터를 수정하려면 해당 데이터를 클릭하세요.");
                    return;
                }
            }

            if (string.IsNullOrEmpty(sChkSlipYmd))
            {
                XtraMessageBox.Show("전표일자를 선택하세요.");
                DateEditSlip.Focus();
                DateEditSlip.SelectAll();
                return;
            }
            if (string.IsNullOrEmpty(sChkSlipGb))
            {
                XtraMessageBox.Show("전표구분을 선택하세요.");
                LkupSlipType.Focus();
                LkupSlipType.SelectAll();
                return;
            }

            if(GridViewSlip.RowCount == 0)
            {
                XtraMessageBox.Show("전표항목을 추가하세요.");
                return;
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            try
            {
                Cursor = Cursors.WaitCursor;
                
                object[] objArrVerification = VerifyDataRow(GridViewSlip, sChkSlipGb); //인덱스의 자세한 설명은 해당 함수에 주석처리로 설명 
                if (objArrVerification[0].ToString().Equals("N"))
                {
                    XtraMessageBox.Show(objArrVerification[2].ToString() + "을(를) 입력하세요.");
                    GridViewSlip.FocusedRowHandle = Convert.ToInt32(objArrVerification[1]);
                    GridViewSlip.FocusedColumn = (GridColumn)objArrVerification[4];
                    GridViewSlip.SelectCell(Convert.ToInt32(objArrVerification[1]), (GridColumn)objArrVerification[4]);
                    return;
                }

                if(sChkSlipGb.Equals("3") || sChkSlipGb.Equals("4")) //대체, 결산 시 차대변 합 체크
                {
                    if (VerificateChaDaeSummary(GridViewSlip))
                    {
                        Cursor = Cursors.Default;
                        XtraMessageBox.Show("차변과 대변의 합이 서로 다릅니다.");
                        return;
                    }
                }

                objArrVerification = null;
                objArrVerification = VerifyCvCodAndCvNam(GridViewSlip);
                if (objArrVerification[0].ToString().Equals("N"))
                {
                    XtraMessageBox.Show(string.Format("{0}을(를) 입력하세요.", objArrVerification[2].ToString()));
                    GridViewSlip.FocusedRowHandle = Convert.ToInt32(objArrVerification[1]);
                    GridViewSlip.FocusedColumn = (GridColumn)objArrVerification[4];
                    GridViewSlip.SelectCell(Convert.ToInt32(objArrVerification[1]), (GridColumn)objArrVerification[4]);
                    return;
                }

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.ToString());
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DataTable dtSave = (DataTable)GridSlip.DataSource;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sSlipNo = string.Empty;

                //전표이동 일 경우 기존 데이터 삭제
                if (COPYGB.Equals("2"))
                {
                    string sCopySlipNo = DT_COPY.Rows[0]["SEQNO"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" DELETE FROM ACTRAN ");
                    strSql.AppendLine("       WHERE TDATE = @TDATE ");
                    strSql.AppendLine("         AND ATGUB = @ATGUB ");
                    strSql.AppendLine("         AND SEQNO = @SEQNO ");
                    strSql.AppendLine(" ");

                    cmd.Parameters.Clear();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.AddWithValue("@TDATE", SLIPDT.Replace("-", ""));
                    cmd.Parameters.AddWithValue("@ATGUB", SLIPGB);
                    cmd.Parameters.AddWithValue("@SEQNO", sCopySlipNo);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    strSql.Clear();
                }
                
                if (AddModifyGb.Equals("ADD")) //추가모드
                {
                    sSlipNo = GetSlipNo(cmd, sChkSlipYmd, sChkSlipGb);
                }

                //string sTDATE = dtSave.Rows[0]["TDATE"]?.ToString().Replace("-", "").Substring(0, 8);
                //string sATGUB = dtSave.Rows[0]["ATGUB"]?.ToString();
                //string sSEQNO = dtSave.Rows[0]["SEQNO"]?.ToString();
                string sTDATE = string.Empty;
                string sATGUB = string.Empty;
                string sSEQNO = string.Empty;

                for (int i = 0; i < dtSave.Rows.Count; i++)
                {
                    string sTDate = string.Empty;
                    string sAtGub = string.Empty;
                    string sSeqNo = string.Empty;
                    string sLinNo = string.Empty;
                    string sAcCod = string.Empty;
                    string sAcNam = string.Empty;
                    string sCvCod = string.Empty;
                    string sCvNam = string.Empty;
                    string sAText = string.Empty;
                    string sAcAmt = string.Empty;
                    double dAcAmt = 0;
                    string sAdAmt = string.Empty;
                    double dAdAmt = 0;
                    string sAdpNm = string.Empty;
                    //string sCVGB = string.Empty;
                    string sRk = string.Empty;
                    string sRef1 = string.Empty;
                    string sRef2 = string.Empty;
                    string sRef3 = string.Empty;
                    string sId = FmMainToolBar2.UserID;

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT B.EMP_NM ");
                    strSql.AppendLine("      , B.EMP_ID ");
                    strSql.AppendLine(" 	 , D.DEPT_CD ");
                    strSql.AppendLine("   FROM ZUSRLST A ");
                    strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                    strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
                    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD D ");
                    strSql.AppendLine("     ON D.DEPT_CD = B.REAL_DUTY_DEPT ");
                    strSql.AppendLine("  WHERE A.USRCD = '" + sId + "' ");

                    DataTable dtUser = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    string sEmpId = dtUser.Rows[0]["EMP_ID"]?.ToString();
                    string sDeptCd = dtUser.Rows[0]["DEPT_CD"]?.ToString();

                    bool bExistsYn = false;
                    if (AddModifyGb.Equals("ADD")) //추가모드
                    {
                        sTDate = sChkSlipYmd;
                        sAtGub = sChkSlipGb;
                        sSeqNo = sSlipNo;
                        sLinNo = dtSave.Rows[i]["LINNO"]?.ToString();
                    }
                    else //수정모드
                    {
                        bExistsYn = true;
                        sTDate = dtSave.Rows[0]["TDATE"]?.ToString().Replace("-", "").Substring(0, 8);
                        sAtGub = dtSave.Rows[0]["ATGUB"]?.ToString();
                        sSeqNo = dtSave.Rows[0]["SEQNO"]?.ToString();
                        sLinNo = dtSave.Rows[i]["LINNO"]?.ToString();
                    }
                    
                    sAcCod = dtSave.Rows[i]["ACCOD"]?.ToString();
                    sAcNam = dtSave.Rows[i]["ACNAM"]?.ToString();
                    sCvCod = dtSave.Rows[i]["CVCOD"]?.ToString();
                    if (string.IsNullOrEmpty(sCvCod))
                        sCvCod = null;
                    sCvNam = dtSave.Rows[i]["CVNAM"]?.ToString();
                    sAText = dtSave.Rows[i]["ATEXT"]?.ToString();
                    sAcAmt = dtSave.Rows[i]["ACAMT"]?.ToString();
                    dAcAmt = string.IsNullOrEmpty(sAcAmt) ? 0 : Convert.ToDouble(sAcAmt);
                    sAdAmt = dtSave.Rows[i]["ADAMT"]?.ToString();
                    dAdAmt = string.IsNullOrEmpty(sAdAmt) ? 0 : Convert.ToDouble(sAdAmt);

                    sAdpNm = dtSave.Rows[i]["ADPNM"]?.ToString();
                    //sCVGB = dtSave.Rows[i]["CVGB"]?.ToString();
                    sRk = dtSave.Rows[i]["RK"]?.ToString();
                    
                    if (PARENT_FORM.Name.Equals("AC13001F01"))
                    {
                        sRef1 = ExternalParam["MG_NO"];
                        sRef2 = ExternalParam["MG_HIS_SEQ"];
                        sRef3 = "MAKE_EXPENSE";
                    }

                    /*
                     * 2021-03-10
                     * Reference Key : #0001
                     * 로그 적용을 위하여 이전 데이터 값 조회
                     */
                    int iLogCnt = 0;
                    int iCCNT = 0;
                    //#0002
                    string sSTD_COLS = string.Empty;
                    string sREF_RMK = string.Empty;
                    string sLog_Msg = string.Empty;
                    if (bExistsYn)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT CONVERT(VARCHAR(10),CONVERT(DATE, X1.TDATE), 23) AS TDATE ");
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
                        //strSql.AppendLine("      , X1.CVGB ");
                        strSql.AppendLine("   FROM ACTRAN X1 ");
                        strSql.AppendLine("   LEFT JOIN COM_BASE_CD X2 ");
                        strSql.AppendLine("     ON X1.ATGUB = X2.COM_CD ");
                        strSql.AppendLine("    AND X2.CD_GB = 'AC02001_01' ");
                        strSql.AppendLine("  WHERE TDATE = '" + sTDate + "'");
                        strSql.AppendLine("    AND ATGUB = '" + sAtGub + "'");
                        strSql.AppendLine("    AND SEQNO = '" + sSeqNo + "'");
                        strSql.AppendLine("    AND LINNO = " + sLinNo + "");

                        DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        if(dtPrv.Rows.Count > 0)
                        {
                            string sPrv_TDate = dtPrv.Rows[0]["TDATE"]?.ToString();
                            string sPrv_AtGub = dtPrv.Rows[0]["ATGUB"]?.ToString();
                            string sPrv_AtGub_Nm = dtPrv.Rows[0]["ATGUB_NM"]?.ToString();
                            string sPrv_SeqNo = dtPrv.Rows[0]["SEQNO"]?.ToString();
                            string sPrv_LinNo = dtPrv.Rows[0]["LINNO"]?.ToString();
                            
                            string sPrv_AcCod = dtPrv.Rows[0]["ACCOD"]?.ToString();
                            string sPrv_AcNam = dtPrv.Rows[0]["ACNAM"]?.ToString();
                            string sPrv_CvCod = dtPrv.Rows[0]["CVCOD"]?.ToString();
                            string sPrv_CvNam = dtPrv.Rows[0]["CVNAM"]?.ToString();
                            string sPrv_AText = dtPrv.Rows[0]["ATEXT"]?.ToString();
                            //string sPrv_CVGB = dtPrv.Rows[0]["CVGB"]?.ToString();

                            double dPrv_AcAmt = 0;
                            double.TryParse(dtPrv.Rows[0]["ACAMT"]?.ToString(), out dPrv_AcAmt);

                            double dPrv_AdAmt = 0;
                            double.TryParse(dtPrv.Rows[0]["ADAMT"]?.ToString(), out dPrv_AdAmt);

                            string sPrv_AdpCd = dtPrv.Rows[0]["ADPCD"]?.ToString();
                            string sPrv_AdpNm = dtPrv.Rows[0]["ADPNM"]?.ToString();

                            sSTD_COLS = string.Format("{0}/{1}/전표번호:{2}/순번:{3}"
                                , sPrv_TDate, sPrv_AtGub_Nm, sPrv_SeqNo, sPrv_LinNo);

                            if (!sAcCod.Equals(sPrv_AcCod))
                            {
                                if (iCCNT++ != 0)
                                    sLog_Msg += " | ";

                                sLog_Msg += string.Format("계정코드 : {0} ▶ {1}", sPrv_AcNam, sAcNam);
                                iLogCnt++;
                            }

                            if (!sCvCod.Equals(sPrv_CvCod))
                            {
                                if (iCCNT++ != 0)
                                    sLog_Msg += " | ";

                                sLog_Msg += string.Format("거래처 : {0} ▶ {1}", sPrv_CvNam, sCvNam);
                                iLogCnt++;
                            }

                            if (dAcAmt != dPrv_AcAmt)
                            {
                                if (iCCNT++ != 0)
                                    sLog_Msg += " | ";

                                sLog_Msg += string.Format("차변 : {0} ▶ {1}", dPrv_AcAmt, dAcAmt);
                                iLogCnt++;
                            }

                            if (dAdAmt != dPrv_AdAmt)
                            {
                                if (iCCNT++ != 0)
                                    sLog_Msg += " | ";

                                sLog_Msg += string.Format("대변 : {0} ▶ {1}", dPrv_AdAmt, dAdAmt);
                                iLogCnt++;
                            }

                            if (!sAText.Equals(sPrv_AText))
                            {
                                if (iCCNT++ != 0)
                                    sLog_Msg += " | ";

                                sLog_Msg += string.Format("적요 : {0} ▶ {1}", sPrv_AText, sAText);
                                iLogCnt++;
                            }

                            sREF_RMK += string.Format("TABLE : ACTRAN / TDATE : {0}, ATGUB : {1}, SEQNO : {2}, LINNO {3} ", sPrv_TDate.Replace("-", ""), sPrv_AtGub, sPrv_SeqNo, sPrv_LinNo);
                            //if (sDeptCd.Equals(sPrv_AdpCd))
                            //{
                            //    sLog_Msg += string.Format(", 대변 {0} ▶ {1} ", sPrv_AdpNm, sAdpNm);
                            //    iLogCnt++;
                            //}
                        }
                    }

                    sLog_Msg = sLog_Msg.Length > 500 ? sLog_Msg.Substring(0, 500) : sLog_Msg;

                    sTDATE = sTDate;
                    sATGUB = sAtGub;
                    sSEQNO = sSeqNo;

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO ACTRAN ");
                    //strSql.AppendLine("           ( TDATE ");
                    //strSql.AppendLine("           , ATGUB ");
                    //strSql.AppendLine("           , SEQNO ");
                    //strSql.AppendLine("           , LINNO ");
                    //strSql.AppendLine("           , ACCOD ");
                    //strSql.AppendLine("           , ACNAM ");
                    //strSql.AppendLine("           , ATEXT ");
                    //strSql.AppendLine("           , CVCOD ");
                    //strSql.AppendLine("           , CVNAM ");
                    //strSql.AppendLine("           , ACAMT ");
                    //strSql.AppendLine("           , ADAMT ");
                    //strSql.AppendLine("           , RK ");
                    //strSql.AppendLine("           , REF1 ");
                    //strSql.AppendLine("           , REF2 ");
                    //strSql.AppendLine("           , REF3 ");
                    //strSql.AppendLine("           , CUSER "); 
                    //strSql.AppendLine("           , ADPCD ");
                    //strSql.AppendLine("           , ADPNM ");
                    //strSql.AppendLine("           , CDATE )");
                    //strSql.AppendLine("     VALUES( @TDATE ");
                    //strSql.AppendLine("           , @ATGUB ");
                    //strSql.AppendLine("           , @SEQNO ");
                    //strSql.AppendLine("           , @LINNO ");
                    ////strSql.AppendLine("           ,  " + sLinNo + " ");
                    //strSql.AppendLine("           , @ACCOD ");
                    //strSql.AppendLine("           , (SELECT ACNAM     ");
                    //strSql.AppendLine("                FROM ACMSTF    ");
                    //strSql.AppendLine("               WHERE ACCOD = @ACCOD )");
                    //strSql.AppendLine("           , '" + sAText + "' ");
                    //strSql.AppendLine("           , @CVCOD ");
                    //strSql.AppendLine("           , '" + sCvNam + "' ");
                    //strSql.AppendLine("           , " + dAcAmt + " ");
                    //strSql.AppendLine("           , " + dAdAmt + " ");
                    //strSql.AppendLine("           , '" + sRk + "' ");
                    //strSql.AppendLine("           , '" + sRef1 + "' ");
                    //strSql.AppendLine("           , '" + sRef2 + "' ");
                    //strSql.AppendLine("           , '" + sRef3 + "' ");
                    //strSql.AppendLine("           , '" + sEmpId + "' ");
                    //strSql.AppendLine("           , '" + sDeptCd + "' ");
                    //strSql.AppendLine("           , ( SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = '" + sDeptCd + "' ) ");
                    //strSql.AppendLine("           , NOW() ) ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine(" 		    ACCOD = @ACCOD ");
                    //strSql.AppendLine("           , ACNAM = ( SELECT ACNAM     ");
                    //strSql.AppendLine("                         FROM ACMSTF    ");
                    //strSql.AppendLine("                        WHERE ACCOD = @ACCOD )");
                    //strSql.AppendLine("           , ATEXT = '" + sAText + "' ");
                    //strSql.AppendLine("           , CVCOD = @CVCOD ");
                    //strSql.AppendLine("           , CVNAM = '" + sCvNam + "' ");
                    //strSql.AppendLine("           , ACAMT = " + dAcAmt + " ");
                    //strSql.AppendLine("           , ADAMT = " + dAdAmt + " ");
                    //strSql.AppendLine("           , RK = '" + sRk + "' ");
                    //strSql.AppendLine("           , MUSER = '" + sEmpId + "' ");
                    //strSql.AppendLine("           , ADPCD = '" + sDeptCd + "' ");
                    //strSql.AppendLine("           , ADPNM = ( SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = '" + sDeptCd + "' ) ");
                    //strSql.AppendLine("           , MDATE = NOW() ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM ACTRAN WHERE TDATE = @TDATE AND ATGUB = @ATGUB AND SEQNO = @SEQNO AND LINNO = @LINNO)            ");
                    strSql.AppendLine("   BEGIN                                                                                               ");
                    strSql.AppendLine("         UPDATE ACTRAN                                                                                 ");
                    strSql.AppendLine("            SET ACCOD = @ACCOD                                                                         ");
	                strSql.AppendLine("              , ACNAM = (SELECT ACNAM                                                                  ");
                    strSql.AppendLine("                           FROM ACMSTF                                                                 ");
                    strSql.AppendLine("                          WHERE ACCOD = @ACCOD )                                                       ");
	                strSql.AppendLine("              , ATEXT = '" + sAText + "'                                                               ");
	                strSql.AppendLine("              , CVCOD = @CVCOD                                                                         ");
	                strSql.AppendLine("              , CVNAM = '" + sCvNam + "'                                                               ");
	                strSql.AppendLine("              , ACAMT = " + dAcAmt + "                                                                 ");
	                strSql.AppendLine("              , ADAMT = " + dAdAmt + "                                                                 ");
	                strSql.AppendLine("              , RK = '" + sRk + "'                                                                     ");
	                strSql.AppendLine("              , MUSER = '" + sEmpId + "'                                                               ");
	                strSql.AppendLine("              , ADPCD = '" + sDeptCd + "'                                                              ");
	                strSql.AppendLine("              , ADPNM = (SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = '" + sDeptCd + "' )  ");
                    //strSql.AppendLine("              , CVGB = '" + sCVGB + "'");
                    

                    strSql.AppendLine("              , MDATE = CONVERT(VARCHAR(19), GETDATE(), 20)                                            ");
                    strSql.AppendLine("          WHERE TDATE = @TDATE AND ATGUB = @ATGUB AND SEQNO = @SEQNO AND LINNO = @LINNO               ");
                    strSql.AppendLine("     END                                                                                               ");
                    strSql.AppendLine("ELSE                                                                                                   ");
                    strSql.AppendLine("   BEGIN                                                                                               ");
                    strSql.AppendLine("         INSERT INTO ACTRAN                                                                            ");
                    strSql.AppendLine("              (TDATE                                                                                   ");
                    strSql.AppendLine("              , ATGUB                                                                                  ");
                    strSql.AppendLine("              , SEQNO                                                                                  ");
                    strSql.AppendLine("              , LINNO                                                                                  ");
                    strSql.AppendLine("              , ACCOD                                                                                  ");
                    strSql.AppendLine("              , ACNAM                                                                                  ");
                    strSql.AppendLine("              , ATEXT                                                                                  ");
                    strSql.AppendLine("              , CVCOD                                                                                  ");
                    strSql.AppendLine("              , CVNAM                                                                                  ");
                    strSql.AppendLine("              , ACAMT                                                                                  ");
                    strSql.AppendLine("              , ADAMT                                                                                  ");
                    //strSql.AppendLine("              , CVGB                                                                                  ");
                    strSql.AppendLine("              , RK                                                                                     ");
                    strSql.AppendLine("              , REF1                                                                                   ");
                    strSql.AppendLine("              , REF2                                                                                   ");
                    strSql.AppendLine("              , REF3                                                                                   ");
                    strSql.AppendLine("              , CUSER                                                                                  ");
                    strSql.AppendLine("              , ADPCD                                                                                  ");
                    strSql.AppendLine("              , ADPNM                                                                                  ");
                    strSql.AppendLine("              , CDATE )                                                                                ");
                    strSql.AppendLine("        VALUES(@TDATE                                                                                  ");
                    strSql.AppendLine("              , @ATGUB                                                                                 ");
                    strSql.AppendLine("              , @SEQNO                                                                                 ");
                    strSql.AppendLine("              , @LINNO                                                                                 ");
                    strSql.AppendLine("              , @ACCOD                                                                                 ");
                    strSql.AppendLine("              , (SELECT ACNAM                                                                          ");
                    strSql.AppendLine("                   FROM ACMSTF                                                                         ");
                    strSql.AppendLine("                  WHERE ACCOD = @ACCOD)                                                                ");
	                strSql.AppendLine("              , '" + sAText + "'                                                                       ");
	                strSql.AppendLine("              , @CVCOD                                                                                 ");
	                strSql.AppendLine("              , '" + sCvNam + "'                                                                       ");
	                strSql.AppendLine("              , " + dAcAmt + "                                                                         ");
	                strSql.AppendLine("              , " + dAdAmt + "                                                                         ");
                    //strSql.AppendLine("              , '" + sCVGB + "'                                                                          ");
                    strSql.AppendLine("              , '" + sRk + "'                                                                          ");
	                strSql.AppendLine("              , '" + sRef1 + "'                                                                        ");
	                strSql.AppendLine("              , '" + sRef2 + "'                                                                        ");
	                strSql.AppendLine("              , '" + sRef3 + "'                                                                        ");
	                strSql.AppendLine("              , '" + sEmpId + "'                                                                       ");
	                strSql.AppendLine("              , '" + sDeptCd + "'                                                                      ");
	                strSql.AppendLine("              , (SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = '" + sDeptCd + "' )          ");
	                strSql.AppendLine("              , CONVERT(VARCHAR(19), GETDATE(), 20))                                                   ");
                    strSql.AppendLine("     END                                                                                               ");

                    cmd.Parameters.Clear();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.AddWithValue("@CVCOD", sCvCod);
                    cmd.Parameters.AddWithValue("@TDATE", sTDate);
                    cmd.Parameters.AddWithValue("@ATGUB", sAtGub);
                    cmd.Parameters.AddWithValue("@SEQNO", sSeqNo);
                    cmd.Parameters.AddWithValue("@ACCOD", sAcCod);
                    cmd.Parameters.AddWithValue("@LINNO", Convert.ToInt32(sLinNo));
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    
                    if (PARENT_FORM.Name.Equals("AC13001F01"))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" UPDATE EQUIP_CD_HISTORY ");
                        strSql.AppendLine("    SET SLIP_YN = 'Y' ");
                        strSql.AppendLine("  WHERE MG_NO = '" + ExternalParam["MG_NO"] + "' ");
                        strSql.AppendLine("    AND MG_HIS_SEQ = " + ExternalParam["MG_HIS_SEQ"] + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    //LogInsert
                    //Reference : #0001
                    if (iLogCnt > 0)
                    {
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
                        cmd.Parameters.AddWithValue("@PGM_ID", PARENT_FORM.Name);
                        cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                        cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                        cmd.Parameters.AddWithValue("@STD_COLS", sSTD_COLS);
                        cmd.Parameters.AddWithValue("@EDIT_RMK", sLog_Msg);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }

                }
                
                /*
                 * 2021-01-27
                 * 입출금전표 등록 시 현금이 자동계산되는 부분에서 다시 원래대로 현금을 상대계정으로 넣지않는 것으로 수정
                 */
                #region[이전 로직]

                ///*
                // * 2021-01-18 현업요청
                // * 입/출금전표(ATGUB : 1, 2)는 등록 시 현금이 자동계산되어 차대에 들어가야함
                // */
                //if(sChkSlipGb.Equals("1") || sChkSlipGb.Equals("2"))
                //{
                //    //현금을 제외한 출금 (ATGUB, 1) (ADAMT) 합계 , 입금 (ATGUB, 2)(ACAMT) 합계
                //    strSql.Clear();
                //    strSql.AppendLine(" SELECT CASE '" + sChkSlipGb + "' WHEN '1' THEN IFNULL(SUM(IFNULL(ACAMT, 0)), 0) ");
                //    strSql.AppendLine("                                           ELSE IFNULL(SUM(IFNULL(ADAMT, 0)), 0) END AS SUMMARY_AMT  ");
                //    strSql.AppendLine("   FROM ACTRAN A ");
                //    strSql.AppendLine("  WHERE TDATE = @TDATE ");
                //    strSql.AppendLine("    AND ATGUB = @ATGUB ");
                //    strSql.AppendLine("    AND SEQNO = @SEQNO ");
                //    strSql.AppendLine("    AND ACCOD <> '" + ComnEtcFunc.CashCode + "' ");

                //    DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                //    int iSummaryAmt = 0;
                //    if (dt.Rows.Count > 0)
                //    {
                //        int.TryParse(dt.Rows[0]["SUMMARY_AMT"]?.ToString(), out iSummaryAmt);
                //    }

                //    //현금계정이 속한 행번(LINNO)찾고 없다면 채번
                //    strSql.Clear();
                //    strSql.AppendLine(" ");
                //    strSql.AppendLine(" SELECT LINNO ");
                //    strSql.AppendLine("   FROM ACTRAN ");
                //    strSql.AppendLine("  WHERE TDATE = @TDATE ");
                //    strSql.AppendLine("    AND ATGUB = @ATGUB ");
                //    strSql.AppendLine("    AND SEQNO = @SEQNO ");
                //    strSql.AppendLine("    AND ACCOD = '" + ComnEtcFunc.CashCode + "' ");

                //    string sLINNO = string.Empty;
                //    dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                //    if (dt.Rows.Count > 0)
                //    {
                //        sLINNO = dt.Rows[0]["LINNO"]?.ToString();
                //    }
                //    else
                //    {
                //        strSql.Clear();
                //        strSql.AppendLine(" ");
                //        strSql.AppendLine(" SELECT IFNULL(MAX(LINNO), 0) + 1 AS MAX_VAL ");
                //        strSql.AppendLine("   FROM ACTRAN ");
                //        strSql.AppendLine("  WHERE TDATE = @TDATE ");
                //        strSql.AppendLine("    AND ATGUB = @ATGUB ");
                //        strSql.AppendLine("    AND SEQNO = @SEQNO ");

                //        dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                //        sLINNO = dt.Rows[0]["MAX_VAL"]?.ToString();
                //    }

                //    strSql.Clear();
                //    strSql.AppendLine(" ");
                //    strSql.AppendLine(" SELECT D.DEPT_CD ");
                //    strSql.AppendLine("   FROM ZUSRLST A ");
                //    strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                //    strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
                //    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD D ");
                //    strSql.AppendLine("     ON D.DEPT_CD = B.REAL_DUTY_DEPT ");
                //    strSql.AppendLine("  WHERE A.USRCD = '" + FmMainToolBar2.UserID + "' ");

                //    DataTable dtUser = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                //    string sDeptCd = dtUser.Rows[0]["DEPT_CD"]?.ToString();

                //    Dictionary<string, string> dicMySqlParam = new Dictionary<string, string>();

                //    dicMySqlParam.Add("TDATE", sTDATE);
                //    dicMySqlParam.Add("ATGUB", sATGUB);
                //    dicMySqlParam.Add("SEQNO", sSEQNO);
                //    dicMySqlParam.Add("LINNO", sLINNO);
                //    dicMySqlParam.Add("ACCOD", ComnEtcFunc.CashCode);
                //    dicMySqlParam.Add("CVCOD", "-1");
                //    dicMySqlParam.Add("CVNAM", "");
                //    //출금시 대변에 세팅, 입금시 차변에 세팅
                //    if (sChkSlipGb.Equals("1"))
                //    {
                //        dicMySqlParam.Add("ACAMT", "0");
                //        dicMySqlParam.Add("ADAMT", iSummaryAmt.ToString());
                //    }
                //    else if (sChkSlipGb.Equals("2"))
                //    {
                //        dicMySqlParam.Add("ACAMT", iSummaryAmt.ToString());
                //        dicMySqlParam.Add("ADAMT", "0");
                //    }

                //    dicMySqlParam.Add("ADPCD", sDeptCd);
                //    dicMySqlParam.Add("MUSER", FmMainToolBar2.UserID);
                //    dicMySqlParam.Add("APVYN", "N");
                //    dicMySqlParam.Add("CUSER", FmMainToolBar2.UserID);

                //    //현금계정 INSERT / UPDATE
                //    strSql.Clear();
                //    strSql.AppendLine(" ");
                //    strSql.AppendLine(" IF ( SELECT 1 = 1 FROM ACTRAN WHERE TDATE = @TDATE AND ATGUB = @ATGUB AND SEQNO = @SEQNO AND LINNO = @LINNO ) THEN ");
                //    strSql.AppendLine("    BEGIN  ");
                //    strSql.AppendLine("          UPDATE ACTRAN ");
                //    strSql.AppendLine("             SET ACCOD = @ACCOD ");
                //    strSql.AppendLine("               , ACNAM = ( SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = @ACCOD ) ");
                //    strSql.AppendLine("               , CVCOD = @CVCOD ");
                //    strSql.AppendLine("               , CVNAM = @CVNAM ");
                //    strSql.AppendLine("               , ATEXT = ( SELECT GROUP_CONCAT(ATEXT) FROM ACTRAN X1 WHERE X1.TDATE = @TDATE AND X1.ATGUB = @ATGUB AND X1.SEQNO = @SEQNO AND X1.ACCOD <> '0101' GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO ) ");
                //    strSql.AppendLine("               , ADAMT = @ADAMT ");
                //    strSql.AppendLine("               , ACAMT = @ACAMT ");
                //    strSql.AppendLine("               , ADPCD = @ADPCD ");
                //    strSql.AppendLine("               , ADPNM = ( SELECT X1.DEPT_CD FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = @ADPCD ) ");
                //    strSql.AppendLine("               , MUSER = @MUSER ");
                //    strSql.AppendLine("               , MDATE = NOW() ");
                //    strSql.AppendLine("           WHERE TDATE = @TDATE  ");
                //    strSql.AppendLine("             AND ATGUB = @ATGUB  ");
                //    strSql.AppendLine("             AND SEQNO = @SEQNO  ");
                //    strSql.AppendLine("             AND LINNO = @LINNO; ");
                //    strSql.AppendLine("      END; ");
                //    strSql.AppendLine(" ELSE ");
                //    strSql.AppendLine("    BEGIN ");
                //    strSql.AppendLine("          INSERT INTO ACTRAN ");
                //    strSql.AppendLine("                    ( TDATE ");
                //    strSql.AppendLine("                    , ATGUB ");
                //    strSql.AppendLine("                    , SEQNO ");
                //    strSql.AppendLine("                    , LINNO ");
                //    strSql.AppendLine("                    , ACCOD ");
                //    strSql.AppendLine("                    , ACNAM ");
                //    strSql.AppendLine("                    , CVCOD ");
                //    strSql.AppendLine("                    , CVNAM ");
                //    strSql.AppendLine("                    , ATEXT ");
                //    strSql.AppendLine("                    , ADAMT ");
                //    strSql.AppendLine("                    , ACAMT ");
                //    strSql.AppendLine("                    , ADPCD ");
                //    strSql.AppendLine("                    , ADPNM ");
                //    strSql.AppendLine("                    , APVYN ");
                //    strSql.AppendLine("                    , CUSER ");
                //    strSql.AppendLine("                    , CDATE ) ");
                //    strSql.AppendLine("              VALUES( @TDATE ");
                //    strSql.AppendLine("                    , @ATGUB ");
                //    strSql.AppendLine("                    , @SEQNO ");
                //    strSql.AppendLine("                    , @LINNO ");
                //    strSql.AppendLine("                    , @ACCOD ");
                //    strSql.AppendLine("                    , ( SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = @ACCOD ) ");
                //    strSql.AppendLine("                    , @CVCOD ");
                //    strSql.AppendLine("                    , @CVNAM ");
                //    strSql.AppendLine("                    , ( SELECT GROUP_CONCAT(X1.ATEXT) FROM ACTRAN X1 WHERE X1.TDATE = @TDATE AND X1.ATGUB = @ATGUB AND X1.SEQNO = @SEQNO AND X1.ACCOD <> '0101' GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO ) ");
                //    strSql.AppendLine("                    , @ADAMT ");
                //    strSql.AppendLine("                    , @ACAMT ");
                //    strSql.AppendLine("                    , @ADPCD ");
                //    strSql.AppendLine("                    , ( SELECT X1.DEPT_CD FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = @ADPCD ) ");
                //    strSql.AppendLine("                    , @APVYN ");
                //    strSql.AppendLine("                    , @CUSER ");
                //    strSql.AppendLine("                    , NOW() ); ");
                //    strSql.AppendLine("      END; ");
                //    strSql.AppendLine(" END IF; ");

                //    cmd.Parameters.Clear();
                //    cmd.CommandType = CommandType.Text;
                //    cmd.CommandText = strSql.ToString();
                //    foreach (KeyValuePair<string, string> param in dicMySqlParam)
                //    {
                //        cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                //    }
                //    //foreach(KeyValuePair<string, MySqlParameter> param in dicMySqlParam)
                //    //{
                //    //    cmd.Parameters.Add(param.Value);
                //    //}
                //    cmd.ExecuteNonQuery();
                //    cmd.Parameters.Clear();
                //}

                #endregion[이전 로직]

                /*
                 * 2021-01-06
                 * 지불관리(AAUTO : 'D01')에서 올라온 데이터 수정 시 SYNC 맞춤
                 * REF1은 SUMGF의 PK인 SLIPNO와 매칭
                 */
                if (_SYNC == SYNC_GB.Jibul)
                {
                    string sREF1 = DrSlipInfo["REF1"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");

                    #region mariaDB
                    //strSql.AppendLine(" SELECT IF(ACAMT > ADAMT, ACAMT, ADAMT) AS TRSUM ");
                    //strSql.AppendLine("   FROM ( SELECT SUM(ACAMT) AS ACAMT ");
                    //strSql.AppendLine("               , SUM(ADAMT) AS ADAMT ");
                    //strSql.AppendLine("            FROM ACTRAN A ");
                    //strSql.AppendLine("           WHERE A.REF1 = '" + sREF1 + "' ");
                    //strSql.AppendLine("           GROUP BY A.REF1 ) X1 ");
                    #endregion

                    strSql.AppendLine("SELECT CASE WHEN ACAMT > ADAMT THEN ACAMT ELSE ADAMT END AS TRSUM");
                    strSql.AppendLine("  FROM(SELECT SUM(ACAMT) AS ACAMT                                ");
                    strSql.AppendLine("              , SUM(ADAMT) AS ADAMT                              ");
                    strSql.AppendLine("           FROM ACTRAN A                                         ");
                    strSql.AppendLine("          WHERE A.REF1 = '"+ sREF1 + "'                          ");
                    strSql.AppendLine("          GROUP BY A.REF1) X1                                    ");

                   DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    int iTRSUM = 0;
                    int.TryParse(dtChk.Rows[0]["TRSUM"]?.ToString(), out iTRSUM);

                    if (!string.IsNullOrEmpty(sREF1))
                    {
                        //변수세팅
                        dicParams.Clear();
                        dicParams.Add("TDATE", sTDATE);
                        dicParams.Add("TRSUM", iTRSUM.ToString());
                        dicParams.Add("SLIPNO", sREF1);

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE SUGMF ");
                        strSql.AppendLine("    SET TDATE = CONVERT(VARCHAR(10), CONVERT([DATETIME],@TDATE), 23) ");
                        strSql.AppendLine("      , TRSUM = @TRSUM ");
                        strSql.AppendLine("  WHERE SLIPNO = @SLIPNO ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        foreach (KeyValuePair<string, string> param in dicParams)
                        {
                            cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                        }
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                dicParams.Clear();
                dicParams.Add("TDATE", sTDATE);
                dicParams.Add("ATGUB", sATGUB);
                dicParams.Add("SEQNO", sSEQNO);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
                sSaveGb = "ERROR";
            }
            finally
            {
                Cursor = Cursors.Default;
                if (sSaveGb.Equals("ERROR"))
                {

                }
                else if (sSaveGb.Equals("SAVE"))
                {
                    DataRowSendEvent(dicParams);
                    Dispose();
                    //DialogResult = DialogResult.OK;
                }
                else if (sSaveGb.Equals("CONTINUE"))
                {
                    SettingInitializedControl();
                }
            }
        }
        
        #endregion[Get Data By Query]

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

        #region[GridView Row's Stripe Pattern]

        private void GridViewSlip_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewSlip_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            //if(e.RowHandle < 0)
            //{
            //    ResourceManager rm = Resources.ResourceManager;
            //    Bitmap myImage = (Bitmap)rm.GetObject("GuideImage");
                
            //    e.Graphics.DrawImage((Image)myImage);
            //}
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
            
        }

        #endregion[GridView Row's Stripe Pattern]

        private void RepoBtnEditAccCd_Leave(object sender, EventArgs e)
        {
            
        }

        private void GetAccountInfo(int iRowHandle, string sAccCd)
        {

        }

        private void SetLookUpEdit(DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit lkup, string sGb, string sNullYn)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("  UNION ALL");
            }
            else if (sNullYn.Equals("N"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , '' AS SUB ");
                strSql.AppendLine("  UNION ALL");
            }

            if (sGb.Equals("1"))//계정명
            {
                strSql.AppendLine(" SELECT A1.ACCOD AS CD");
                strSql.AppendLine("      , A1.ACNAM AS NM ");
                strSql.AppendLine("      , A2.ACNAM AS SUB ");
                strSql.AppendLine("   FROM ACMSTF A1   ");
                strSql.AppendLine("   LEFT JOIN ACTOPF A2  ");
                strSql.AppendLine("     ON A1.ACCOD BETWEEN AFROM AND ATO  ");
                strSql.AppendLine("    AND SEQNO <> '0'  ");
                strSql.AppendLine("    AND USEYN = 'Y' ");
            }
            else if (sGb.Equals("2"))//거래처명
            {
                strSql.AppendLine(" SELECT DEALER_CD AS CD  ");
                strSql.AppendLine(" 	 , DEALER_NM AS NM  ");
                strSql.AppendLine("   FROM ACC_DEALER_CD    ");
            }

            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            lkup.DataSource = dt;
            lkup.DisplayMember = "NM";
            lkup.ValueMember = "CD";
        }
        
        private void RepoLkupAcNam_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEdit lkupEdit = (LookUpEdit)sender;
            string sAccod = lkupEdit.EditValue.ToString();

            GridViewSlip.SetFocusedRowCellValue("ACCOD", sAccod);
        }

        private void RepoLkupAcNam_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewSlip.SelectedRowsCount < 1)
                return;

            if (e.KeyCode == Keys.Enter)
            {
                LookUpEdit lkupEdit = (LookUpEdit)sender;
                string sAccod = lkupEdit.EditValue.ToString();

                GridViewSlip.SetFocusedRowCellValue("ACCOD", sAccod);
            }
        }

        private void RepoLkupDealerNm_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEdit lkupEdit = (LookUpEdit)sender;
            string sDealerCd = lkupEdit.EditValue.ToString();

            GridViewSlip.SetFocusedRowCellValue("CVCOD", sDealerCd);
        }

        private void RepoLkupDealerNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (GridViewSlip.SelectedRowsCount < 1)
                return;

            if (e.KeyCode == Keys.Enter)
            {
                LookUpEdit lkupEdit = (LookUpEdit)sender;
                string sDealerCd = lkupEdit.EditValue.ToString();

                GridViewSlip.SetFocusedRowCellValue("CVCOD", sDealerCd);
            }
        }

        private void GridViewSlip_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                GridView view = sender as GridView;
                GridHitInfo hitInfo = view.CalcHitInfo(e.Location);
                if (hitInfo.InRow || hitInfo.InColumnPanel)
                {
                    DevExpress.Utils.DXMouseEventArgs args = DevExpress.Utils.DXMouseEventArgs.GetMouseArgs(e);
                    if (args != null) args.Handled = true;
                    view.FocusedRowHandle = hitInfo.RowHandle;
                    view.FocusedColumn = hitInfo.Column;
                    //show context menu here  
                }
            }
        }

        private void RepoTxtAcAmt_Leave(object sender, EventArgs e)
        {
            SetPriceDifferency();
        }

        private void RepoTxtAdAmt_Leave(object sender, EventArgs e)
        {
            SetPriceDifferency();
        }

        private void SetPriceDifferency()
        {
            double iAcAmtSummary = 0;
            double iAdAmtSummary = 0;

            for (int i = 0; i < GridViewSlip.RowCount; i++)
            {
                double iAcAmt = 0;
                double iAdAmt = 0;

                double.TryParse(GridViewSlip.GetRowCellValue(i, GridColAcAmt)?.ToString(), out iAcAmt);
                double.TryParse(GridViewSlip.GetRowCellValue(i, GridColAdAmt)?.ToString(), out iAdAmt);

                iAcAmtSummary += iAcAmt;
                iAdAmtSummary += iAdAmt;
            }

            double dDiff = iAcAmtSummary - iAdAmtSummary;
            LblDiff.Text = String.Format("{0:#,0}", dDiff);
            if(dDiff == 0)
            {
                LblDiff.ForeColor = Color.Blue;
            }
            else
            {
                LblDiff.ForeColor = Color.Red;
            }
            
        }

        private void RepoCboCvgb_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Space)
            {
                ComboBoxEdit cbo = (ComboBoxEdit)sender;

                if (cbo.SelectedIndex == 0)
                    cbo.SelectedIndex = 1;
                else
                    cbo.SelectedIndex = 0;
            }
        }
    }
}