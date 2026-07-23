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
using DevExpress.XtraBars;
using System.Data.SqlClient;

/*
 * 1. 2021-02-01(현업요청)
 *    기존 -> 전표일자는 From 일자를 조회때마다 세팅
 *    수정 -> Load 시에만 오늘날짜로 세팅
 *    
 * 1. 2021-02-07 (현업요청)
 *    거래처 초성검색 추가
 *    
 *    
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2021-10-19
 * 수정자   : 정은영
 * 수정내용 : 속도이슈로 인해 조회 거래처 전체 검색 쿼리수정
 * 참조 ID  : #00001
 * 
 * 수정일자 : 2022-10-13
 * 수정자   : 정은영
 * 수정내용 : 1. 지불전표 적요 수정 -> 거래처명,지불금액,외상대지불
 *            2. 은행조회 은행구분으로 변경
 * ID       : #00002
 */
namespace AccAdm
{
    public partial class AC16001F00 : DevExpress.XtraEditors.XtraForm
    {
        public AC16001F00()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC16001F00_Load(object sender, EventArgs e)
        {
            //ComnEtcFunc.SetDateFromToValue(DateEditFrom, DateEditTo);
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            DateEditTDate.EditValue = DateEditFrom.EditValue;
            TxtAText.EditValue = "외상대 지불";
            
            SetLookup(); //초기 LookupEdit 세팅, 메소드 참조
            LkupPmntGb.EditValue = "0001"; //지급방법 초기 세팅 (송금->보통예금)
            
            InitControls();
            UpdateDropDownButton(BtnOne);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr, GridViewRetr2, GridViewRetr3 };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout2.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl2.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout3.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl3.RestoreLayoutFromXml(sFile);
            }

            if (CboFindSbj.SelectedIndex < 0)
                CboFindSbj.SelectedIndex = 0;
            if (string.IsNullOrEmpty(RdgbBalYn.EditValue?.ToString()))
                RdgbBalYn.SelectedIndex = 1;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            /*
             * 2021-02-01(현업요청)
             * 전표일자는 오늘날짜로 Load시에만 세팅하고 조회할때마다 전표일자를 바꾸는 것은 안됨
             */
            //DateEditTDate.EditValue = DateEditFrom.EditValue;
            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("전표일자를 입력하세요.");
                DateEditFrom.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("전표일자를 입력하세요.");
                DateEditTo.Focus();
                return;
            }

            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                XtraMessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }

            
            string sFindIdx = CboFindSbj.SelectedIndex.ToString();
            string sFindWord = TxtFindWord.EditValue?.ToString().Trim();
            string sBalYn = RdgbBalYn.EditValue?.ToString();
            string sPmntGb = CboPmntTarget.EditValue?.ToString();
            
            if (sFindIdx.Equals("1")) //사업자번호의 경우 "-" 제외한다
                sFindWord.Replace("-", "");

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("DATE_F", sYmdFrom);
            dicParams.Add("DATE_T", sYmdTo);
            dicParams.Add("FIND_IDX", sFindIdx);
            dicParams.Add("FIND_WORD", sFindWord);
            dicParams.Add("BAL_YN", sBalYn);

            //결제대상의 값에 따라 계정코드 따로 부여 HARDCODING
            if (sPmntGb.Equals("외상매입금"))
                dicParams.Add("ACCOD", "0251");
            else if(sPmntGb.Equals("미지급금"))
                dicParams.Add("ACCOD", "0253");

            if (TabControl.SelectedTabPage == TabPagePurc)
            {
                if (sPmntGb.Equals("외상매입금"))
                {
                    LayoutGrpText.Text = "미지급금 리스트 ( 해당 기간에 발생된 거래처별 외상매입금의 총액을 나타냅니다. )";
                    TabPagePurc.Text = "외상매입금 합계";
                    GridRetr.MainView = GridViewRetr;
                }
                else if (sPmntGb.Equals("미지급금"))
                {
                    LayoutGrpText.Text = "미지급금 리스트 ( 해당 기간에 발생된 거래처별 미지급금의 총액을 나타냅니다. )";
                    TabPagePurc.Text = "미지급금 합계";
                    GridRetr.MainView = GridViewRetr3;
                }

                GridRetr.DataSource = GetInfo(dicParams);
                if (GridViewRetr.RowCount > 0)
                {
                    GridViewRetr.Focus();
                }
                else
                {
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }
            }
            else if (TabControl.SelectedTabPage == TabPageSugm)
            {
                dicParams["DATE_F"] = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                dicParams["DATE_T"] = DateEditTo.EditValue?.ToString().Substring(0, 10);
                dicParams.Remove("BAL_YN"); //BAL_YN은 해당 쿼리에서 사용되지 않음
                dicParams.Remove("ACCOD");
                
                GridRetr2.DataSource = GetPaymentInfo(dicParams);
                if (GridViewRetr2.RowCount > 0)
                {
                    GridViewRetr2.Focus();
                }
                else
                {
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                }
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            AC16001F01 frm = new AC16001F01();
            frm.P_AC16001F00 = this;
            frm.AddAndMdi = "ADD";
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }

        private void GridViewRetr2_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                AC16001F01 frm = new AC16001F01();
                frm.P_AC16001F00 = this;
                frm.DR_INFO = GridViewRetr2.GetFocusedDataRow();
                frm.AddAndMdi = "MOD";
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr.PerformClick();
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //외상매입금 합계 탭
            if (TabControl.SelectedTabPage != TabPagePurc)
            {
                XtraMessageBox.Show("외상매입금 합계 탭에서만 가능한 버튼입니다.");
                return;
            }

            string sPmntTarget = CboPmntTarget.EditValue?.ToString(); //결제대상
            string sTDate = DateEditTDate.EditValue?.ToString().Substring(0, 10); //전표일자
            string sPmntGb = LkupPmntGb.EditValue?.ToString(); //지급방법
            string sAcnt = LkupAcnt.EditValue?.ToString(); //계좌
            string sAcCod1 = string.Empty; //적용계정코드 (대변)
            string sAcCod2 = string.Empty; //적용계정코드 (차변)
            string sJGubn = string.Empty; //결제대상

            //필수입력 값 체크
            if (string.IsNullOrEmpty(sPmntTarget))
            {
                XtraMessageBox.Show("결제대상을 선택하세요.");
                CboPmntTarget.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sTDate))
            {
                XtraMessageBox.Show("전표일자를 입력하세요.");
                DateEditTDate.SelectAll();
                DateEditTDate.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sPmntGb))
            {
                XtraMessageBox.Show("지급방법을 선택하세요.");
                LkupPmntGb.SelectAll();
                LkupPmntGb.Focus();
                return;
            }
            //#00002
            //else if (string.IsNullOrEmpty(sAText))
            //{
            //    XtraMessageBox.Show("적요를 입력하세요.");
            //    TxtAText.SelectAll();
            //    TxtAText.Focus();
            //    return;
            //}

            /*
             * COM_BASE_CD 테이블 -> CD_GB : PAY_METHOD
             * COM_CD -> 0001 : 송금
             * COM_CD -> 0002 : 현금
             * 송금일 경우 계좌정보 필수 입력
             * 현금일 경우 계정코드 0101 HARDCODING으로 이후 저장로직 처리
             */
            if (sPmntGb.Equals("0001"))
            {
                if (string.IsNullOrEmpty(sAcnt))
                {
                    XtraMessageBox.Show("계좌를 선택하세요!");
                    LkupAcnt.SelectAll();
                    LkupAcnt.Focus();
                    return;
                }
            }

            if (sPmntGb.Equals("0001"))
            {
                sAcCod2 = "0103"; //송금의 경우 계정코드 0251 HARDCODING -> 계정코드 변경 시 해당 코드도 변경하여야함
            }
            else if (sPmntGb.Equals("0002"))
            {
                sAcCod2 = "0101"; //현금의 경우 계정코드 0101 HARDCODING;
            }

            if (sPmntTarget.Equals("외상매입금"))
            {
                if (GridRetr.MainView != GridViewRetr)
                {
                    XtraMessageBox.Show("외상매입금에 해당하는 그리드가 아닙니다.\r\n관리자에게 문의하세요.");
                    return;
                }
            }
            else if (sPmntTarget.Equals("미지급금"))
            {
                if (GridRetr.MainView != GridViewRetr3)
                {
                    XtraMessageBox.Show("미지급금에 해당하는 그리드가 아닙니다.\r\n관리자에게 문의하세요.");
                    return;
                }
            }

            /*
             * COM_BASE_CD 테이블 -> CD_GB : AC16001_01
             * COM_CD -> 1 : 외상매입금
             * COM_CD -> 2 : 미지급금
             * COM_CD -> 3 : 외상매출금
             * COM_CD -> 4 : 미수금
             */
            if (sPmntTarget.Equals("외상매입금"))
            {
                sJGubn = "1";
                sAcCod1 = "0251";
            }
            else if (sPmntTarget.Equals("미지급금"))
            {
                sJGubn = "2";
                sAcCod1 = "0253";
            }

            //메소드 참조
            if (GridRetr.MainView == GridViewRetr)
            {
                if (ValidateData(GridViewRetr))
                    return;
            }
            else
            {
                if (ValidateData(GridViewRetr3))
                    return;
            }

            Dictionary<int, DataRow> dicResult;
            
            //메소드 참조
            if (GridRetr.MainView == GridViewRetr) 
                dicResult = ReturnData(GridViewRetr);
            else
                dicResult = ReturnData(GridViewRetr3);

            //#00002 적요 제거
            string sMsg = string.Format("-----적용사항-----" +
                        "\r\n결제대상 : {0}" +
                        "\r\n전표일자 : {1}" +
                        "\r\n지급방법 : {2}" +
                        "\r\n계좌 : {3}" +
                        "\r\n총 {4} 건에 대하여 전표를 발행하시겠습니까?"
                        , CboPmntTarget.EditValue?.ToString()
                        , DateEditTDate.EditValue?.ToString().Substring(0, 10)
                        , LkupPmntGb.Text
                        , LkupAcnt.Text
                        , dicResult.Keys.Count);

            if (XtraMessageBox.Show(sMsg, "발행확인여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;
            
            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                foreach (KeyValuePair<int, DataRow> row in dicResult)
                {
                    string sCvCod = row.Value["CVCOD"]?.ToString(); //거래처코드
                    string sOfsAmt = row.Value["OFS_AMT"]?.ToString();//지급금액
                    string sRmk = row.Value["RMK"]?.ToString(); //비고
                    //#00002
                    string sAText = row.Value["CVNAM"]?.ToString()+","+ TxtAText.EditValue?.ToString().Trim(); //적요 (요청)금액제거 2022-12-09

                    //PK 채번
                    strSql.Clear();
                    #region mariaDB
                    //strSql.AppendLine(" SELECT CONCAT('T', REPLACE(@TDATE, '-', ''), LPAD(IFNULL(SUBSTRING(MAX(A.SLIPNO), 10, 6), 0) + 1, 6, '0') ) AS PK_RST ");
                    //strSql.AppendLine("   FROM SUGMF A  ");
                    //strSql.AppendLine("  WHERE A.TDATE = @TDATE ");
                    #endregion

                    strSql.AppendLine("SELECT CONCAT('T', REPLACE(@TDATE, '-', ''),                                                    ");
				    strSql.AppendLine("                CAST(REPLICATE(0, 6 - LEN(ISNULL(SUBSTRING(MAX(A.SLIPNO), 10, 6), 0) + 1)) AS VARCHAR)");
                    strSql.AppendLine("                + CAST(ISNULL(SUBSTRING(MAX(A.SLIPNO), 10, 6), 0) + 1 AS VARCHAR)) AS PK_RST          ");
                    strSql.AppendLine("  FROM SUGMF A                                                                                        ");
                    strSql.AppendLine(" WHERE A.TDATE = @TDATE                                                                               ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@TDATE", sTDate);
                    string sSlipNo = cmd.ExecuteScalar()?.ToString();

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Add("@SLIPNO", sSlipNo);
                    dicParams.Add("@TDATE", sTDate);
                    dicParams.Add("@JGUBN", sJGubn);
                    dicParams.Add("@BILGU", sPmntGb);
                    dicParams.Add("@CVCOD", sCvCod);
                    dicParams.Add("@ACTNO", sAcnt);
                    dicParams.Add("@TRSUM", sOfsAmt);
                    dicParams.Add("@RK", sRmk);
                    dicParams.Add("@CUSER", FmMainToolBar2.UserID);

                    //수금/지급 테이블 INSERT
                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO SUGMF ");
                    strSql.AppendLine("           ( SLIPNO ");
                    strSql.AppendLine("           , TDATE ");
                    strSql.AppendLine("           , JGUBN ");
                    strSql.AppendLine("           , BILGU ");
                    strSql.AppendLine("           , CVCOD ");
                    strSql.AppendLine("           , ACTNO ");
                    strSql.AppendLine("           , TRSUM ");
                    strSql.AppendLine("           , RK ");
                    strSql.AppendLine("           , CDATE ");
                    strSql.AppendLine("           , CUSER ) ");
                    strSql.AppendLine("     VALUES( @SLIPNO ");
                    strSql.AppendLine("           , @TDATE ");
                    strSql.AppendLine("           , @JGUBN ");
                    strSql.AppendLine("           , @BILGU ");
                    strSql.AppendLine("           , @CVCOD ");
                    strSql.AppendLine("           , @ACTNO ");
                    strSql.AppendLine("           , @TRSUM ");
                    strSql.AppendLine("           , @RK ");
                    strSql.AppendLine("           , CONVERT(VARCHAR(20),GETDATE(),20) ");
                    strSql.AppendLine("           , @CUSER ); ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    foreach(KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                    }
                    cmd.ExecuteNonQuery();

                    #region mariaDB
                    //dicParams.Clear();
                    //dicParams.Add("@TDATE", sTDate);
                    //dicParams.Add("@ACCOD1", sAcCod1); //외상매입금, 미지급금에 대한 계정코드
                    //dicParams.Add("@ACCOD2", sAcCod2); //상대계정, 보통예금이나 현금에 대한 부분
                    //dicParams.Add("@CVCOD1", sCvCod); //외상매입금, 미지급금에 대한 거래처코드
                    //dicParams.Add("@CVCOD2", sAcnt); //상대계정, 보통예금이나 현금에 대한 거래처코드
                    //dicParams.Add("@ATEXT", sAText); //적요
                    //dicParams.Add("@REF1", sSlipNo); //SUGMF의 PK값
                    //dicParams.Add("@ACAMT", sOfsAmt); //대변 혼돈금지
                    //dicParams.Add("@ADAMT", sOfsAmt); //차변 혼돈금지
                    //dicParams.Add("@USRCD", FmMainToolBar2.UserID); //유저코드
                    //dicParams.Add("@RK", sRmk); //비고
                    #endregion

                    //ACTRAN 테이블 전표생성
                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO ACTRAN ");
                    strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO ");
                    strSql.AppendLine("           , ACCOD, ACNAM, CVCOD, CVNAM ");
                    strSql.AppendLine("           , ATEXT, ADAMT, ACAMT, ADPCD ");
                    strSql.AppendLine("           , ADPNM, REF1, APVYN, AAUTO ");
                    strSql.AppendLine("           , ADATE, AUSER, RK, CUSER ");
                    strSql.AppendLine("           , CDATE ) ");

                    #region mariaDB
                    //strSql.AppendLine("SELECT DATE_FORMAT(@TDATE, '%Y%m%d') AS TDATE  ");
                    //strSql.AppendLine("     , '3' AS ATGUB #대체전표 ");
                    //strSql.AppendLine("     , ( SELECT LPAD(CAST(IFNULL(MAX(Z1.SEQNO), 5000) AS DOUBLE) + 1, 4, '0') FROM ACTRAN Z1 WHERE Z1.TDATE = DATE_FORMAT(@TDATE, '%Y%m%d') AND Z1.ATGUB = '3' AND Z1.SEQNO >= '5000' ) AS SEQNO ");
                    //strSql.AppendLine("     , Y1.* ");
                    //strSql.AppendLine("  FROM ( ");
                    //strSql.AppendLine("         SELECT 0 AS LINNO #외상매입금 부분 ");
                    //strSql.AppendLine("              , CAST(@ACCOD1 AS CHAR) AS ACCOD ");
                    //strSql.AppendLine("              , ( SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = @ACCOD1 ) AS ACNAM ");
                    //strSql.AppendLine("              , CAST(@CVCOD1 AS DOUBLE) AS CVCOD ");
                    //strSql.AppendLine("              , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = @CVCOD1 ) AS CVNAM ");
                    //strSql.AppendLine("              , @ATEXT AS ATEXT ");
                    //strSql.AppendLine("              , 0 AS ADAMT #외상매입금, 미지급금 ( ADAMT는 차변임!!! ) ");
                    //strSql.AppendLine("              , @ACAMT AS ACAMT #외상매입금, 미지급금 ( ACAMT는 대변임!!! ) ");
                    //strSql.AppendLine("              , ( SELECT X2.DEPT_CD ");
                    //strSql.AppendLine("                    FROM ZUSRLST X1 ");
                    //strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2 ");
                    //strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID ");
                    //strSql.AppendLine("                   WHERE USRCD = @USRCD ) AS ADPCD ");
                    //strSql.AppendLine("              , ( SELECT X3.DEPT_NM ");
                    //strSql.AppendLine("                    FROM ZUSRLST X1 ");
                    //strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2 ");
                    //strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID ");
                    //strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3 ");
                    //strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD ");
                    //strSql.AppendLine("                   WHERE USRCD = @USRCD ) AS ADPNM ");
                    //strSql.AppendLine("              , @REF1 AS REF1 #SUGMF 테이블의 PK값 ");
                    //strSql.AppendLine("              , 'N' AS APVYN ");
                    //strSql.AppendLine("              , 'D01' AS AAUTO #자동전표 구분 HARDCODING ");
                    //strSql.AppendLine("              , NULL AS ADATE ");
                    //strSql.AppendLine("              , NULL AS AUSER ");
                    //strSql.AppendLine("              , @RK AS RK ");
                    //strSql.AppendLine("              , @USRCD AS CUSER ");
                    //strSql.AppendLine("              , NOW() AS CDATE ");
                    //strSql.AppendLine("          UNION ALL ");
                    //strSql.AppendLine("         SELECT 1 AS LINNO #상대계정 부분 ");
                    //strSql.AppendLine("              , CAST(@ACCOD2 AS CHAR) AS ACCOD ");
                    //strSql.AppendLine("              , ( SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = @ACCOD2 ) AS ACNAM ");
                    //strSql.AppendLine("              , CAST(@CVCOD2 AS DOUBLE) AS CVCOD ");
                    //strSql.AppendLine("              , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = @CVCOD2 ) AS CVNAM ");
                    //strSql.AppendLine("              , @ATEXT AS ATEXT ");
                    //strSql.AppendLine("              , @ADAMT AS ADAMT #상대계정 ( ADAMT는 차변임!!! ) ");
                    //strSql.AppendLine("              , 0 AS ACAMT #상대계정 ( ADAMT는 차변임!!! ) ");
                    //strSql.AppendLine("              , ( SELECT X2.DEPT_CD ");
                    //strSql.AppendLine("                    FROM ZUSRLST X1 ");
                    //strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2 ");
                    //strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID ");
                    //strSql.AppendLine("                   WHERE USRCD = @USRCD ) AS ADPCD ");
                    //strSql.AppendLine("              , ( SELECT X3.DEPT_NM ");
                    //strSql.AppendLine("                    FROM ZUSRLST X1 ");
                    //strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2 ");
                    //strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID ");
                    //strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3 ");
                    //strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD ");
                    //strSql.AppendLine("                   WHERE USRCD = @USRCD ) AS ADPNM ");
                    //strSql.AppendLine("              , @REF1 AS REF1 #SUGMF 테이블의 PK값 ");
                    //strSql.AppendLine("              , 'N' AS APVYN ");
                    //strSql.AppendLine("              , 'D01' AS AAUTO #자동전표 구분 ");
                    //strSql.AppendLine("              , NULL AS ADATE ");
                    //strSql.AppendLine("              , NULL AS AUSER ");
                    //strSql.AppendLine("              , @RK AS RK ");
                    //strSql.AppendLine("              , @USRCD AS CUSER ");
                    //strSql.AppendLine("              , NOW() AS CDATE ");
                    //strSql.AppendLine("       ) Y1 ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //foreach (KeyValuePair<string, string> param in dicParams)
                    //{
                    //    cmd.Parameters.AddWithValue(param.Key, param.Value);
                    //}
                    //cmd.ExecuteNonQuery();
                    #endregion

                    strSql.AppendLine("SELECT '" + sTDate.Replace("-","") + "' AS TDATE                                                                                          ");
                    strSql.AppendLine("     , '3' AS ATGUB --대체전표                                                                                                          ");
                    strSql.AppendLine("     , (SELECT CAST(REPLICATE(0, 4 - LEN(ISNULL(MAX(Z1.SEQNO), 5000) + 1)) AS VARCHAR) + CAST(ISNULL(MAX(Z1.SEQNO), 5000) + 1 AS VARCHAR)");
                    strSql.AppendLine("              FROM ACTRAN Z1                                                                                                             ");
                    strSql.AppendLine("             WHERE Z1.TDATE = CONVERT(VARCHAR(8), CONVERT(DATE,'" + sTDate.Replace("-", "") + "'), 112) AND Z1.ATGUB = '3' AND Z1.SEQNO >= '5000') AS SEQNO                 ");
                    strSql.AppendLine("      , Y1.*                                                                                                                            ");
                    strSql.AppendLine("   FROM(                                                                                                                                ");
                    strSql.AppendLine("          SELECT 0 AS LINNO--외상매입금 부분                                                                                            ");
                    strSql.AppendLine("               , '"+ sAcCod1 + "' AS ACCOD                                                                                                       ");
                    strSql.AppendLine("               , (SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = '" + sAcCod1 + "') AS ACNAM                                                     ");
                    strSql.AppendLine("             , "+ sCvCod + " AS CVCOD                                                                                                         ");
                    strSql.AppendLine("             , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = "+ sCvCod + " ) AS CVNAM                                      ");
                    strSql.AppendLine("             , '"+ sAText + "' AS ATEXT                                                                                                          ");
                    strSql.AppendLine("             , 0 AS ADAMT --외상매입금, 미지급금(ADAMT는 차변임!!! )                                                                    ");
                    strSql.AppendLine("              , "+ sOfsAmt + " AS ACAMT--외상매입금, 미지급금(ACAMT는 대변임!!! )                                                               ");
                    strSql.AppendLine("              , (SELECT X2.DEPT_CD                                                                                                      ");
                    strSql.AppendLine("                   FROM ZUSRLST X1                                                                                                      ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                                                          ");
                    strSql.AppendLine("                   WHERE USRCD = "+ FmMainToolBar2.UserID + " ) AS ADPCD                                                                                      ");
                    strSql.AppendLine("              , ( SELECT X3.DEPT_NM                                                                                                     ");
                    strSql.AppendLine("                    FROM ZUSRLST X1                                                                                                     ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                                                          ");
                    strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3                                                                                            ");
                    strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD                                                                                        ");
                    strSql.AppendLine("                   WHERE USRCD = " + FmMainToolBar2.UserID + " ) AS ADPNM                                                                                      ");
                    strSql.AppendLine("              , '"+ sSlipNo + "' AS REF1 --SUGMF 테이블의 PK값                                                                                     ");
                    strSql.AppendLine("              , 'N' AS APVYN                                                                                                            ");
                    strSql.AppendLine("              , 'D01' AS AAUTO --자동전표 구분 HARDCODING                                                                               ");
                    strSql.AppendLine("              , NULL AS ADATE                                                                                                           ");
                    strSql.AppendLine("              , NULL AS AUSER                                                                                                           ");
                    strSql.AppendLine("              , '"+ sRmk + "' AS RK                                                                                                               ");
                    strSql.AppendLine("              , '" + FmMainToolBar2.EmpID + "' AS CUSER                                                                                                         ");
                    strSql.AppendLine("              , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE                                                                                                      ");
                    strSql.AppendLine("          UNION ALL                                                                                                                     ");
                    strSql.AppendLine("         SELECT 1 AS LINNO --상대계정 부분                                                                                              ");
                    strSql.AppendLine("              , '" + sAcCod2 + "' AS ACCOD                                                                                                        ");
                    strSql.AppendLine("              , (SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = '" + sAcCod2 + "' ) AS ACNAM                                                     ");
                    strSql.AppendLine("             , "+ sAcnt + " AS CVCOD                                                                                                         ");
                    strSql.AppendLine("             , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = "+ sAcnt + " ) AS CVNAM                                      ");
                    strSql.AppendLine("             , '" + sAText + "' AS ATEXT                                                                                                          ");
                    strSql.AppendLine("             , "+ sOfsAmt + " AS ADAMT --상대계정(ADAMT는 차변임!!! )                                                                           ");
                    strSql.AppendLine("              , 0 AS ACAMT --상대계정(ADAMT는 차변임!!! )                                                                               ");
                    strSql.AppendLine("              , (SELECT X2.DEPT_CD                                                                                                      ");
                    strSql.AppendLine("                   FROM ZUSRLST X1                                                                                                      ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                                                          ");
                    strSql.AppendLine("                   WHERE USRCD = " + FmMainToolBar2.UserID + " ) AS ADPCD                                                                                      ");
                    strSql.AppendLine("              , ( SELECT X3.DEPT_NM                                                                                                     ");
                    strSql.AppendLine("                    FROM ZUSRLST X1                                                                                                     ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                                                          ");
                    strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3                                                                                            ");
                    strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD                                                                                        ");
                    strSql.AppendLine("                   WHERE USRCD = " + FmMainToolBar2.UserID + " ) AS ADPNM                                                                                      ");
                    strSql.AppendLine("              , '" + sSlipNo + "' AS REF1 --SUGMF 테이블의 PK값                                                                                     ");
                    strSql.AppendLine("              , 'N' AS APVYN                                                                                                            ");
                    strSql.AppendLine("              , 'D01' AS AAUTO --자동전표 구분                                                                                          ");
                    strSql.AppendLine("              , NULL AS ADATE                                                                                                           ");
                    strSql.AppendLine("              , NULL AS AUSER                                                                                                           ");
                    strSql.AppendLine("              , '" + sRmk + "' AS RK                                                                                                               ");
                    strSql.AppendLine("              , '" + FmMainToolBar2.EmpID + "' AS CUSER                                                                                                         ");
                    strSql.AppendLine("              , CONVERT(VARCHAR(20),GETDATE(),20) AS CDATE                                                                                                      ");
                    strSql.AppendLine("       ) Y1                                                                                                                             ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                
                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");
                BtnRetr.PerformClick();
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        ///     저장 전 그리드 내용 점검
        ///     1 : 그리드에 단 하나라도 지급금액이 설정되어 있지 않은 경우
        ///     2 : 그리드에 비고에는 내용이 존재하지만 지급금액이 설정되어 있지 않은 경우
        /// </summary>
        /// <param name="view"> 그리드 뷰 </param> 
        private bool ValidateData(GridView view)
        {
            bool bYn = false;
            double dTotVal = 0;
            for(int i = 0; i < view.RowCount; i++)
            {
                string sVal = view.GetRowCellValue(i, GridColOfsAmt)?.ToString();
                double dVal = string.IsNullOrEmpty(sVal) ? 0 : Convert.ToDouble(sVal);
                dTotVal += dVal;

                string sRmk = view.GetRowCellValue(i, GridColRmk)?.ToString().Trim();
                if(!string.IsNullOrEmpty(sRmk) && dVal == 0) //비고에는 값이 있지만 지급금액이 설정되어있지 않은 경우 break
                {
                    XtraMessageBox.Show("지급금액을 입력하세요.");
                    view.FocusedRowHandle = i;
                    view.FocusedColumn = GridColOfsAmt;
                    bYn = true;
                    break;
                }
            }

            if(dTotVal == 0)
            {
                bYn = true;
            }

            return bYn;
        }
        
        /// <summary>
        ///     저장을 위하여 값이 있는 DataRow 체크 후 해당 RowHandle 값과 DataRow 값 리턴
        /// </summary>
        /// <param name="view"> 그리드 뷰 </param> 
        /// 
        private Dictionary<int, DataRow> ReturnData(GridView view)
        {
            Dictionary<int, DataRow> dicReturn = new Dictionary<int, DataRow>();

            for(int i = 0; i < view.RowCount; i++)
            {
                string sVal = view.GetRowCellValue(i, GridColOfsAmt)?.ToString();
                double dVal = string.IsNullOrEmpty(sVal) ? 0 : Convert.ToDouble(sVal);

                if (dVal > 0)
                    dicReturn.Add(i, view.GetDataRow(i));
            }

            return dicReturn;
        }

        private void DropBtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "건별삭제(1개)")
            {
                DeleteOne();
            }
            else if (tag == "선택삭제(1개 이상)(구현 중..)")
            {
                XtraMessageBox.Show("구현 예정입니다.");
                return;
            }
        }

        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnOne;
        BarButtonItem BtnOverOne;
        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;

            popupMenu1 = new PopupMenu(barManager1);
            BtnOne = new BarButtonItem(barManager1, "건별삭제(1개)");
            BtnOverOne = new BarButtonItem(barManager1, "선택삭제(1개 이상)");
            popupMenu1.AddItem(BtnOne);
            popupMenu1.AddItem(BtnOverOne);

            DropBtnDelete.DropDownControl = popupMenu1;

            BtnOne.Tag = "건별삭제(1개)";
            BtnOne.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlipCopy_ItemClick);

            BtnOverOne.Tag = "선택삭제(1개 이상)";
            BtnOverOne.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlipMove_ItemClick);
        }

        private void BtnSlipCopy_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            DeleteOne();
        }

        private void BtnSlipMove_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            //DeleteOverOne();
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnDelete.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnDelete.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnDelete.Tag = submenuItem.Tag;
            DropBtnDelete.Text = string.Format("{0}", submenuItem.Tag);
        }

        private void DeleteOne()
        {
            if(GridViewRetr2.RowCount == 0)
            {
                return;
            }

            string sSlipNo = GridViewRetr2.GetFocusedRowCellValue(GridCol2SlipNo)?.ToString();
            string sTDate = GridViewRetr2.GetFocusedRowCellValue(GridCol2TDate)?.ToString();
            string sJGubn = GridViewRetr2.GetFocusedRowCellValue(GridCol2JGubn2)?.ToString();
            string sCvNam = GridViewRetr2.GetFocusedRowCellValue(GridCol2DealerNm)?.ToString();
            
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(CASE WHEN APVYN = 'Y' THEN 1 END) AS CNT  ");
            strSql.AppendLine("   FROM ACTRAN A ");
            strSql.AppendLine("  WHERE A.AAUTO = 'D01' ");
            strSql.AppendLine("    AND A.REF1 = '" + sSlipNo + "' ");
            strSql.AppendLine("  GROUP BY A.TDATE, A.ATGUB, A.SEQNO ");

            DataTable dtCnt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            int iCnt = 0;
            if(dtCnt.Rows.Count > 0)
            {
                int.TryParse(dtCnt.Rows[0]["CNT"]?.ToString(), out iCnt);
            }
            else
            {
                XtraMessageBox.Show("로직에러 : 승인카운트 조회불가 \r\n[담당자에게 문의하세요.]");
                return;
            }

            if(iCnt > 0)
            {
                XtraMessageBox.Show("해당 건은 전표승인 상태입니다.");
                return;
            }
            
            string sMsg = string.Format("지불번호 : {0}" +
                "\r\n발행일자 : {1}" +
                "\r\n결제대상 : {2}" +
                "\r\n매입처 : {3}" +
                "\r\n해당 건에 대하여 관련된 전표내역까지 삭제됩니다." +
                "\r\n그래도 진행하시겠습니까?"
                , sSlipNo
                , sTDate
                , sJGubn
                , sCvNam);

            if (XtraMessageBox.Show(sMsg, "삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;
                
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                //수금 SUGMF DELETE
                dicParams.Clear();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM SUGMF ");
                strSql.AppendLine("  WHERE SLIPNO = @SLIPNO ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@SLIPNO", sSlipNo);
                cmd.ExecuteNonQuery();
                
                /*
                 * 전표 DELETE
                 */
                dicParams.Clear();
                dicParams.Add("SLIPNO", sSlipNo);

                strSql.Clear();
                strSql.AppendLine(" SELECT TDATE ");
                strSql.AppendLine("      , ATGUB ");
                strSql.AppendLine("      , SEQNO ");
                strSql.AppendLine("   FROM ACTRAN ");
                strSql.AppendLine("  WHERE REF1 = @SLIPNO ");
                strSql.AppendLine("    AND AAUTO = 'D01' ");

                DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                if (dt.Rows.Count >= 1 && dt.Rows.Count <= 10)
                {
                    string sRstTDate = dt.Rows[0]["TDATE"]?.ToString();
                    string sAtGub = dt.Rows[0]["ATGUB"]?.ToString();
                    string sSeqNo = dt.Rows[0]["SEQNO"]?.ToString();

                    if (string.IsNullOrEmpty(sRstTDate) || string.IsNullOrEmpty(sAtGub) || string.IsNullOrEmpty(sSeqNo))
                    {
                        throw new Exception("업데이트 중 전표정보가 없습니다.\r\n관리자에게 문의하세요.");
                    }

                    dicParams.Clear();
                    dicParams.Add("@TDATE", sRstTDate);
                    dicParams.Add("@ATGUB", sAtGub);
                    dicParams.Add("@SEQNO", sSeqNo);

                    strSql.Clear();
                    strSql.AppendLine(" DELETE FROM ACTRAN ");
                    strSql.AppendLine("       WHERE TDATE = @TDATE ");
                    strSql.AppendLine("         AND ATGUB = @ATGUB ");
                    strSql.AppendLine("         AND SEQNO = @SEQNO ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    foreach (KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                    }
                    cmd.ExecuteNonQuery();
                }

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (TabControl.SelectedTabPage == TabPagePurc)
            {
                ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}_", this.Text, TabPagePurc.Text), GridRetr);
            }
            else if (TabControl.SelectedTabPage == TabPageSugm)
            {
                ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}_", this.Text, TabPageSugm.Text), GridRetr2);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC16001F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        /// <summary>
        ///     지급금액 및 비고 값에 따라 셀색 변경
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private void GridViewRetr_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if(e.Column == GridColOfsAmt)
            {
                double dVal = string.IsNullOrEmpty(e.CellValue?.ToString()) ? 0 : Convert.ToDouble(e.CellValue);
                if(dVal > 0)
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else
                {
                    e.Appearance.BackColor = SystemColors.Info;
                }
            }
            else if(e.Column == GridColRmk)
            {
                string sVal = e.CellValue?.ToString().Trim();
                if (!string.IsNullOrEmpty(sVal))
                    e.Appearance.BackColor = Color.PaleGreen;
                else
                    e.Appearance.BackColor = SystemColors.Info;
            }
        }

        #region[Query]


        /// <summary>
        ///     거래처별 미지급 현황 조회
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();
            DataTable dt = new DataTable();
            try
            {
                Cursor = Cursors.WaitCursor;
                /*
                 * #00001
                 * 속도 이슈로 인해 거래처 전체 출력 삭제
                 */

                strSql.Clear();
                strSql.AppendLine("WITH INFO AS(                                                                                                                                                      ");
                strSql.AppendLine("          SELECT X1.CVCOD                                                                                                                                          ");
                strSql.AppendLine("               , X1.ACCOD                                                                                                                                          ");
                strSql.AppendLine("               , SUM(X1.JAMT) AS JAMT                                                                                                                              ");
                strSql.AppendLine("            FROM(                                                                                                                                                  ");
                strSql.AppendLine("                   SELECT A1.ACCOD                                                                                                                                 ");
                strSql.AppendLine("                        , A1.CVCOD                                                                                                                                 ");
                strSql.AppendLine("                        , ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) AS JAMT                                                                                      ");
                strSql.AppendLine("                     FROM ACJANF A1                                                                                                                                ");
                strSql.AppendLine("                    WHERE A1.ACYEAR = SUBSTRING('" + dicParams["DATE_F"] + "', 1, 4)                                                                                                  ");
                strSql.AppendLine("                      AND A1.ACCOD = '" + dicParams["ACCOD"] + "'                                                                                                                        ");
                strSql.AppendLine("                    GROUP BY A1.CVCOD, A1.ACCOD, ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0)                                                                       ");
                strSql.AppendLine("                    UNION ALL                                                                                                                                      ");
                strSql.AppendLine("                   SELECT A1.ACCOD                                                                                                                                 ");
                strSql.AppendLine("                        , A1.CVCOD                                                                                                                                 ");
                strSql.AppendLine("                        , ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) END), 0) AS JAMT   ");
                strSql.AppendLine("                     FROM ACTRAN A1                                                                                                                                ");
                strSql.AppendLine("                     LEFT JOIN ACMSTF B1                                                                                                                           ");
                strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD                                                                                                                      ");
                strSql.AppendLine("                    WHERE A1.TDATE BETWEEN SUBSTRING('" + dicParams["DATE_F"] + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + dicParams["DATE_F"] + "')), 112)          ");
                strSql.AppendLine("                      AND A1.ACCOD = '" + dicParams["ACCOD"] + "'                                                                                                                        ");
                strSql.AppendLine("                      AND A1.APVYN = 'Y'                                                                                                                           ");
                strSql.AppendLine("                    GROUP BY A1.CVCOD, A1.ACCOD                                                                                                                    ");
                strSql.AppendLine("                 ) X1                                                                                                                                              ");
                strSql.AppendLine("           GROUP BY X1.CVCOD, X1.ACCOD                                                                                                                             ");
                strSql.AppendLine("), TEMP1 AS(                                                                                                                                                       ");
                strSql.AppendLine("    SELECT X1.DEALER_CD AS CVCOD                                                                                                                                   ");
                strSql.AppendLine("              , X1.DEALER_NM AS CVNAM                                                                                                                              ");
                strSql.AppendLine("              , X1.IDT_NO                                                                                                                                          ");
                strSql.AppendLine("              , X1.REP_NM                                                                                                                                          ");
                strSql.AppendLine("              , CASE WHEN X1.CHRG_TEL_NO IS NULL OR X1.CHRG_TEL_NO = '' THEN X1.CHRG_HP_NO ELSE X1.CHRG_TEL_NO END AS PHONE                                        ");
                strSql.AppendLine("              , ISNULL(MAX(X3.JAMT), 0) AS JAMT                                                                                                                    ");
                strSql.AppendLine("              , SUM(ISNULL(X2.ADAMT, 0)) AS OCR_AMT                                                                                                                ");
                strSql.AppendLine("              , SUM(ISNULL(X2.ACAMT, 0)) AS CPT_OFS_AMT                                                                                                            ");
                strSql.AppendLine("              , ISNULL(MAX(JAMT), 0) +SUM(ISNULL(X2.ADAMT, 0)) - SUM(ISNULL(X2.ACAMT, 0)) AS OFS_BLC                                                               ");
                strSql.AppendLine("             , 0 AS OFS_AMT                                                                                                                                        ");
                strSql.AppendLine("           FROM ACC_DEALER_CD X1                                                                                                                                   ");
                strSql.AppendLine("           LEFT JOIN ACTRAN X2                                                                                                                                     ");
                strSql.AppendLine("             ON X1.DEALER_CD = X2.CVCOD                                                                                                                            ");
                strSql.AppendLine("           LEFT OUTER JOIN INFO X3                                                                                                                                 ");
                strSql.AppendLine("             ON X2.CVCOD = X3.CVCOD                                                                                                                                ");
                strSql.AppendLine("          WHERE X2.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
                strSql.AppendLine("            AND X2.ACCOD = '" + dicParams["ACCOD"] + "' ");
                strSql.AppendLine("            AND X2.APVYN = 'Y' ");
                strSql.AppendLine("            AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
                strSql.AppendLine("                 OR  ");
                strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '0' AND (X1.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR X1.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) "); //거래처초성검색 추가(2021-02-07)
                strSql.AppendLine("                 OR  ");
                strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '1' AND X1.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
                strSql.AppendLine("                 OR  ");
                strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '2' AND X1.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
                strSql.AppendLine("          GROUP BY X1.DEALER_CD, X1.DEALER_NM, X1.IDT_NO, X1.REP_NM                                                                                                ");
                strSql.AppendLine("                 , CASE WHEN X1.CHRG_TEL_NO IS NULL OR X1.CHRG_TEL_NO = '' THEN X1.CHRG_HP_NO ELSE X1.CHRG_TEL_NO END                                              ");
                strSql.AppendLine(")                                                                                                                                                                  ");
                strSql.AppendLine(" SELECT Y1.CVCOD                                                                                                                                                   ");
                strSql.AppendLine("      , Y1.CVNAM                                                                                                                                                   ");
                strSql.AppendLine("      , Y1.IDT_NO                                                                                                                                                  ");
                strSql.AppendLine("      , Y1.REP_NM                                                                                                                                                  ");
                strSql.AppendLine("      , Y1.PHONE                                                                                                                                                   ");
                strSql.AppendLine("      , Y1.JAMT                                                                                                                                                    ");
                strSql.AppendLine("      , Y1.OCR_AMT                                                                                                                                                 ");
                strSql.AppendLine("      , Y1.CPT_OFS_AMT                                                                                                                                             ");
                strSql.AppendLine("      , Y1.OFS_AMT                                                                                                                                                 ");
                strSql.AppendLine("      , Y1.OFS_BLC                                                                                                                                                 ");
                strSql.AppendLine("      , '' AS RMK                                                                                                                                                  ");
                strSql.AppendLine("   FROM TEMP1 Y1                                                                                                                                                   ");
                strSql.AppendLine("   WHERE ((('" + dicParams["BAL_YN"] + "' = 'A' AND 1 = 1 ) ");
                strSql.AppendLine("          OR ");
                strSql.AppendLine("          ('" + dicParams["BAL_YN"] + "' = 'Y' AND Y1.OFS_BLC <> 0 ) ");
                strSql.AppendLine("          OR ");
                strSql.AppendLine("          ('" + dicParams["BAL_YN"] + "' = 'N' AND Y1.OFS_BLC = 0))) "); 
                strSql.AppendLine("   ORDER BY REPLACE(Y1.CVNAM, '(주)', '')                                                                                                                          ");
                                                                                                                                                                                                      
                #region mariaDB
                //strSql.AppendLine("   WITH INFO AS (           ");
                //strSql.AppendLine("          SELECT X1.CVCOD  ");
                //strSql.AppendLine("               , X1.ACCOD  ");
                //strSql.AppendLine("               , SUM(X1.JAMT) AS JAMT  ");
                //strSql.AppendLine("            FROM (  ");
                //strSql.AppendLine("                   SELECT A1.ACCOD   ");
                //strSql.AppendLine("                        , A1.CVCOD  ");
                //strSql.AppendLine("                        , IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) AS JAMT  ");
                //strSql.AppendLine("                     FROM ACJANF A1    ");
                //strSql.AppendLine("                    WHERE A1.ACYEAR = DATE_FORMAT('" + dicParams["DATE_F"] + "', '%Y')    ");
                //strSql.AppendLine("                      AND A1.ACCOD = '" + dicParams["ACCOD"] + "'  ");
                //strSql.AppendLine("                    GROUP BY A1.CVCOD  ");
                //strSql.AppendLine("                    UNION ALL  ");
                //strSql.AppendLine("                   SELECT A1.ACCOD  ");
                //strSql.AppendLine("                        , A1.CVCOD  ");
                //strSql.AppendLine("                        , IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT  ");
                //strSql.AppendLine("                     FROM ACTRAN A1   ");
                //strSql.AppendLine("                     LEFT JOIN ACMSTF B1    ");
                //strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD   ");
                //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN DATE_FORMAT('" + dicParams["DATE_F"] + "','%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + dicParams["DATE_F"] + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')   ");
                //strSql.AppendLine("                      AND A1.ACCOD = '" + dicParams["ACCOD"] + "'  ");
                //strSql.AppendLine("                      AND A1.APVYN = 'Y'  ");
                //strSql.AppendLine("                    GROUP BY A1.CVCOD  ");
                //strSql.AppendLine("                 ) X1  ");
                //strSql.AppendLine("           GROUP BY X1.CVCOD  ");
                //strSql.AppendLine("     )     ");
                //strSql.AppendLine(" SELECT Y1.CVCOD   ");
                //strSql.AppendLine("      , Y1.CVNAM   ");
                //strSql.AppendLine("      , Y1.IDT_NO   ");
                //strSql.AppendLine("      , Y1.REP_NM   ");
                //strSql.AppendLine("      , Y1.PHONE   ");
                //strSql.AppendLine("      , Y1.JAMT  ");
                //strSql.AppendLine("      , Y1.OCR_AMT   ");
                //strSql.AppendLine("      , Y1.CPT_OFS_AMT   ");
                //strSql.AppendLine("      , Y1.OFS_AMT   ");
                //strSql.AppendLine("      , Y1.OFS_BLC ");
                //strSql.AppendLine("      , '' AS RMK   ");
                //strSql.AppendLine("   FROM (   ");
                //strSql.AppendLine("         SELECT X1.DEALER_CD AS CVCOD   ");
                //strSql.AppendLine("              , X1.DEALER_NM AS CVNAM   ");
                //strSql.AppendLine("              , X1.IDT_NO   ");
                //strSql.AppendLine("              , X1.REP_NM   ");
                //strSql.AppendLine("              , CASE WHEN X1.CHRG_TEL_NO IS NULL OR X1.CHRG_TEL_NO = '' THEN X1.CHRG_HP_NO ELSE X1.CHRG_TEL_NO END AS PHONE   ");
                //strSql.AppendLine("              , IFNULL(X3.JAMT, 0) AS JAMT  ");
                //strSql.AppendLine("              , SUM(IFNULL(X2.ADAMT, 0)) AS OCR_AMT   ");
                //strSql.AppendLine("              , SUM(IFNULL(X2.ACAMT, 0)) AS CPT_OFS_AMT  ");
                //strSql.AppendLine("              , IFNULL(JAMT, 0) + SUM(IFNULL(X2.ADAMT, 0)) - SUM(IFNULL(X2.ACAMT, 0)) AS OFS_BLC ");
                //strSql.AppendLine("              , 0 AS OFS_AMT   ");
                //strSql.AppendLine("           FROM ACC_DEALER_CD X1 ");
                //strSql.AppendLine("           LEFT JOIN ACTRAN X2 ");
                //strSql.AppendLine("             ON X1.DEALER_CD = X2.CVCOD ");
                //strSql.AppendLine("           LEFT OUTER JOIN INFO X3 ");
                //strSql.AppendLine("             ON X2.CVCOD = X3.CVCOD ");
                //strSql.AppendLine("          WHERE X2.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
                //strSql.AppendLine("            AND X2.ACCOD = '" + dicParams["ACCOD"] + "' ");
                //strSql.AppendLine("            AND X2.APVYN = 'Y' ");
                //strSql.AppendLine("            AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
                //strSql.AppendLine("                 OR  ");
                //strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '0' AND (X1.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR X1.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) "); //거래처초성검색 추가(2021-02-07)
                //strSql.AppendLine("                 OR  ");
                //strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '1' AND X1.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
                //strSql.AppendLine("                 OR  ");
                //strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '2' AND X1.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
                //strSql.AppendLine("          GROUP BY X1.DEALER_CD ) Y1 ");
                //strSql.AppendLine("   WHERE ((('" + dicParams["BAL_YN"] + "' = 'A' AND 1 = 1 ) ");
                //strSql.AppendLine("          OR ");
                //strSql.AppendLine("          ('" + dicParams["BAL_YN"] + "' = 'Y' AND Y1.OFS_BLC <> 0 ) ");
                //strSql.AppendLine("          OR ");
                //strSql.AppendLine("          ('" + dicParams["BAL_YN"] + "' = 'N' AND Y1.OFS_BLC = 0))) ");
                //strSql.AppendLine("   ORDER BY REPLACE(Y1.CVNAM, '(주)', '') ");
                #endregion

                #region 2021-10-19 이전 로직
                /*
                 * 2021-01-14 현업요청
                 * 거래처별 이월금액이 존재하는데도 리스트에 나오지 않는 현상이 있어
                 * 거래처 전체로 출력
                 */
                //strSql.Clear();
                //strSql.AppendLine("   WITH INFO AS (           ");
                //strSql.AppendLine("          SELECT X1.CVCOD  ");
                //strSql.AppendLine("               , X1.ACCOD  ");
                //strSql.AppendLine("               , SUM(X1.JAMT) AS JAMT  ");
                //strSql.AppendLine("            FROM (  ");
                //strSql.AppendLine("                   SELECT A1.ACCOD   ");
                //strSql.AppendLine("                        , A1.CVCOD  ");
                //strSql.AppendLine("                        , IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) AS JAMT  ");
                //strSql.AppendLine("                     FROM ACJANF A1    ");
                //strSql.AppendLine("                    WHERE A1.ACYEAR = DATE_FORMAT('" + dicParams["DATE_F"] + "', '%Y')    ");
                //strSql.AppendLine("                      AND A1.ACCOD = '"+ dicParams["ACCOD"] + "'  ");
                //strSql.AppendLine("                    GROUP BY A1.CVCOD  ");
                //strSql.AppendLine("                    UNION ALL  ");
                //strSql.AppendLine("                   SELECT A1.ACCOD  ");
                //strSql.AppendLine("                        , A1.CVCOD  ");
                //strSql.AppendLine("                        , IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT  ");
                //strSql.AppendLine("                     FROM ACTRAN A1   ");
                //strSql.AppendLine("                     LEFT JOIN ACMSTF B1    ");
                //strSql.AppendLine("                       ON A1.ACCOD = B1.ACCOD   ");
                //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN DATE_FORMAT('" + dicParams["DATE_F"] + "','%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + dicParams["DATE_F"] + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')   ");
                //strSql.AppendLine("                      AND A1.ACCOD = '" + dicParams["ACCOD"] + "'  ");
                //strSql.AppendLine("                      AND A1.APVYN = 'Y'  ");
                //strSql.AppendLine("                    GROUP BY A1.CVCOD  ");
                //strSql.AppendLine("                 ) X1  ");
                //strSql.AppendLine("           GROUP BY X1.CVCOD  ");
                //strSql.AppendLine("     )     ");
                //strSql.AppendLine(" SELECT Y1.CVCOD   ");
                //strSql.AppendLine("      , Y1.CVNAM   ");
                //strSql.AppendLine("      , Y1.IDT_NO   ");
                //strSql.AppendLine("      , Y1.REP_NM   ");
                //strSql.AppendLine("      , Y1.PHONE   ");
                //strSql.AppendLine("      , Y1.JAMT  ");
                //strSql.AppendLine("      , Y1.OCR_AMT   ");
                //strSql.AppendLine("      , Y1.CPT_OFS_AMT   ");
                //strSql.AppendLine("      , Y1.OFS_AMT   ");
                //strSql.AppendLine("      , Y1.OFS_BLC ");
                //strSql.AppendLine("      , '' AS RMK   ");
                //strSql.AppendLine("   FROM (   ");
                //strSql.AppendLine("         SELECT X1.DEALER_CD AS CVCOD   ");
                //strSql.AppendLine("              , X1.DEALER_NM AS CVNAM   ");
                //strSql.AppendLine("              , X1.IDT_NO   ");
                //strSql.AppendLine("              , X1.REP_NM   ");
                //strSql.AppendLine("              , CASE WHEN X1.CHRG_TEL_NO IS NULL OR X1.CHRG_TEL_NO = '' THEN X1.CHRG_HP_NO ELSE X1.CHRG_TEL_NO END AS PHONE   ");
                //strSql.AppendLine("              , IFNULL(X3.JAMT, 0) AS JAMT  ");
                //strSql.AppendLine("              , SUM(IFNULL(X2.ADAMT, 0)) AS OCR_AMT   ");
                //strSql.AppendLine("              , SUM(IFNULL(X2.ACAMT, 0)) AS CPT_OFS_AMT  ");
                //strSql.AppendLine("              , IFNULL(JAMT, 0) + SUM(IFNULL(X2.ADAMT, 0)) - SUM(IFNULL(X2.ACAMT, 0)) AS OFS_BLC ");
                //strSql.AppendLine("              , 0 AS OFS_AMT   ");
                //strSql.AppendLine("           FROM ACC_DEALER_CD X1 ");
                //strSql.AppendLine("           LEFT JOIN ACTRAN X2 ");
                //strSql.AppendLine("             ON X1.DEALER_CD = X2.CVCOD ");
                //strSql.AppendLine("           LEFT OUTER JOIN INFO X3 ");
                //strSql.AppendLine("             ON X2.CVCOD = X3.CVCOD ");
                //strSql.AppendLine("          WHERE X2.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
                //strSql.AppendLine("            AND X2.ACCOD = '" + dicParams["ACCOD"] + "' ");
                //strSql.AppendLine("            AND X2.APVYN = 'Y' ");
                //strSql.AppendLine("            AND (('"+ dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
                //strSql.AppendLine("                 OR  ");
                //strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '0' AND (X1.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR X1.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) "); //거래처초성검색 추가(2021-02-07)
                //strSql.AppendLine("                 OR  ");
                //strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '1' AND X1.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
                //strSql.AppendLine("                 OR  ");
                //strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '2' AND X1.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
                //strSql.AppendLine("          GROUP BY X1.DEALER_CD ");
                //strSql.AppendLine("          UNION ALL ");
                //strSql.AppendLine("          SELECT X1.DEALER_CD AS CVCOD   ");
                //strSql.AppendLine("               , X1.DEALER_NM AS CVNAM   ");
                //strSql.AppendLine("               , X1.IDT_NO   ");
                //strSql.AppendLine("               , X1.REP_NM   ");
                //strSql.AppendLine("               , CASE WHEN X1.CHRG_TEL_NO IS NULL OR X1.CHRG_TEL_NO = '' THEN X1.CHRG_HP_NO ELSE X1.CHRG_TEL_NO END AS PHONE   ");
                //strSql.AppendLine("               , IFNULL(X3.JAMT, 0) AS JAMT  ");
                //strSql.AppendLine("               , 0 AS OCR_AMT   ");
                //strSql.AppendLine("               , 0 AS CPT_OFS_AMT   ");
                //strSql.AppendLine("               , IFNULL(X3.JAMT, 0) AS OFS_BLC ");
                //strSql.AppendLine("               , 0 AS OFS_AMT   ");
                //strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                //strSql.AppendLine("            LEFT JOIN ACTRAN X2 ");
                //strSql.AppendLine("              ON X1.DEALER_CD = X2.CVCOD ");
                //strSql.AppendLine("            LEFT OUTER JOIN INFO X3 ");
                //strSql.AppendLine("              ON X1.DEALER_CD = X3.CVCOD ");
                //strSql.AppendLine("           WHERE X2.TDATE IS NULL ");
                //strSql.AppendLine("             AND X1.EOB_YN = 'N' ");
                //strSql.AppendLine("             AND X1.DEALER_NM <> '' ");
                //strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1) ");
                //strSql.AppendLine("                  OR  ");
                //strSql.AppendLine("                  ('"+ dicParams["FIND_IDX"] + "' = '0' AND (X1.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR X1.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
                //strSql.AppendLine("                  OR  ");
                //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '1' AND X1.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%') ");
                //strSql.AppendLine("                  OR  ");
                //strSql.AppendLine("                  ('" + dicParams["FIND_IDX"] + "' = '2' AND X1.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%')) ");
                //strSql.AppendLine("           GROUP BY X1.DEALER_CD ) Y1 ");
                //strSql.AppendLine("   WHERE ((('" + dicParams["BAL_YN"] + "' = 'A' AND 1 = 1 ) ");
                //strSql.AppendLine("          OR ");
                //strSql.AppendLine("          ('" + dicParams["BAL_YN"] + "' = 'Y' AND Y1.OFS_BLC <> 0 ) ");
                //strSql.AppendLine("          OR ");
                //strSql.AppendLine("          ('" + dicParams["BAL_YN"] + "' = 'N' AND Y1.OFS_BLC = 0))) ");
                //strSql.AppendLine("   ORDER BY REPLACE(Y1.CVNAM, '(주)', '') ");

                #endregion

                #region[2021-01-14 수정 전 쿼리]

                //strSql.Clear();
                //strSql.AppendLine("    WITH INFO AS (          ");
                //strSql.AppendLine("         SELECT X1.CVCOD ");
                //strSql.AppendLine("              , X1.ACCOD ");
                //strSql.AppendLine("              , SUM(X1.JAMT) AS JAMT ");
                //strSql.AppendLine("           FROM ( ");
                //strSql.AppendLine("                  SELECT A1.ACCOD  ");
                //strSql.AppendLine("                       , A1.CVCOD ");
                //strSql.AppendLine("                       , IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) AS JAMT ");
                //strSql.AppendLine("                    FROM ACJANF A1   ");
                //strSql.AppendLine("                   WHERE A1.ACYEAR = DATE_FORMAT(@DATE_F, '%Y')   ");
                //strSql.AppendLine("                     AND A1.ACCOD = @ACCOD   ");
                //strSql.AppendLine("                   GROUP BY A1.CVCOD ");
                //strSql.AppendLine("                   UNION ALL ");
                //strSql.AppendLine("                  SELECT A1.ACCOD  ");
                //strSql.AppendLine("                       , A1.CVCOD ");
                //strSql.AppendLine("                       , IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN IFNULL(ACAMT, 0) - IFNULL(ADAMT, 0) ELSE IFNULL(ADAMT, 0) - IFNULL(ACAMT, 0) END),0) AS JAMT ");
                //strSql.AppendLine("                    FROM ACTRAN A1  ");
                //strSql.AppendLine("                    LEFT JOIN ACMSTF B1   ");
                //strSql.AppendLine("                      ON A1.ACCOD = B1.ACCOD  ");
                //strSql.AppendLine("                   WHERE A1.TDATE BETWEEN DATE_FORMAT(@DATE_F,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT(@DATE_F, '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d')  ");
                //strSql.AppendLine("                     AND A1.ACCOD = @ACCOD    ");
                //strSql.AppendLine("                     AND A1.APVYN = 'Y' ");
                //strSql.AppendLine("                   GROUP BY A1.CVCOD ");
                //strSql.AppendLine("                ) X1 ");
                //strSql.AppendLine("          GROUP BY X1.CVCOD ");
                //strSql.AppendLine("    )    ");
                //strSql.AppendLine(" SELECT Y1.CVCOD  ");
                //strSql.AppendLine("      , Y1.CVNAM  ");
                //strSql.AppendLine("      , Y1.IDT_NO  ");
                //strSql.AppendLine("      , Y1.REP_NM  ");
                //strSql.AppendLine("      , Y1.PHONE  ");
                //strSql.AppendLine("      , Y1.JAMT ");
                //strSql.AppendLine("      , Y1.OCR_AMT  ");
                //strSql.AppendLine("      , Y1.CPT_OFS_AMT  ");
                //strSql.AppendLine("      , Y1.OFS_AMT  ");
                //strSql.AppendLine("      , Y1.JAMT + Y1.OCR_AMT - Y1.CPT_OFS_AMT AS OFS_BLC  ");
                //strSql.AppendLine("      , '' AS RMK  ");
                //strSql.AppendLine("   FROM (  ");
                //strSql.AppendLine("          SELECT X1.CVCOD  ");
                //strSql.AppendLine("               , X1.CVNAM  ");
                //strSql.AppendLine("               , X2.IDT_NO  ");
                //strSql.AppendLine("               , X2.REP_NM  ");
                //strSql.AppendLine("               , CASE WHEN X2.CHRG_TEL_NO IS NULL OR X2.CHRG_TEL_NO = '' THEN X2.CHRG_HP_NO ELSE X2.CHRG_TEL_NO END AS PHONE  ");
                //strSql.AppendLine("               , IFNULL(X3.JAMT, 0) AS JAMT ");
                //strSql.AppendLine("               , SUM(IFNULL(X1.ADAMT, 0)) AS OCR_AMT  ");
                //strSql.AppendLine("               , SUM(IFNULL(X1.ACAMT, 0)) AS CPT_OFS_AMT  ");
                //strSql.AppendLine("               , 0 AS OFS_AMT  ");
                //strSql.AppendLine("            FROM ACTRAN X1  ");
                //strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD X2  ");
                //strSql.AppendLine("              ON X1.CVCOD = X2.DEALER_CD  ");
                //strSql.AppendLine("            LEFT OUTER JOIN INFO X3  ");
                //strSql.AppendLine("              ON X1.CVCOD = X3.CVCOD ");
                //strSql.AppendLine("           WHERE X1.TDATE BETWEEN @DATE_F AND @DATE_T ");
                //strSql.AppendLine("             AND X1.ACCOD = @ACCOD #HARDCODING  ");
                //strSql.AppendLine("             AND X1.APVYN = 'Y'  ");
                //strSql.AppendLine("             AND ((@FIND_WORD = '' AND 1 = 1)    ");
                //strSql.AppendLine("                 OR");
                //strSql.AppendLine("                 ( @FIND_IDX = '0' AND X2.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' )");
                //strSql.AppendLine("                 OR");
                //strSql.AppendLine("                 ( @FIND_IDX = '1' AND X2.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )");
                //strSql.AppendLine("                 OR");
                //strSql.AppendLine("                 ( @FIND_IDX = '2' AND X2.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))");
                //strSql.AppendLine("           GROUP BY X1.CVCOD ");
                //strSql.AppendLine("        ) Y1 ");
                //strSql.AppendLine("  WHERE ((@BAL_YN = 'A' AND 1 = 1 )");
                //strSql.AppendLine("         OR");
                //strSql.AppendLine("         ( @BAL_YN = 'Y' AND (Y1.JAMT + Y1.OCR_AMT - Y1.CPT_OFS_AMT) <> 0 )");
                //strSql.AppendLine("         OR");
                //strSql.AppendLine("         ( @BAL_YN = 'N' AND (Y1.JAMT + Y1.OCR_AMT - Y1.CPT_OFS_AMT) = 0))");
                //strSql.AppendLine("  ORDER BY REPLACE(Y1.CVNAM, '(주)', '')");

                #endregion[2021-01-14 수정 전 쿼리]

                dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {

            }

            return dt;
        }

        /// <summary>
        ///     거래처별 지불 현황 조회
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetPaymentInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();
                                                                                                                                                                               
            strSql.AppendLine("WITH ACTRAN_INFO AS(                                                                                                                            ");
            strSql.AppendLine("        SELECT X1.TDATE, X1.REF1, X1.APVYN, MAX(X1.ATEXT) AS ATEXT");//#00002
            strSql.AppendLine("        FROM ACTRAN X1                                                                                                                          ");
            strSql.AppendLine("        WHERE X1.TDATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATE, '" + dicParams["DATE_F"] + "'), 112) AND CONVERT(VARCHAR(8), CONVERT(DATE, '" + dicParams["DATE_T"] + "'), 112)  ");
            strSql.AppendLine("        AND X1.AAUTO = 'D01'--자동전표 구분->지불                                                                                               ");
            strSql.AppendLine("        GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO, X1.REF1, X1.APVYN                                                                                ");
            strSql.AppendLine(")                                                                                                                                               ");
            strSql.AppendLine("SELECT A.SLIPNO                                                                                                                                 ");
            strSql.AppendLine("     , A.TDATE                                                                                                                                  ");
            strSql.AppendLine("     , A.JGUBN                                                                                                                                  ");
            strSql.AppendLine("     , B.COM_NM AS JGUBN2                                                                                                                       ");
            strSql.AppendLine("     , A.BILGU                                                                                                                                  ");
            strSql.AppendLine("     , C.COM_NM AS BILGU2                                                                                                                       ");
            strSql.AppendLine("     , A.CVCOD                                                                                                                                  ");
            strSql.AppendLine("     , D.DEALER_NM                                                                                                                              ");
            strSql.AppendLine("     , D.IDT_NO                                                                                                                                 ");
            strSql.AppendLine("     , D.REP_NM                                                                                                                                 ");
            strSql.AppendLine("     , CASE WHEN D.CHRG_TEL_NO IS NULL OR D.CHRG_TEL_NO = '' THEN D.CHRG_HP_NO ELSE D.CHRG_TEL_NO END AS PHONE                                  ");
            strSql.AppendLine("     , A.ACTNO                                                                                                                                  ");
            strSql.AppendLine("     , E.DEALER_NM AS ACNT                                                                                                                      ");
            strSql.AppendLine("     , A.TRSUM                                                                                                                                  ");
            strSql.AppendLine("     , CASE WHEN G.TDATE IS NULL THEN '미발행' ELSE '발행' END AS ISSUE_YN                                                                      ");
            strSql.AppendLine("     , CASE WHEN G.APVYN = 'Y' THEN '승인' ELSE '미승인' END AS APVYN                                                                           ");
            strSql.AppendLine("     , A.RK                                                                                                                                     ");
            strSql.AppendLine("     , G.ATEXT");//#00002
            strSql.AppendLine("     , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER                                     ");
            strSql.AppendLine("     , A.CDATE                                                                                                                                  ");
            strSql.AppendLine("     , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER                                     ");
            strSql.AppendLine("     , A.MDATE                                                                                                                                  ");
            strSql.AppendLine("  FROM SUGMF A                                                                                                                                  ");
            strSql.AppendLine("  LEFT JOIN COM_BASE_CD B                                                                                                                       ");
            strSql.AppendLine("    ON A.JGUBN = B.COM_CD                                                                                                                       ");
            strSql.AppendLine("   AND B.CD_GB = 'AC16001_01'                                                                                                                   ");
            strSql.AppendLine("  LEFT JOIN COM_BASE_CD C                                                                                                                       ");
            strSql.AppendLine("    ON A.BILGU = C.COM_CD                                                                                                                       ");
            strSql.AppendLine("   AND C.CD_GB = 'PAY_METHOD'                                                                                                                   ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD D                                                                                                                     ");
            strSql.AppendLine("    ON A.CVCOD = D.DEALER_CD                                                                                                                    ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD E                                                                                                                     ");
            strSql.AppendLine("    ON A.ACTNO = E.DEALER_CD                                                                                                                    ");
            strSql.AppendLine("  LEFT JOIN ACTRAN_INFO G                                                                                                                       ");
            strSql.AppendLine("    ON A.SLIPNO = G.REF1                                                                                                                        ");
            strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("    AND A.JGUBN IN ('1', '2') ");
            strSql.AppendLine("    AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1)    ");
            strSql.AppendLine("        OR");
            strSql.AppendLine("        ( '" + dicParams["FIND_IDX"] + "' = '0' AND (D.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR D.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))");
            strSql.AppendLine("        OR");
            strSql.AppendLine("        ( '" + dicParams["FIND_IDX"] + "' = '1' AND D.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )");
            strSql.AppendLine("        OR");
            strSql.AppendLine("        ( '" + dicParams["FIND_IDX"] + "' = '2' AND D.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))");

            #region mariaDB
            //strSql.AppendLine(" WITH ACTRAN_INFO AS ( ");
            //strSql.AppendLine(" 	    SELECT X1.TDATE, X1.REF1, X1.APVYN ");
            //strSql.AppendLine(" 	      FROM ACTRAN X1 ");
            //strSql.AppendLine(" 	     WHERE X1.TDATE BETWEEN DATE_FORMAT(@DATE_F, '%Y%m%d') AND DATE_FORMAT(@DATE_T, '%Y%m%d') ");
            //strSql.AppendLine(" 	       AND X1.AAUTO = 'D01' #자동전표 구분 -> 지불 ");
            //strSql.AppendLine(" 	     GROUP BY X1.TDATE, X1.ATGUB, X1.SEQNO ");
            //strSql.AppendLine(" )    ");
            //strSql.AppendLine("SELECT A.SLIPNO ");
            //strSql.AppendLine("     , A.TDATE ");
            //strSql.AppendLine("     , A.JGUBN ");
            //strSql.AppendLine("     , B.COM_NM AS JGUBN2 ");
            //strSql.AppendLine("     , A.BILGU ");
            //strSql.AppendLine("     , C.COM_NM AS BILGU2 ");
            //strSql.AppendLine("     , A.CVCOD ");
            //strSql.AppendLine("     , D.DEALER_NM ");
            //strSql.AppendLine("     , D.IDT_NO ");
            //strSql.AppendLine("     , D.REP_NM ");
            //strSql.AppendLine("     , CASE WHEN D.CHRG_TEL_NO IS NULL OR D.CHRG_TEL_NO = '' THEN D.CHRG_HP_NO ELSE D.CHRG_TEL_NO END AS PHONE ");
            //strSql.AppendLine("     , A.ACTNO ");
            //strSql.AppendLine("     , E.DEALER_NM AS ACNT ");
            //strSql.AppendLine("     , A.TRSUM ");
            //strSql.AppendLine("     , CASE WHEN G.TDATE IS NULL THEN '미발행' ELSE '발행' END AS ISSUE_YN ");
            //strSql.AppendLine("     , CASE WHEN G.APVYN = 'Y' THEN '승인' ELSE '미승인' END AS APVYN ");
            //strSql.AppendLine("     , A.RK ");
            //strSql.AppendLine("     , F1.USRNM AS CUSER  ");
            //strSql.AppendLine("     , A.CDATE  ");
            //strSql.AppendLine("     , F2.USRNM AS MUSER ");
            //strSql.AppendLine("     , A.MDATE ");
            //strSql.AppendLine("  FROM SUGMF A ");
            //strSql.AppendLine("  LEFT JOIN COM_BASE_CD B ");
            //strSql.AppendLine("    ON A.JGUBN = B.COM_CD ");
            //strSql.AppendLine("   AND B.CD_GB = 'AC16001_01' ");
            //strSql.AppendLine("  LEFT JOIN COM_BASE_CD C  ");
            //strSql.AppendLine("    ON A.BILGU = C.COM_CD ");
            //strSql.AppendLine("   AND C.CD_GB = 'PAY_METHOD' ");
            //strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD D  ");
            //strSql.AppendLine("    ON A.CVCOD = D.DEALER_CD ");
            //strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD E  ");
            //strSql.AppendLine("    ON A.ACTNO = E.DEALER_CD ");
            //strSql.AppendLine("  LEFT JOIN ZUSRLST F1 ");
            //strSql.AppendLine("    ON A.CUSER = F1.USRCD ");
            //strSql.AppendLine("  LEFT JOIN ZUSRLST F2 ");
            //strSql.AppendLine("    ON A.MUSER = F2.USRCD ");
            //strSql.AppendLine("  LEFT JOIN ACTRAN_INFO G  ");
            //strSql.AppendLine("    ON A.SLIPNO = G.REF1 ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T ");
            //strSql.AppendLine("    AND A.JGUBN IN ('1', '2') ");
            //strSql.AppendLine("    AND ((@FIND_WORD = '' AND 1 = 1)    ");
            //strSql.AppendLine("        OR");
            //strSql.AppendLine("        ( @FIND_IDX = '0' AND (D.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR D.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))");
            //strSql.AppendLine("        OR");
            //strSql.AppendLine("        ( @FIND_IDX = '1' AND D.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )");
            //strSql.AppendLine("        OR");
            //strSql.AppendLine("        ( @FIND_IDX = '2' AND D.REP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))");
            #endregion

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        #endregion[Query]

        //초기 LookupEdit 세팅
        private void SetLookup()
        {
            DataTable dtPmntGb = GetLookUpData("1", "Y", ""); //지급방법
            LkupPmntGb.Properties.ValueMember = "CD";
            LkupPmntGb.Properties.DisplayMember = "NM";
            LkupPmntGb.Properties.DataSource = dtPmntGb;

            DataTable dtAcnt = GetLookUpData("2", "Y", ""); //계좌정보
            LkupAcnt.Properties.ValueMember = "CD";
            LkupAcnt.Properties.DisplayMember = "NM";
            LkupAcnt.Properties.DataSource = dtAcnt;
        }

        /// <summary>
        ///     Lookup 세팅을 위한 DataTable Return Method
        /// </summary>
        /// /// <param name="sGb"> Index 값 </param> 
        /// /// <param name="sNullYn"> 그냥 Y </param> 
        /// /// <param name="sParam"> 그냥 "" 처리 </param> 
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '선택' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine("  WHERE CD_GB = 'PAY_METHOD' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT CAST(DEALER_CD AS VARCHAR) AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A ");
                strSql.AppendLine("  WHERE A.BANKYN = 'Y' "); //계좌는 거래처코드에서 관리하여 9000~10000 사이로 관리 중 -> #00002 은행구분YN 으로 변경 
                strSql.AppendLine("    AND A.EOB_YN = 'N' ");
            }
            //else if (sGb.Equals("3"))
            //{
            //    strSql.AppendLine(" SELECT BANK_ACNT_NO AS CD ");
            //    strSql.AppendLine("      , BANK_ACNT_NO AS NM ");
            //    strSql.AppendLine("   FROM ACC_ACNT_CD ");
            //    strSql.AppendLine("  WHERE TRMN_YN = 'N' ");
            //    strSql.AppendLine("    AND RPLC_PMNT_YN = 'Y' ");
            //}
            //else if (sGb.Equals("4"))
            //{
            //    strSql.AppendLine(" SELECT PRNOTE_NO AS CD ");
            //    strSql.AppendLine("      , PRNOTE_NO AS NM ");
            //    strSql.AppendLine("   FROM ACC_PRNOTE_MGT ");
            //    strSql.AppendLine("  WHERE PRNOTE_GB = '2' ");
            //    strSql.AppendLine("    AND DISCNT_FIN_YN = 'N' ");

            //}
            //else if (sGb.Equals("5"))
            //{
            //    strSql.AppendLine(" SELECT ACC_CD AS CD  ");
            //    strSql.AppendLine("      , ACC_NM AS NM ");
            //    strSql.AppendLine("   FROM ACC_ACC_CD ");
            //    strSql.AppendLine("  WHERE PYBC_YN = 'Y' ");
            //    strSql.AppendLine("    AND DBCR_GB = 'C' ");
            //    strSql.AppendLine("    AND USE_YN = 'Y' ");
            //    strSql.AppendLine("    AND CRT_YN = 'Y' ");
            //}

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY CASE WHEN CD = '' THEN '100' ELSE CD END");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        /// <summary>
        ///     라디오버튼 잔액여부 바뀔때마다 Select처리
        /// </summary>
        private void RdgbBalYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        /// <summary>
        ///     지급금액 0 이하 체크
        /// </summary>
        private void RepoTxtNumEdit_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridColOfsAmt, txt.EditValue);
            double dVal = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColOfsAmt)?.ToString().Trim()) ? 0 : Convert.ToDouble(txt.EditValue);

            if (dVal <= 0)
            {
                GridViewRetr.SetFocusedRowCellValue(GridColOfsAmt, 0);
            }
        }

        /// <summary>
        ///     지급금액 0 이하일 경우 비고 
        /// </summary>
        private void RepoTxtNumEdit_Leave(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
        }

        /// <summary>
        ///     1. 지급구분 0001 (송금->보통예금)의 경우 계좌정보 보이고 아니면 끔
        ///     2. 값에 따라 색변경
        /// </summary>
        private void LkupPmntGb_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEdit lkup = (LookUpEdit)sender;
            string sVal = string.IsNullOrEmpty(lkup.EditValue?.ToString()) ? string.Empty : lkup.EditValue?.ToString();
            if (sVal.Equals("0001"))
            {
                LayoutAcnt.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else
            {
                LayoutAcnt.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }

            if (string.IsNullOrEmpty(sVal))
            {
                lkup.Properties.Appearance.BackColor = Color.OrangeRed;
            }
            else if (!string.IsNullOrEmpty(sVal))
            {
                lkup.Properties.Appearance.BackColor = Color.PaleGreen;
            }
        }

        private void GridViewRetr_FocusedColumnChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedColumnChangedEventArgs e)
        {
            //if(e.FocusedColumn == GridColOfsAmt)
            //{
            //    double dVal = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(e.FocusedColumn)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(e.FocusedColumn));
            //    if(dVal == 0)
            //    {
            //        GridViewRetr.SetFocusedRowCellValue(e.FocusedColumn, GridViewRetr.GetFocusedRowCellValue(GridColOfsBlc));
            //    }
            //}
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                double dVal = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColOfsAmt)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColOfsAmt));
                double dOfsBls = Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColOfsBlc));
                if (dVal == 0)
                {
                    GridViewRetr.SetFocusedRowCellValue(GridColOfsAmt, GridViewRetr.GetFocusedRowCellValue(GridColOfsBlc));
                }
                else if(dVal == dOfsBls)
                {
                    GridViewRetr.SetFocusedRowCellValue(GridColOfsAmt, 0);
                }
            }
        }

        private void GridViewRetr_ShownEditor(object sender, EventArgs e)
        {
            
        }

        private void GridViewRetr2_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if(e.Column == GridCol2IssueYn)
            {
                string sVal = string.IsNullOrEmpty(e.CellValue?.ToString()) ? string.Empty : e.CellValue?.ToString();
                if (sVal.Equals("발행"))
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else if (sVal.Equals("미발행"))
                {
                    e.Appearance.BackColor = Color.OrangeRed;
                }
            }
            else if(e.Column == GridColApvYn)
            {
                string sVal = string.IsNullOrEmpty(e.CellValue?.ToString()) ? string.Empty : e.CellValue?.ToString();
                if (sVal.Equals("승인"))
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else if (sVal.Equals("미승인"))
                {
                    e.Appearance.BackColor = Color.OrangeRed;
                }
            }
        }

        /// <summary>
        ///     탭페이지에 따라 검색조건 및 버튼 visible 다르게 세팅
        /// </summary>
        private void TabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if(e.Page == TabPagePurc)
            {
                LayoutRdgbBlcYn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                LayoutBtnAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnSave.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                /*
                 * 2021-01-06 현업요청
                 * 지불관리 삭제
                 */
                //LayoutBtnDelete.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else if(e.Page == TabPageSugm)
            {
                LayoutRdgbBlcYn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                LayoutBtnAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutBtnSave.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //LayoutBtnDelete.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }

            BtnRetr.PerformClick();
        }

        /// <summary>
        ///        값에 따라 색변경
        /// </summary>
        private void CboPmntTarget_EditValueChanged(object sender, EventArgs e)
        {
            ComboBoxEdit cbo = (ComboBoxEdit)sender;
            string sVal = cbo.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
            {
                cbo.Properties.Appearance.BackColor = Color.OrangeRed;
            }
            else if (!string.IsNullOrEmpty(sVal))
            {
                cbo.Properties.Appearance.BackColor = Color.PaleGreen;
            }
        }

        /// <summary>
        ///        값에 따라 색변경
        /// </summary>
        private void DateEditTDate_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit date = (DateEdit)sender;
            string sVal = date.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
            {
                date.Properties.Appearance.BackColor = Color.OrangeRed;
            }
            else if (!string.IsNullOrEmpty(sVal))
            {
                date.Properties.Appearance.BackColor = Color.PaleGreen;
            }
        }

        /// <summary>
        ///        값에 따라 색변경
        /// </summary>
        private void LkupAcnt_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEdit lkup = (LookUpEdit)sender;
            string sVal = string.IsNullOrEmpty(lkup.EditValue?.ToString()) ? string.Empty : lkup.EditValue?.ToString();
            
            if (string.IsNullOrEmpty(sVal))
            {
                lkup.Properties.Appearance.BackColor = Color.OrangeRed;
            }
            else if (sVal.Equals("****"))
            {
                lkup.Properties.Appearance.BackColor = Color.OrangeRed;
            }
            else if (!string.IsNullOrEmpty(sVal))
            {
                lkup.Properties.Appearance.BackColor = Color.PaleGreen;
            }
        }

        /// <summary>
        ///        값에 따라 색변경
        /// </summary>
        private void TxtAText_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit text = (TextEdit)sender;
            string sVal = text.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
            {
                text.Properties.Appearance.BackColor = Color.OrangeRed;
            }
            else if (!string.IsNullOrEmpty(sVal))
            {
                text.Properties.Appearance.BackColor = Color.PaleGreen;
            }
        }

        private void AC16001F00_TextChanged(object sender, EventArgs e)
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
                layoutControl2.SaveLayoutToXml(path + @"\" + this.Name + "_Layout2.xaml");
                layoutControl3.SaveLayoutToXml(path + @"\" + this.Name + "_Layout3.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        /// <summary>
        ///        결제대상에 따라 재조회
        /// </summary>
        private void CboPmntTarget_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBoxEdit cbo = (ComboBoxEdit)sender;
            string sVal = cbo.EditValue?.ToString();
            if (string.IsNullOrEmpty(sVal))
            {
                XtraMessageBox.Show("결제대상을 정확히 선택하세요.");
                cbo.SelectAll();
                cbo.Focus();
                return;
            }

            BtnRetr.PerformClick();
        }

        private void GridViewRetr3_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridCol3OfsAmt)
            {
                double dVal = string.IsNullOrEmpty(e.CellValue?.ToString()) ? 0 : Convert.ToDouble(e.CellValue);
                if (dVal > 0)
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
                else
                {
                    e.Appearance.BackColor = SystemColors.Info;
                }
            }
            else if (e.Column == GridCol3Rmk)
            {
                string sVal = e.CellValue?.ToString().Trim();
                if (!string.IsNullOrEmpty(sVal))
                    e.Appearance.BackColor = Color.PaleGreen;
                else
                    e.Appearance.BackColor = SystemColors.Info;
            }
        }

        private void GridViewRetr3_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr3_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr3_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                double dVal = string.IsNullOrEmpty(GridViewRetr3.GetFocusedRowCellValue(GridCol3OfsAmt)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr3.GetFocusedRowCellValue(GridCol3OfsAmt));
                double dOfsBls = Convert.ToDouble(GridViewRetr3.GetFocusedRowCellValue(GridCol3OfsBlc));
                if (dVal == 0)
                {
                    GridViewRetr3.SetFocusedRowCellValue(GridCol3OfsAmt, GridViewRetr3.GetFocusedRowCellValue(GridCol3OfsBlc));
                }
                else if (dVal == dOfsBls)
                {
                    GridViewRetr3.SetFocusedRowCellValue(GridCol3OfsAmt, 0);
                }
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}